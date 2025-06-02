/**
 * @license
 * Copyright 2025 Daniel Joseph Sapit
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 *
 * @version 1.0.0
 * @file Code.js
 * @description Server-side Google Apps Script code for automating peer evaluation
 * form creation, distribution via Google Classroom, and response processing.
 * This script is intended to be used with its companion WebApp.html.
 */

// Global constant for the web app developer's error log spreadsheet
const WEBAPP_ERROR_LOG_SPREADSHEET_ID = 'YOUR_DEVELOPER_ERROR_LOG_SPREADSHEET_ID'; // <<< PASTE YOUR ERROR LOG SPREADSHEET ID HERE
const WEBAPP_ERROR_LOG_SHEET_NAME = 'Errors'; // Name of the sheet within the error log spreadsheet where errors will be recorded.
const XOR_KEY = 'YourSecretKeyForXOR'; // <<< IMPORTANT: CHANGE THIS TO A UNIQUE RANDOM STRING

// --- User Defaults Functions ---
function saveUserDefaults(formData) {
  try {
    PropertiesService.getUserProperties().setProperty('userDefaults', JSON.stringify(formData));
    return {
      status: "success",
      message: "Defaults saved successfully."
    };
  } catch (e) {
    Logger.log(`Error saving user defaults: ${e}`);
    return {
      status: "error",
      message: `Error saving defaults: ${e.message}`
    };
  }
}

function getUserDefaults() {
  try {
    const defaults = PropertiesService.getUserProperties().getProperty('userDefaults');
    return defaults ? JSON.parse(defaults) : {};
  } catch (e) {
    Logger.log(`Error retrieving user defaults: ${e}`);
    return {}; // Return empty object on error to avoid breaking UI
  }
}

// --- Trigger Management Functions ---
function getProjectTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const scriptProperties = PropertiesService.getScriptProperties();
  const webAppTriggers = [];

  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "onFormSubmit_WebApp") {
      const triggerId = trigger.getUniqueId();
      const contextString = scriptProperties.getProperty('trigger_' + triggerId);
      let context = {
        formTitle: "N/A (or older trigger)",
        sheet: "N/A",
        group: "N/A",
        course: "N/A",
        created: "N/A"
      };
      if (contextString) {
        try {
          context = JSON.parse(contextString);
          // Ensure essential fields have defaults
          context.formTitle = context.formTitle || "N/A (context missing title)";
          context.sheet = context.sheet || "N/A";
          context.group = context.group || "N/A";
          context.course = context.course || "N/A";
          context.created = context.created || "N/A";
        } catch (e) {
          Logger.log(`Failed to parse context for trigger ${triggerId}: ${e.message}`);
          context.formTitle = "Error parsing context";
          context.created = "Error parsing date context";
        }
      }
      webAppTriggers.push({
        id: triggerId,
        handlerFunction: trigger.getHandlerFunction(),
        eventType: trigger.getEventType().toString(),
        formTitle: context.formTitle,
        sheetName: context.sheet,
        groupName: context.group,
        courseName: context.course,
        creationDate: context.created
      });
    }
  });
  return webAppTriggers;
}

function deleteSelectedTriggers(triggerIds) {
  if (!triggerIds || triggerIds.length === 0) {
    return {
      status: "info",
      message: "No triggers selected for deletion."
    };
  }
  const scriptProperties = PropertiesService.getScriptProperties();
  let deletedCount = 0;
  let errorCount = 0;
  const projectTriggers = ScriptApp.getProjectTriggers(); // Get a fresh list

  triggerIds.forEach(id => {
    const triggerToDelete = projectTriggers.find(t => t.getUniqueId() === id);
    if (triggerToDelete) {
      try {
        ScriptApp.deleteTrigger(triggerToDelete);
        scriptProperties.deleteProperty('trigger_' + id);
        deletedCount++;
      } catch (e) {
        Logger.log(`Failed to delete trigger ${id}: ${e.message}`);
        errorCount++;
      }
    } else {
      Logger.log(`Trigger ID ${id} not found for deletion (might have been already deleted).`);
      // Also attempt to delete property in case it's orphaned
      scriptProperties.deleteProperty('trigger_' + id);
    }
  });

  let message = "";
  if (deletedCount > 0) message += `${deletedCount} trigger(s) deleted successfully. `;
  if (errorCount > 0) message += `${errorCount} trigger(s) failed to delete (see logs).`;
  if (deletedCount === 0 && errorCount === 0) message = "Selected triggers were not found or already deleted.";

  return {
    status: deletedCount > 0 ? "success" : "error",
    message: message
  };
}

function deleteAllWebAppTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const scriptProperties = PropertiesService.getScriptProperties();
  let deletedCount = 0;
  let errors = [];

  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "onFormSubmit_WebApp") {
      const triggerId = trigger.getUniqueId();
      try {
        ScriptApp.deleteTrigger(trigger);
        scriptProperties.deleteProperty('trigger_' + triggerId);
        deletedCount++;
      } catch (e) {
        Logger.log(`Error deleting trigger ${triggerId}: ${e.message}`);
        errors.push(e.message);
      }
    }
  });

  if (errors.length > 0) {
    return {
      status: "error",
      message: `Deleted ${deletedCount} triggers, but encountered errors: ${errors.join("; ")}`
    };
  }
  return {
    status: "success",
    message: `Successfully deleted ${deletedCount} 'onFormSubmit_WebApp' triggers.`
  };
}


// --- XOR Utility Functions ---
function xorString(input, key) {
  let output = '';
  for (let i = 0; i < input.length; i++) {
    output += String.fromCharCode(input.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }
  return output;
}

// --- Web App Setup ---
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('Form & Classroom Automation')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Server-side function for WebApp previews and checks ---
function getPreviewAndCellData(sheetName, namesRangeStr, groupNamesRangeStr, evalStartCol, urlCellA1, formFolder, courseName, courseTopicName, dueDateStr, dueTimeStr, utcOffsetStr, scheduleDateStr, scheduleTimeStr, includeJustification, includeFeedback) {
  let criticalErrorForProcessing = null;
  let errorMessages = [];

  const response = {
    names: [],
    groupNames: [],
    studentCount: 0,
    allNamesHaveGroup: true,
    namesMissingGroupCount: 0,
    uniqueGroupCount: 0,
    uniqueGroups: [],
    uniqueGroupsPreview: [],
    evalHeader: {
      isEmpty: true,
      value: ""
    },
    urlCell: {
      isEmpty: true,
      value: ""
    },
    evalTableAreaClear: true,
    evalTableCheckMessage: "Not checked.",
    driveFolderExists: false,
    driveFolderError: null,
    driveFolderCheckMessage: "Not checked.",
    classroomTopicExists: false,
    classroomTopicError: null,
    classroomTopicCheckMessage: "Not checked.",
    dueDateInFuture: null,
    dueDateCheckMessage: "Not checked.",
    classroomNameMismatches: [],
    classroomNameMismatchCount: 0,
    classroomNameMismatchesPreview: [],
    classroomRosterCheckMessage: "Not checked.",
    classroomRoster: [],
    scheduleDateCheckMessage: "Not checked.",
    error: null,
    errors: null
  };

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetName) {
      response.error = "Sheet Name is required.";
      return response;
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      response.error = `Sheet "${sheetName}" not found.`;
      return response;
    }

    // 1. Read Names and Group Names
    let allRawNames = [];
    let allRawGroupNames = [];
    try {
      if (!namesRangeStr) {
        throw new Error("Names Range is required.");
      }
      allRawNames = sheet.getRange(namesRangeStr).getValues();
      if (!groupNamesRangeStr) {
        throw new Error("Group Names Range is required.");
      }
      allRawGroupNames = sheet.getRange(groupNamesRangeStr).getValues();
      response.namesRangeValid = true;
    } catch (e) {
      const msg = `Error reading Names/Group Names Ranges: ${e.message}. Ensure ranges are valid A1 notation (e.g., A2:A17).`;
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + msg;
      errorMessages.push(msg);
      response.namesRangeValid = false;
    }

    const validStudentData = [];
    if (response.namesRangeValid && allRawNames.length > 0) {
      for (let i = 0; i < allRawNames.length; i++) {
        const name = allRawNames[i][0];
        if (name && name.toString().trim() !== "") {
          const group = (allRawGroupNames[i] && allRawGroupNames[i][0]) ? allRawGroupNames[i][0].toString().trim() : "";
          validStudentData.push({
            name: name.toString().trim(),
            group: group
          });
        }
      }
    }
    response.studentCount = validStudentData.length;
    const filteredStudentNamesFromSheet = validStudentData.map(s => s.name);

    if (response.namesRangeValid && response.studentCount === 0) {
      const noStudentMsg = "No valid student names found in the specified 'Names Range'. Ensure the range is correct and contains names.";
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + noStudentMsg;
      errorMessages.push(noStudentMsg);
    }

    if (response.studentCount > 0) {
      const MAX_PREVIEW_ITEMS = 5;
      if (response.studentCount > MAX_PREVIEW_ITEMS * 2) { // Show first 5, ellipsis, last 5
        response.names = validStudentData.slice(0, MAX_PREVIEW_ITEMS).map(s => s.name)
          .concat(["..."])
          .concat(validStudentData.slice(-MAX_PREVIEW_ITEMS).map(s => s.name));
        response.groupNames = validStudentData.slice(0, MAX_PREVIEW_ITEMS).map(s => s.group)
          .concat(["..."])
          .concat(validStudentData.slice(-MAX_PREVIEW_ITEMS).map(s => s.group));
      } else {
        response.names = validStudentData.map(s => s.name);
        response.groupNames = validStudentData.map(s => s.group);
      }
    }

    // 2. Check: All non-empty names have a group name
    if (response.studentCount > 0) {
      validStudentData.forEach(student => {
        if (!student.group || student.group.trim() === "") {
          response.namesMissingGroupCount++;
        }
      });
      response.allNamesHaveGroup = response.namesMissingGroupCount === 0;
      if (response.namesMissingGroupCount === response.studentCount && response.studentCount > 0) {
        const allMissingGroupsMsg = "Critical: All students found are missing group names. Group names are required for processing.";
        criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + allMissingGroupsMsg;
        errorMessages.push(allMissingGroupsMsg);
      } else if (response.namesMissingGroupCount > 0) {
        errorMessages.push(`Warning: ${response.namesMissingGroupCount} student(s) are missing group names and will be skipped during processing if not assigned a group.`);
      }
    }

    // 3. Unique Groups
    if (response.studentCount > 0) {
      response.uniqueGroups = [...new Set(validStudentData.map(s => s.group).filter(g => g && g.trim() !== ""))];
      response.uniqueGroupCount = response.uniqueGroups.length;
      const MAX_UNIQUE_GROUP_PREVIEW = 10;
      if (response.uniqueGroupCount > MAX_UNIQUE_GROUP_PREVIEW) {
        response.uniqueGroupsPreview = response.uniqueGroups.slice(0, MAX_UNIQUE_GROUP_PREVIEW).concat([`... and ${response.uniqueGroupCount - MAX_UNIQUE_GROUP_PREVIEW} more`]);
      } else {
        response.uniqueGroupsPreview = response.uniqueGroups;
      }
      if (response.uniqueGroupCount === 0 && response.studentCount > 0 && response.namesMissingGroupCount < response.studentCount) {
        const noValidGroupsMsg = "No valid (non-empty) group names found for students who are listed with groups.";
        criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + noValidGroupsMsg;
        errorMessages.push(noValidGroupsMsg);
      }
    }

    // 4. Eval Header & URL Cell
    var evalTableStartCol;
    var namesStartRow;
    var numColsToCheckInEvalTable;
    try {
      if (!evalStartCol) {
        throw new Error("Evaluation Table Starting Column is required." + `evalStartCol: ${evalStartCol}`);
      }
      if (!/^[A-Z]+$/.test(evalStartCol)) {
        throw new Error('Invalid column letter.');
      }
      evalTableStartCol = sheet.getRange(evalStartCol + '1').getColumn();

      const tempGetMembersByGroupName = {};
      validStudentData.forEach(s => {
        if (s.group && s.group.trim() !== "") {
          if (!tempGetMembersByGroupName[s.group]) tempGetMembersByGroupName[s.group] = [];
          tempGetMembersByGroupName[s.group].push(s.name);
        }
      });
      let tempMaxGroupSize = 0;
      Object.values(tempGetMembersByGroupName).forEach(members => {
        if (members.length > tempMaxGroupSize) tempMaxGroupSize = members.length;
      });
      if (tempMaxGroupSize === 0 && response.studentCount > 0) tempMaxGroupSize = 1; // For individuals not in groups but still processed

      // Calculate number of columns to check based on flags
      const numColsPerMemberInTable = 1 + (includeJustification ? 1 : 0);
      const numDataColsForScoresAndJustifications = (tempMaxGroupSize > 0 ? tempMaxGroupSize : 1) * numColsPerMemberInTable;
      numColsToCheckInEvalTable = numDataColsForScoresAndJustifications + (includeFeedback ? 1 : 0) + 2; // + Avg + Last Submitted

      namesStartRow = sheet.getRange(namesRangeStr).getRow();
      if (namesStartRow > 1) {
        const checkHeaderRange = sheet.getRange(namesStartRow - 1, evalTableStartCol, 1, numColsToCheckInEvalTable);

        response.evalHeader.value = `Not empty: ${checkHeaderRange.getDisplayValues()}`;
        response.evalHeader.isEmpty = checkHeaderRange.isBlank();
      } else {
        response.evalHeader.value = "Warning: No space for Eval Table Headers.";
        response.evalHeader.isEmpty = false;
      }

    } catch (e) {
      const msg = `Eval Start Column (${evalStartCol || 'not specified'}): ${e.message}`;
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + msg;
      errorMessages.push(msg);
      response.evalHeader.value = "Error";
    }
    try {
      if (!urlCellA1) {
        throw new Error("Cell to Write Form URLs is required.");
      }
      const urlCell = sheet.getRange(urlCellA1);
      response.urlCell.value = urlCell.getDisplayValue();
      response.urlCell.isEmpty = urlCell.isBlank();
    } catch (e) {
      const msg = `URL Links Cell (${urlCellA1 || 'not specified'}): ${e.message}`;
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + msg;
      errorMessages.push(msg);
      response.urlCell.value = "Error";
    }

    // 5. Evaluation Table Area Check
    if (response.studentCount > 0 && evalStartCol && response.namesRangeValid) {
      try {
        if (!evalTableStartCol) {
          throw new Error('Evaluation Table Starting Column is required.');
        }
        let firstProblemCellNotation = "";
        let problemStudentName = "";

        for (let i = 0; i < allRawNames.length; i++) {
          const currentName = allRawNames[i][0];
          if (currentName && currentName.toString().trim() !== "") {
            const studentRowInSheet = namesStartRow + i;
            // Ensure we only check rows for students present in validStudentData
            if (validStudentData.find(s => s.name === currentName.toString().trim())) {
              const checkRange = sheet.getRange(studentRowInSheet, evalTableStartCol, 1, numColsToCheckInEvalTable);
              if (!checkRange.isBlank()) {
                response.evalTableAreaClear = false;
                const cellValues = checkRange.getDisplayValues()[0];
                for (let j = 0; j < cellValues.length; j++) {
                  if (cellValues[j] !== "") {
                    firstProblemCellNotation = sheet.getRange(studentRowInSheet, evalTableStartCol + j).getA1Notation();
                    break;
                  }
                }
                problemStudentName = currentName.toString().trim();
                break;
              }
            }
          }
        }
        response.evalTableCheckMessage = response.evalTableAreaClear ? "Evaluation table area for students appears clear." : `Evaluation table area for students is NOT clear. First non-empty cell found at ${firstProblemCellNotation} (row of student '${problemStudentName}').`;
      } catch (e) {
        response.evalTableCheckMessage = `Error: ${e.message}`;
        response.evalTableAreaClear = false;
        errorMessages.push(response.evalTableCheckMessage);
      }
    } else if (response.studentCount === 0 && response.namesRangeValid) {
      response.evalTableCheckMessage = "No students to check.";
    } else {
      response.evalTableCheckMessage = "Skipped (student data or eval starting column invalid).";
    }

    // 6. Drive Folder Check
    try {
      if (!formFolder) {
        throw new Error("Drive Folder Name is required.");
      }
      response.driveFolderExists = DriveApp.getFoldersByName(formFolder).hasNext();
      response.driveFolderCheckMessage = response.driveFolderExists ? "Folder exists." : "Folder does NOT exist (will be created).";
    } catch (e) {
      response.driveFolderCheckMessage = `Error: ${e.message}`;
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + `Drive Folder: ${e.message}`;
      errorMessages.push(response.driveFolderCheckMessage);
    }

    // 7. Classroom Topic Check & Roster Fetch
    let courseIdForChecks;
    try {
      if (!courseName) {
        throw new Error("Classroom Course Name is required.");
      }
      courseIdForChecks = getCourseIdByName(courseName);
      if (!courseTopicName) {
        throw new Error("Classroom Topic Name is required.");
      }
      const topicsResponse = Classroom.Courses.Topics.list(courseIdForChecks);
      response.classroomTopicExists = topicsResponse.topic && topicsResponse.topic.find(t => t.name === courseTopicName);
      response.classroomTopicCheckMessage = response.classroomTopicExists ? "Topic exists." : "Topic does NOT exist (will be created).";
      const classroomStudentsResponse = Classroom.Courses.Students.list(courseIdForChecks);
      const classroomStudents = classroomStudentsResponse.students;
      if (classroomStudents && classroomStudents.length > 0) {
        response.classroomRoster = classroomStudents
          .map(s => s.profile.name.fullName)
          .sort((a, b) => {
            // Extract last names
            const lastNameA = a.trim().split(' ').slice(-1)[0].toLowerCase();
            const lastNameB = b.trim().split(' ').slice(-1)[0].toLowerCase();

            if (lastNameA < lastNameB) return -1;
            if (lastNameA > lastNameB) return 1;

            // If last names are the same, sort by full name as tie-breaker
            return a.toLowerCase() < b.toLowerCase() ? -1 : (a.toLowerCase() > b.toLowerCase() ? 1 : 0);
          });
      } else {
        response.classroomRoster = ["No students found in Classroom course."];
      }
    } catch (e) {
      response.classroomTopicCheckMessage = `Error related to topic/course: ${e.message}`;
      response.classroomRosterCheckMessage = `Could not fetch roster: ${e.message}`;
      response.classroomRoster = [`Error fetching roster: ${e.message}`];
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + `Classroom Course/Topic/Roster: ${e.message}`;
      errorMessages.push(response.classroomTopicCheckMessage);
    }

    // 8. Due Date Check
    try {
      if (!dueDateStr || !dueTimeStr || !utcOffsetStr) {
        throw new Error("Due Date, Time, and UTC Offset are all required.");
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDateStr) || !/^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$/.test(dueTimeStr) || !/^[+\-]\d{4}$/.test(utcOffsetStr)) {
        throw new Error("Invalid due date, time, or UTC offset format.");
      }
      const dueDateTime = new Date(`${dueDateStr}T${dueTimeStr}${utcOffsetStr}`);
      if (isNaN(dueDateTime.getTime())) {
        throw new Error("Could not parse due date/time.");
      }
      response.dueDateInFuture = dueDateTime > new Date();
      response.dueDateCheckMessage = response.dueDateInFuture ? "Due date is in the future." : "Due date is NOT in the future.";
      if (!response.dueDateInFuture) {
        criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + "Due date must be in the future.";
        errorMessages.push("Due date must be in the future.");
      }
    } catch (e) {
      response.dueDateCheckMessage = `Error: ${e.message}`;
      response.dueDateInFuture = false;
      criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + `Due Date: ${e.message}`;
      errorMessages.push(response.dueDateCheckMessage);
    }

    // 9. Classroom Name Mismatch Check (uses roster fetched in step 7 if successful)
    if (courseIdForChecks && response.studentCount > 0 && response.classroomRoster.length > 0 && !response.classroomRoster[0].startsWith("Error") && !response.classroomRoster[0].startsWith("No students")) {
      const classroomFullNamesSet = new Set(response.classroomRoster);
      filteredStudentNamesFromSheet.forEach(sheetStdName => {
        if (!classroomFullNamesSet.has(sheetStdName)) {
          response.classroomNameMismatches.push(sheetStdName);
        }
      });
      response.classroomNameMismatchCount = response.classroomNameMismatches.length;
      const MAX_MISMATCH_PREVIEW = 5;
      response.classroomNameMismatchesPreview = response.classroomNameMismatchCount > MAX_MISMATCH_PREVIEW * 2 ? response.classroomNameMismatches.slice(0, MAX_MISMATCH_PREVIEW).concat(["..."]).concat(response.classroomNameMismatches.slice(-MAX_MISMATCH_PREVIEW)) : response.classroomNameMismatches;
      if (response.classroomNameMismatchCount > 0) {
        response.classroomRosterCheckMessage = `${response.classroomNameMismatchCount} name(s) from sheet NOT found in Classroom roster.`;
        criticalErrorForProcessing = (criticalErrorForProcessing ? criticalErrorForProcessing + "; " : "") + response.classroomRosterCheckMessage + " This is a critical error.";
        errorMessages.push(response.classroomRosterCheckMessage + " Assignments will not be correctly made for these students.");
      } else {
        response.classroomRosterCheckMessage = "All names from sheet appear to match Classroom roster.";
      }
    } else if (!courseIdForChecks && response.studentCount > 0) {
      response.classroomRosterCheckMessage = "Skipped (Classroom course not found/specified or error fetching roster).";
    } else if (response.studentCount === 0 && response.namesRangeValid) {
      response.classroomRosterCheckMessage = "No students on sheet to check.";
    } else if (response.classroomRoster.length > 0 && (response.classroomRoster[0].startsWith("Error") || response.classroomRoster[0].startsWith("No students"))) {
      response.classroomRosterCheckMessage = "Skipped (Problem with Classroom roster for comparison).";
    }

    // 10. Schedule Date Future Check (Warning only)
    try {
      if (scheduleDateStr && scheduleTimeStr) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(scheduleDateStr) || !/^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$/.test(scheduleTimeStr) || !utcOffsetStr || !/^[+\-]\d{4}$/.test(utcOffsetStr)) {
          response.scheduleDateCheckMessage = "Invalid schedule date, time, or missing UTC offset for this check.";
          errorMessages.push(response.scheduleDateCheckMessage);
        } else {
          const scheduleDateTime = new Date(`${scheduleDateStr}T${scheduleTimeStr}${utcOffsetStr}`);
          if (isNaN(scheduleDateTime.getTime())) {
            response.scheduleDateCheckMessage = "Could not parse schedule date/time.";
            errorMessages.push(response.scheduleDateCheckMessage);
          } else {
            response.scheduleDateCheckMessage = scheduleDateTime <= new Date() ? "Warning: Schedule date/time is in the past or is now. Assignment would publish immediately if not far enough in future for Classroom API." : "Schedule date/time is in the future.";
            if (scheduleDateTime <= new Date()) errorMessages.push(response.scheduleDateCheckMessage);
          }
        }
      } else if (scheduleDateStr || scheduleTimeStr) {
        response.scheduleDateCheckMessage = "Provide both Schedule Date and Time for check.";
      } else {
        response.scheduleDateCheckMessage = "Not scheduled.";
      }
    } catch (e) {
      response.scheduleDateCheckMessage = `Error checking schedule date: ${e.message}`;
      errorMessages.push(response.scheduleDateCheckMessage);
    }

    if (criticalErrorForProcessing) {
      response.error = criticalErrorForProcessing;
    }
    if (errorMessages.length > 0) {
      response.errors = errorMessages.join("; ");
    }
    return response;
  } catch (e) {
    Logger.log(`CRITICAL Error in getPreviewAndCellData: ${e.toString()}\nStack: ${e.stack}`);
    response.error = `Server error in getPreviewAndCellData: ${e.message}. ` + (criticalErrorForProcessing || "");
    response.errors = (response.errors ? response.errors + "; " : "") + errorMessages.join("; ");
    response.classroomRoster = ["Error occurred before roster could be fetched."];
    logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : "N/A", "getPreviewAndCellData Crash", e.toString(), e.stack);
    return response;
  }
}

// --- Main Processing Function (called from WebApp.html) ---
function processUserInputs(formObject, deleteAllOldTriggersFlag) {
  try {
    if (deleteAllOldTriggersFlag) {
      const deletionResult = deleteAllWebAppTriggers();
      Logger.log(`Trigger deletion before processing: ${deletionResult.message}`);
      if (deletionResult.status === "error" && deletionResult.message.includes("encountered errors")) {
        // Log this to developer sheet as a warning, but proceed if possible
        logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), SpreadsheetApp.getActiveSpreadsheet().getId(), "Processing - Trigger Deletion Warning", deletionResult.message);
      }
    }

    // Convert checkbox values from string "true"/"false" or boolean to actual booleans
    const includeJustificationInTable = formObject.includeJustificationInTable === true || formObject.includeJustificationInTable === 'true';
    const includeFeedbackInTable = formObject.includeFeedbackInTable === true || formObject.includeFeedbackInTable === 'true';
    const allowResponseEdits = formObject.allowResponseEdits === true || formObject.allowResponseEdits === 'true';

    const validationResult = getPreviewAndCellData(
      formObject.sheetName, formObject.namesRange, formObject.groupNamesRange,
      formObject.evaluationTableStartCol, formObject.urlLinksRange, formObject.formFolder,
      formObject.courseName, formObject.courseTopicName, formObject.dueDate,
      formObject.dueTime, formObject.utcOffset, formObject.assignmentScheduleDate,
      formObject.assignmentScheduleTime,
      includeJustificationInTable,
      includeFeedbackInTable
    );
    if (validationResult.error) {
      throw new Error(`Input validation failed: ${validationResult.error}`);
    }

    formObject.scaleLowerBound = parseInt(formObject.scaleLowerBound, 10);
    formObject.scaleUpperBound = parseInt(formObject.scaleUpperBound, 10);
    formObject.assignmentPoints = parseInt(formObject.assignmentPoints, 10);
    formObject.allowResponseEdits = allowResponseEdits;
    formObject.includeJustificationInTable = includeJustificationInTable;
    formObject.includeFeedbackInTable = includeFeedbackInTable;


    const userSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = userSpreadsheet.getSheetByName(formObject.sheetName);
    const courseId = getCourseIdByName(formObject.courseName);
    const classroomStudentsResponse = Classroom.Courses.Students.list(courseId);
    if (!classroomStudentsResponse.students || classroomStudentsResponse.students.length === 0) {
      throw new Error("No students found in the specified Google Classroom course.");
    }
    const getUserIdByNameFromClassroom = {};
    const getFullNameByEmailFromClassroom = {};
    classroomStudentsResponse.students.forEach(student => {
      getUserIdByNameFromClassroom[student.profile.name.fullName] = student.userId;
      getFullNameByEmailFromClassroom[student.profile.emailAddress] = student.profile.name.fullName;
    });

    const rawNamesFromSheet = userSheet.getRange(formObject.namesRange).getValues();
    const rawGroupNamesFromSheet = userSheet.getRange(formObject.groupNamesRange).getValues();
    const namesForProcessing = [];
    const groupNamesForProcessing = [];
    let studentsSkippedMsg = "";
    for (let i = 0; i < rawNamesFromSheet.length; i++) {
      const name = rawNamesFromSheet[i][0];
      if (name && name.toString().trim() !== "") {
        const groupName = (rawGroupNamesFromSheet[i] && rawGroupNamesFromSheet[i][0]) ? rawGroupNamesFromSheet[i][0].toString().trim() : "";
        if (groupName === "") {
          studentsSkippedMsg += `${name.toString().trim()} (missing group); `;
          continue;
        }
        if (!getUserIdByNameFromClassroom[name.toString().trim()]) {
          studentsSkippedMsg += `${name.toString().trim()} (not in Classroom roster or name mismatch); `;
          logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), userSpreadsheet.getId(), "Processing - Name Mismatch Skip", `Sheet name "${name.toString().trim()}" not found in Classroom roster for course "${formObject.courseName}". Student skipped entirely.`);
          continue;
        }
        namesForProcessing.push(name.toString().trim());
        groupNamesForProcessing.push(groupName);
      }
    }
    if (studentsSkippedMsg) {
      logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), userSpreadsheet.getId(), "Student Processing Info", "Students skipped: " + studentsSkippedMsg.slice(0, -2));
    }
    if (namesForProcessing.length === 0) {
      throw new Error("No students remaining for processing after initial filters (missing group or Classroom name mismatch).");
    }


    const userSpecificMappings = buildUserSpecificMappings(namesForProcessing, groupNamesForProcessing, userSheet.getRange(formObject.namesRange).getRow(), userSheet.getRange(formObject.evaluationTableStartCol + '1').getColumn(), rawNamesFromSheet);
    userSpecificMappings.getUserIdByName = {};
    namesForProcessing.forEach(name => {
      userSpecificMappings.getUserIdByName[name] = getUserIdByNameFromClassroom[name];
    });
    userSpecificMappings.getFullNameByEmail = getFullNameByEmailFromClassroom;

    const topicId = checkUserInputsAndGetTopicId(formObject, courseId);
    const driveFolder = getOrCreateDriveFolder(formObject.formFolder);

    initUserPeerTable(userSheet, userSpecificMappings, formObject);

    const formCreationConfig = {
      sheetName: formObject.sheetName,
      namesRange: formObject.namesRange,
      groupNamesRange: formObject.groupNamesRange,
      evaluationTableStartCol: formObject.evaluationTableStartCol,
      originalFormTitle: formObject.formTitle,
      courseName: formObject.courseName,
      utcOffset: formObject.utcOffset,
      includeJustificationInTable: formObject.includeJustificationInTable,
      includeFeedbackInTable: formObject.includeFeedbackInTable
    };
    const formURLs = createUserForms(formObject, userSpecificMappings, driveFolder, topicId, formCreationConfig, userSpreadsheet.getId());
    userSheet.getRange(formObject.urlLinksRange).setValue(Object.keys(formURLs).length > 0 ? Object.values(formURLs).join('\n') : "No forms were created (e.g., no valid groups after filters).");
    createUserClassroomAssignments(formURLs, formObject, userSpecificMappings, courseId, topicId);
    let successMessage = "Forms and assignments processed successfully! URLs written to " + formObject.urlLinksRange + ". Don't forget to set \"Collect email addresses\" to \"Verified\" under \"Responses\" of each form's settings.";
    if (studentsSkippedMsg) {
      successMessage += " Some students were skipped (see script logs/developer log for details).";
    }
    return {
      status: "success",
      message: successMessage
    };

  } catch (error) {
    Logger.log(`processUserInputs Error: ${error.toString()}\nStack: ${error.stack}`);
    logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : "N/A", "processUserInputs Error", error.toString(), error.stack);
    return {
      status: "error",
      message: "An error occurred: " + error.message
    };
  }
}

function getCourseIdByName(courseName) {
  if (!courseName || courseName.trim() === "") {
    throw new Error("Course name cannot be empty.");
  }
  const coursesResponse = Classroom.Courses.list({
    teacherId: 'me',
    courseStates: ['ACTIVE']
  });
  if (!coursesResponse.courses || coursesResponse.courses.length === 0) {
    throw new Error("No active courses found in Google Classroom.");
  }
  const course = coursesResponse.courses.find(c => c.name === courseName);
  if (course) {
    return course.id;
  }
  const foundCaseInsensitive = coursesResponse.courses.find(c => c.name.toLowerCase() === courseName.toLowerCase());
  if (foundCaseInsensitive) {
    Logger.log(`Course "${courseName}" not found exactly, using case-insensitive "${foundCaseInsensitive.name}".`);
    return foundCaseInsensitive.id;
  }
  throw new Error(`Course "${courseName}" not found. Check name and ensure it's active.`);
}

function buildUserSpecificMappings(filteredNames, filteredGroupNames, namesActualStartRow, evalTableStartCol, rawAllNames) {
  const getMembersByGroupName = {};
  filteredNames.forEach((name, index) => {
    const groupName = filteredGroupNames[index];
    if (groupName && groupName.trim() !== "") {
      const trimmedGroupName = groupName.trim();
      if (!getMembersByGroupName[trimmedGroupName]) {
        getMembersByGroupName[trimmedGroupName] = [];
      }
      getMembersByGroupName[trimmedGroupName].push(name);
    }
  });
  const getIndexInGroupByName = {};
  Object.keys(getMembersByGroupName).forEach(groupName => {
    getMembersByGroupName[groupName].forEach((member, index) => {
      getIndexInGroupByName[member] = index;
    });
  });
  const getRowByName = {};
  rawAllNames.forEach((row, i) => {
    const rawNameCell = row[0];
    if (rawNameCell && rawNameCell.toString().trim() !== "") {
      const currentRawNameTrimmed = rawNameCell.toString().trim();
      if (filteredNames.includes(currentRawNameTrimmed)) {
        getRowByName[currentRawNameTrimmed] = namesActualStartRow + i;
      }
    }
  });
  let maxGroupSize = 0;
  Object.values(getMembersByGroupName).forEach(members => {
    if (members.length > maxGroupSize) maxGroupSize = members.length;
  });
  if (maxGroupSize === 0 && filteredNames.length > 0 && Object.keys(getMembersByGroupName).length === 0) {
    maxGroupSize = 1;
  }
  return {
    getMembersByGroupName,
    getIndexInGroupByName,
    getRowByName,
    maxGroupSize,
    evalTableStartColumn: evalTableStartCol,
    processedNames: filteredNames,
    processedGroupNames: filteredGroupNames
  };
}

function checkUserInputsAndGetTopicId(formObject, courseId) {
  if (!Number.isInteger(formObject.scaleLowerBound) || !Number.isInteger(formObject.scaleUpperBound) || formObject.scaleUpperBound <= formObject.scaleLowerBound) {
    throw new Error("Invalid scale bounds: Must be integers and upper bound must be greater than lower bound.");
  }
  if (formObject.assignmentPoints < 0) {
    throw new Error("Assignment points cannot be negative.");
  }
  const topicsResponse = Classroom.Courses.Topics.list(courseId);
  if (topicsResponse.topic && topicsResponse.topic.length > 0) {
    const topic = topicsResponse.topic.find(t => t.name === formObject.courseTopicName);
    if (topic) return topic.topicId;
  }
  try {
    const newTopic = Classroom.Courses.Topics.create({
      name: formObject.courseTopicName
    }, courseId);
    Logger.log("Created new topic: " + newTopic.name + " with ID: " + newTopic.topicId);
    return newTopic.topicId;
  } catch (e) {
    throw new Error(`Topic "${formObject.courseTopicName}" not found and could not be created: ${e.message}`);
  }
}

function getOrCreateDriveFolder(folderName) {
  if (!folderName || folderName.trim() === "") {
    throw new Error("Drive Folder Name cannot be empty.");
  }
  const trimmedFolderName = folderName.trim();
  const folders = DriveApp.getFoldersByName(trimmedFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  try {
    return DriveApp.createFolder(trimmedFolderName);
  } catch (e) {
    throw new Error(`Could not find or create folder "${trimmedFolderName}": ${e.message}`);
  }
}

function initUserPeerTable(userSheet, userMappings, formObject) {
  const {
    getRowByName,
    evalTableStartColumn,
    maxGroupSize,
    processedNames
  } = userMappings;

  // Retrieve flags from formObject
  const includeJustification = formObject.includeJustificationInTable;
  const includeFeedback = formObject.includeFeedbackInTable;

  if (processedNames.length === 0) {
    Logger.log("No names for initializing peer table.");
    return;
  }

  const allSheetRowsForNames = processedNames.map(name => getRowByName[name]).filter(row => row > 0);
  if (allSheetRowsForNames.length === 0) {
    Logger.log("No valid sheet rows for table initialization.");
    return;
  }

  const topDataRow = Math.min(...allSheetRowsForNames);
  const bottomDataRow = Math.max(...allSheetRowsForNames);
  const numDataRowsInvolved = bottomDataRow - topDataRow + 1;

  const numColsPerMemberPair = 1 + (includeJustification ? 1 : 0); // Score + Optional Justification
  const actualMaxGroupSize = maxGroupSize > 0 ? maxGroupSize : 1; // Ensure at least 1 for calculations if no groups
  const numColsForScoresAndJustifications = actualMaxGroupSize * numColsPerMemberPair;

  // Total columns for table content: Scores/Justifications + Avg Score + Optional Feedback + Last Submitted
  const numColsForTableContent = numColsForScoresAndJustifications + 1 + (includeFeedback ? 1 : 0) + 1;

  if (evalTableStartColumn <= 0) {
    Logger.log(`Invalid evalTableStartColumn: ${evalTableStartColumn}.`);
    return;
  }
  if (numDataRowsInvolved > 0 && topDataRow > 0) {
    userSheet.getRange(topDataRow, evalTableStartColumn, numDataRowsInvolved, numColsForTableContent).clearContent().clearFormat();
  }

  const headerRowForEvalTable = userSheet.getRange(formObject.namesRange).getRow() - 1;
  if (headerRowForEvalTable >= 1) {
    const headers = [];
    for (let i = 1; i <= actualMaxGroupSize; i++) {
      headers.push(`M${i} Score`);
      if (includeJustification) {
        headers.push(`M${i} Just.`);
      }
    }
    headers.push("Avg Score");
    if (includeFeedback) {
      headers.push("Feedback");
    }
    headers.push("Last Submitted");

    try {
      userSheet.getRange(headerRowForEvalTable, evalTableStartColumn, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    } catch (e) {
      Logger.log(`Error setting headers: ${e.message}`);
    }
  } else {
    Logger.log(`Invalid header row for eval table: ${headerRowForEvalTable}.`);
  }

  processedNames.forEach(fullName => {
    const studentSheetRow = getRowByName[fullName];
    if (studentSheetRow > 0) {
      const scoreCellA1Notations = [];
      for (let i = 0; i < actualMaxGroupSize; i++) {
        // Score column is at startCol + i * (1 for score + 1 if justification exists)
        const scoreCol = evalTableStartColumn + (i * numColsPerMemberPair);
        scoreCellA1Notations.push(userSheet.getRange(studentSheetRow, scoreCol).getA1Notation());
      }

      const formula = scoreCellA1Notations.length > 0 ? `=IFERROR(AVERAGE(${scoreCellA1Notations.join(",")}), "")` : "";
      const avgScoreColumn = evalTableStartColumn + numColsForScoresAndJustifications; // Column immediately after all scores/justifications

      if (formula) {
        try {
          userSheet.getRange(studentSheetRow, avgScoreColumn).setFormula(formula);
        } catch (e) {
          Logger.log(`Error setting formula for ${fullName} at row ${studentSheetRow}, col ${avgScoreColumn}: ${e.message}`);
        }
      }
    }
  });
  Logger.log("Peer evaluation table structure updated.");
}

function createUserForms(formObject, userMappings, driveFolder, topicId, formCreationConfig, logSpreadsheetId) {
  const formURLs = {};
  const {
    getMembersByGroupName
  } = userMappings;
  const {
    sheetName,
    namesRange,
    groupNamesRange,
    evaluationTableStartCol,
    originalFormTitle,
    courseName,
    utcOffset,
    includeJustificationInTable,
    includeFeedbackInTable
  } = formCreationConfig;
  const scriptProperties = PropertiesService.getScriptProperties();

  if (Object.keys(getMembersByGroupName).length === 0) {
    Logger.log("No groups to create forms for.");
    return formURLs;
  }

  Object.keys(getMembersByGroupName).forEach(groupName => {
    const groupMembers = getMembersByGroupName[groupName];
    if (!groupMembers || groupMembers.length === 0) {
      Logger.log(`Skipping form for group "${groupName}" (no members).`);
      return;
    }

    const individualFormTitle = `${originalFormTitle} - Group ${groupName}`;
    const form = FormApp.create(individualFormTitle);
    form.setCollectEmail(true)
      .setAllowResponseEdits(formObject.allowResponseEdits)
      .setLimitOneResponsePerUser(true);

    const contextData = {
      sheet: sheetName,
      namesR: namesRange,
      groupsR: groupNamesRange,
      evalR: evaluationTableStartCol,
      course: courseName,
      group: groupName,
      utcOff: utcOffset,
      includeJustification: includeJustificationInTable,
      includeFeedback: includeFeedbackInTable
    };
    const obfuscatedContext = Utilities.base64Encode(xorString(JSON.stringify(contextData), XOR_KEY), Utilities.Charset.UTF_8);
    form.setDescription(`Form Data (do not edit): ${obfuscatedContext}\n\nThis form is for Group ${groupName} only.\nGroup Members: ${groupMembers.join(', ')}\n\n${formObject.formDescription}`);

    groupMembers.forEach(member => {
      form.addScaleItem().setTitle(`Rate ${member}'s performance`).setBounds(formObject.scaleLowerBound, formObject.scaleUpperBound).setLabels(formObject.scaleLowerLabel, formObject.scaleUpperLabel).setRequired(true);
      form.addTextItem().setTitle(`Justify why you gave ${member} that score. Give details about their contribution.`).setRequired(true);
    });
    form.addTextItem().setTitle("Any other concerns or feedback regarding the activity?").setRequired(false);

    // Delete existing triggers for this specific form with the handler "onFormSubmit_WebApp"
    const existingTriggersForForm = ScriptApp.getUserTriggers(form);
    existingTriggersForForm.forEach(tr => {
      if (tr.getHandlerFunction() === "onFormSubmit_WebApp") {
        const oldTriggerId = tr.getUniqueId();
        ScriptApp.deleteTrigger(tr);
        scriptProperties.deleteProperty('trigger_' + oldTriggerId); // Also remove its stored context
        Logger.log(`Deleted existing trigger ${oldTriggerId} for form "${form.getTitle()}" before creating new one.`);
      }
    });

    const trigger = ScriptApp.newTrigger("onFormSubmit_WebApp").forForm(form).onFormSubmit().create();
    const triggerId = trigger.getUniqueId();

    const now = new Date();
    const userUtcOffset = formCreationConfig.utcOffset;
    let formattedTimestampForStorage = "N/A";

    if (userUtcOffset && /^[+\-]\d{4}$/.test(userUtcOffset)) {
      const gmtOffset = "GMT" + userUtcOffset.substring(0, 3) + ":" + userUtcOffset.substring(3, 5);
      try {
        const dateTimePart = Utilities.formatDate(now, gmtOffset, "yyyy/MM/dd hh:mm a");
        formattedTimestampForStorage = `${dateTimePart} (UTC${userUtcOffset})`;
      } catch (e) {
        Logger.log(`Error formatting date for trigger context using offset ${userUtcOffset} (converted to ${gmtOffset}): ${e.message}. Storing ISO as fallback.`);
        formattedTimestampForStorage = now.toISOString() + ` (Error formatting with ${userUtcOffset})`;
      }
    } else {
      Logger.log(`Invalid or missing userUtcOffset ('${userUtcOffset}') for trigger creation. Storing as ISO UTC string.`);
      const utcDateTimePart = Utilities.formatDate(now, "GMT", "yyyy/MM/dd hh:mm a");
      formattedTimestampForStorage = `${utcDateTimePart} (UTC)`;
    }

    const triggerContext = {
      formTitle: individualFormTitle,
      sheet: sheetName,
      group: groupName,
      course: courseName,
      created: formattedTimestampForStorage,
      handlerFunction: "onFormSubmit_WebApp"
    };
    scriptProperties.setProperty('trigger_' + triggerId, JSON.stringify(triggerContext));

    formURLs[groupName] = form.getPublishedUrl();
    try {
      DriveApp.getFileById(form.getId()).moveTo(driveFolder);
    } catch (driveError) {
      Logger.log(`Error moving form ${form.getId()} to folder ${driveFolder.getName()}: ${driveError.message}`);
      logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), logSpreadsheetId, `Form Move Error for Group ${groupName}`, driveError.toString(), driveError.stack);
    }
    Logger.log(`Group ${groupName} form created: ${formURLs[groupName]}. Trigger ID: ${triggerId}`);
  });
  return formURLs;
}

function onFormSubmit_WebApp(e) {
  const formId = e.source.getId();
  const form = FormApp.openById(formId);
  const respondentEmail = e.response.getRespondentEmail();
  const itemResponses = e.response.getItemResponses();
  const formDescription = form.getDescription();
  const descriptionLines = formDescription.split('\n');
  const submissionTimestamp = new Date();
  let activeSpreadsheet, activeSpreadsheetId;

  try {
    activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    activeSpreadsheetId = activeSpreadsheet.getId();
  } catch (ssError) {
    logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", "Failed to get active spreadsheet.", ssError.message, "N/A");
    return;
  }

  try {
    const obfuscatedContextLine = descriptionLines.find(line => line.startsWith("Form Data (do not edit):"));
    if (!obfuscatedContextLine) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", "Obfuscated context line not found.", `Form: ${form.getTitle()}`, activeSpreadsheetId);
      return;
    }
    const obfuscatedContext = obfuscatedContextLine.substring("Form Data (do not edit): ".length).trim();
    let contextDataJSON;
    try {
      contextDataJSON = JSON.parse(xorString(Utilities.newBlob(Utilities.base64Decode(obfuscatedContext, Utilities.Charset.UTF_8)).getDataAsString(), XOR_KEY));
    } catch (decodeJsonError) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", "Failed to decode/parse context.", decodeJsonError.message, activeSpreadsheetId);
      return;
    }

    const {
      sheet: targetSheetName,
      namesR: namesRangeStr,
      groupsR: groupNamesRangeStr,
      evalR: evalTableStartColStr,
      course: courseNameForSubmitterLookup,
      group: formDesignatedGroup,
      utcOff: userUtcOffset,
      includeJustification,
      includeFeedback
    } = contextDataJSON;

    if (!targetSheetName || !namesRangeStr || !groupNamesRangeStr || !evalTableStartColStr || !courseNameForSubmitterLookup || !formDesignatedGroup || !userUtcOffset) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", "Critical fields missing from context.", JSON.stringify(contextDataJSON), activeSpreadsheetId);
      return;
    }

    const targetSheet = activeSpreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", `Target sheet "${targetSheetName}" not found.`, null, activeSpreadsheetId);
      return;
    }

    let submitterFullName = null;
    try {
      const courseId = getCourseIdByName(courseNameForSubmitterLookup);
      const student = Classroom.Courses.Students.get(courseId, respondentEmail);
      if (student && student.profile && student.profile.name) {
        submitterFullName = student.profile.name.fullName;
      }
    } catch (err) {
      Logger.log(`Could not map email ${respondentEmail} for course "${courseNameForSubmitterLookup}" directly. Will try list. Error: ${err.message}`);
    }

    if (!submitterFullName) {
      try {
        const courseIdForSubmitterLookup = getCourseIdByName(courseNameForSubmitterLookup);
        const classroomStudents = Classroom.Courses.Students.list(courseIdForSubmitterLookup).students;
        if (classroomStudents) {
          const studentProfile = classroomStudents.find(s => s.profile.emailAddress === respondentEmail);
          if (studentProfile) submitterFullName = studentProfile.profile.name.fullName;
        }
      } catch (err) {
        logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Warning", `Could not map email ${respondentEmail} for course "${courseNameForSubmitterLookup}" via list. Error: ${err.message}`, `Form Title: ${form.getTitle()}`, activeSpreadsheetId);
      }
    }
    if (!submitterFullName) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", `Respondent email ${respondentEmail} could not be mapped to Classroom full name.`, `Form: ${form.getTitle()}`, activeSpreadsheetId);
      return;
    }


    const rawNamesFromTargetSheet = targetSheet.getRange(namesRangeStr).getValues();
    const rawGroupNamesFromTargetSheet = targetSheet.getRange(groupNamesRangeStr).getValues();
    const namesForContext = [];
    const groupNamesForContext = [];
    rawNamesFromTargetSheet.forEach((row, i) => {
      if (row[0] && row[0].toString().trim() !== "") {
        namesForContext.push(row[0].toString().trim());
        groupNamesForContext.push((rawGroupNamesFromTargetSheet[i] && rawGroupNamesFromTargetSheet[i][0]) ? rawGroupNamesFromTargetSheet[i][0].toString().trim() : "");
      }
    });
    const validNamesForContext = [];
    const validGroupNamesForContext = [];
    namesForContext.forEach((name, idx) => {
      if (groupNamesForContext[idx] !== "") {
        validNamesForContext.push(name);
        validGroupNamesForContext.push(groupNamesForContext[idx]);
      }
    });
    const submissionContext = buildUserSpecificMappings(validNamesForContext, validGroupNamesForContext, targetSheet.getRange(namesRangeStr).getRow(), targetSheet.getRange(evalTableStartColStr + '1').getColumn(), rawNamesFromTargetSheet);


    let submitterActualGroup = null;
    for (const groupKey in submissionContext.getMembersByGroupName) {
      if (submissionContext.getMembersByGroupName[groupKey].includes(submitterFullName)) {
        submitterActualGroup = groupKey;
        break;
      }
    }
    if (!submitterActualGroup) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", `Submitter ${submitterFullName} not found in any group on sheet ${targetSheetName}.`, `Form: Group ${formDesignatedGroup}.`, activeSpreadsheetId);
      return;
    }
    if (submitterActualGroup !== formDesignatedGroup) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Info", `Submitter ${submitterFullName} (Actual Group: ${submitterActualGroup}) for form of Group ${formDesignatedGroup} ignored. Mismatch.`, `Sheet: ${targetSheetName}`, activeSpreadsheetId);
      return;
    }

    const submitterIndexInGroup = submissionContext.getIndexInGroupByName[submitterFullName];
    if (submitterIndexInGroup === undefined || submitterIndexInGroup < 0) {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Error", `Submitter ${submitterFullName} in group ${submitterActualGroup}, invalid index: ${submitterIndexInGroup}.`, `Sheet: ${targetSheetName}`, activeSpreadsheetId);
      return;
    }

    const numColsPerMemberPairInContext = 1 + (includeJustification ? 1 : 0);

    // Process item responses for scores, justifications, and feedback
    for (let k = 0; k < itemResponses.length; k++) {
      const itemResponse = itemResponses[k];
      const itemTitle = itemResponse.getItem().getTitle();

      if (itemTitle.startsWith('Rate ') && itemTitle.includes("'s performance")) {
        const ratedMemberName = itemTitle.substring('Rate '.length, itemTitle.indexOf("'s performance"));
        const score = itemResponse.getResponse();
        const rowForRatedMember = submissionContext.getRowByName[ratedMemberName];

        // The column for this score depends on the submitter's index in their group
        // Each member being rated has a block of columns (score + optional justification)
        // The submitter's responses fill one "M_i Score" (and "M_i Just.") slot for each person they rate.
        // The 'i' in "M_i" corresponds to the submitter's position in their own group.
        const scoreColumnForThisRating = submissionContext.evalTableStartColumn + (submitterIndexInGroup * numColsPerMemberPairInContext);

        if (rowForRatedMember > 0 && scoreColumnForThisRating > 0) {
          try {
            targetSheet.getRange(rowForRatedMember, scoreColumnForThisRating).setValue(parseFloat(score) || score);
          } catch (cellWriteError) {
            logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp CellWrite Error (Score)", `Rated: ${ratedMemberName} (R:${rowForRatedMember}), ScoreCol:${scoreColumnForThisRating}, Score:${score}. Err:${cellWriteError.message}`, activeSpreadsheetId);
          }

          // Check for and process justification if the flag is true and the next item is a justification
          if (includeJustification && (k + 1 < itemResponses.length)) {
            const nextItemResponse = itemResponses[k + 1];
            const nextItemTitle = nextItemResponse.getItem().getTitle();
            if (nextItemTitle.startsWith(`Justify why you gave ${ratedMemberName} that score`)) {
              const justificationText = nextItemResponse.getResponse();
              const justificationColumn = scoreColumnForThisRating + 1; // Justification is in the column next to the score
              try {
                targetSheet.getRange(rowForRatedMember, justificationColumn).setValue(justificationText);
              } catch (cellWriteError) {
                logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp CellWrite Error (Just.)", `Rated: ${ratedMemberName} (R:${rowForRatedMember}), JustCol:${justificationColumn}, Text:${justificationText}. Err:${cellWriteError.message}`, activeSpreadsheetId);
              }
              k++; // Increment k to skip this justification item in the next loop iteration
            }
          }
        } else {
          logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Invalid Row/Col (Score/Just.)", `Rated:${ratedMemberName}(R:${rowForRatedMember}). Submitter:${submitterFullName}(Idx:${submitterIndexInGroup}=>ScoreCol:${scoreColumnForThisRating})`, activeSpreadsheetId);
        }
      } else if (itemTitle.startsWith("Any other concerns or feedback") && includeFeedback) {
        const feedbackText = itemResponse.getResponse();
        const rowForSubmitterFeedback = submissionContext.getRowByName[submitterFullName];

        // Calculate feedback column: After all scores/justifications and Avg Score
        const actualMaxGroupSizeForSubmitter = submissionContext.maxGroupSize > 0 ? submissionContext.maxGroupSize : 1;
        const numColsForScoresAndJustificationsForSubmitter = actualMaxGroupSizeForSubmitter * numColsPerMemberPairInContext;
        const avgScoreColOffset = numColsForScoresAndJustificationsForSubmitter;
        const feedbackCol = submissionContext.evalTableStartColumn + avgScoreColOffset + 1; // +1 for Avg Score column itself

        if (rowForSubmitterFeedback > 0 && feedbackCol > 0) {
          try {
            targetSheet.getRange(rowForSubmitterFeedback, feedbackCol).setValue(feedbackText);
          } catch (cellWriteError) {
            logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp CellWrite Error (Feedback)", `Submitter: ${submitterFullName} (R:${rowForSubmitterFeedback}), FeedbackCol:${feedbackCol}. Err:${cellWriteError.message}`, activeSpreadsheetId);
          }
        } else {
          logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Invalid Row/Col for Feedback", `Submitter:${submitterFullName}(R:${rowForSubmitterFeedback}). FeedbackCol:${feedbackCol}`, activeSpreadsheetId);
        }
      }
    }

    // Update "Last Submitted" timestamp for the submitter
    const rowForSubmitterTimestamp = submissionContext.getRowByName[submitterFullName];
    const actualMaxGroupSizeForTimestamp = submissionContext.maxGroupSize > 0 ? submissionContext.maxGroupSize : 1;
    const numColsForScoresAndJustificationsForTimestamp = actualMaxGroupSizeForTimestamp * numColsPerMemberPairInContext;
    // Last Submitted is after Scores/Justifications, Avg Score, and Optional Feedback
    const lastSubmittedCol = submissionContext.evalTableStartColumn + numColsForScoresAndJustificationsForTimestamp + 1 + (includeFeedback ? 1 : 0);

    if (rowForSubmitterTimestamp > 0 && lastSubmittedCol > 0) {
      try {
        let formattedUtcOffsetForUtil = "GMT" + userUtcOffset.substring(0, 3) + ":" + userUtcOffset.substring(3, 5);
        const formattedTimestamp = Utilities.formatDate(submissionTimestamp, formattedUtcOffsetForUtil, "yyyy/MM/dd HH:mm:ss a");
        targetSheet.getRange(rowForSubmitterTimestamp, lastSubmittedCol).setValue(formattedTimestamp).setHorizontalAlignment("right");
      } catch (tsError) {
        logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp TimestampWrite Error", `Submitter:${submitterFullName}(R:${rowForSubmitterTimestamp}),Col:${lastSubmittedCol}. UTC:${userUtcOffset}. Err:${tsError.message}`, activeSpreadsheetId);
        targetSheet.getRange(rowForSubmitterTimestamp, lastSubmittedCol).setValue(submissionTimestamp).setNumberFormat("yyyy/mm/dd hh:mm:ss AM/PM");
      }
    } else {
      logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp Invalid Row/Col for Timestamp", `Submitter:${submitterFullName}(R:${rowForSubmitterTimestamp}). LastSubmittedCol:${lastSubmittedCol}`, activeSpreadsheetId);
    }
    Logger.log(`Form ${formId} by ${respondentEmail} (Name:${submitterFullName}, Group:${submitterActualGroup}) for sheet ${targetSheetName} (SS:${activeSpreadsheetId}) processed.`);
  } catch (error) {
    Logger.log(`onFormSubmit_WebApp CRITICAL Error: ${error.toString()}\nStack: ${error.stack}`);
    logErrorToDeveloperSheet(respondentEmail, formId, "onFormSubmit_WebApp CRITICAL Error", error.toString(), error.stack, activeSpreadsheetId || "N/A");
  }
}

function createUserClassroomAssignments(formURLs, formObject, userMappings, courseId, topicId) {
  const {
    getMembersByGroupName,
    getUserIdByName
  } = userMappings;
  const dueDateTime = new Date(`${formObject.dueDate}T${formObject.dueTime}${formObject.utcOffset}`);
  const dueDate = {
    year: dueDateTime.getUTCFullYear(),
    month: dueDateTime.getUTCMonth() + 1,
    day: dueDateTime.getUTCDate()
  };
  const dueTime = {
    hours: dueDateTime.getUTCHours(),
    minutes: dueDateTime.getUTCMinutes(),
    seconds: 0
  };
  const formattedDueDate = `${dueDate.year}/${(dueDate.month).toString().padStart(2, '0')}/${dueDate.day.toString().padStart(2, '0')}`;
  const formattedDueTime = dueDateTime.toLocaleTimeString(undefined, {
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  });
  let scheduledTimeISO = null;
  if (formObject.assignmentScheduleDate && formObject.assignmentScheduleTime) {
    try {
      const scheduleDateTime = new Date(`${formObject.assignmentScheduleDate}T${formObject.assignmentScheduleTime}${formObject.utcOffset}`);
      if (scheduleDateTime > new Date(new Date().getTime() + 5 * 60000 + 1000)) {
        scheduledTimeISO = scheduleDateTime.toISOString();
      } else {
        Logger.log(`Scheduled time ${scheduleDateTime.toISOString()} not sufficiently in future. Publishing immediately.`);
      }
    } catch (dateError) {
      Logger.log(`Error parsing schedule date/time: ${dateError}. Publishing immediately.`);
    }
  }
  Object.keys(getMembersByGroupName).forEach(groupName => {
    const groupMembers = getMembersByGroupName[groupName];
    if (!formURLs[groupName]) {
      Logger.log(`Skipping assignment for Group ${groupName} (no URL).`);
      return;
    }
    const studentIdsForAssignment = groupMembers.map(memberName => getUserIdByName[memberName]).filter(id => id);
    if (studentIdsForAssignment.length === 0) {
      Logger.log(`Skipping assignment for Group ${groupName} (no valid student IDs).`);
      return;
    }
    const assignment = {
      title: `${formObject.assignmentTitle} - Group ${groupName}`,
      description: `Please complete this peer evaluation before ${formattedDueDate}, ${formattedDueTime}.\n` +
        `Form Link: ${formURLs[groupName]}\n\n` +
        `This form is for Group ${groupName} only.\nGroup Members: ${groupMembers.join(', ')}\n\n` +
        `${formObject.assignmentDescription}`,
      workType: 'ASSIGNMENT',
      maxPoints: formObject.assignmentPoints,
      topicId: topicId,
      dueDate: dueDate,
      dueTime: dueTime,
      assigneeMode: 'INDIVIDUAL_STUDENTS',
      individualStudentsOptions: {
        studentIds: studentIdsForAssignment
      },
      materials: [{
        link: {
          url: formURLs[groupName]
        }
      }]
    };
    if (scheduledTimeISO) {
      assignment.scheduledTime = scheduledTimeISO;
    } else {
      assignment.state = 'PUBLISHED';
    }
    try {
      const createdAssignment = Classroom.Courses.CourseWork.create(assignment, courseId);
      Logger.log(`Assignment for Group ${groupName} state: ${createdAssignment.state}. ID: ${createdAssignment.id}. Scheduled: ${createdAssignment.scheduledTime || 'No'}`);
    } catch (e) {
      Logger.log(`Error creating assignment for Group ${groupName}: ${e.toString()}`);
      let errorDetails = e.message;
      if (e.details && e.details.error && e.details.error.message) {
        errorDetails += ` | API Error: ${e.details.error.message}`;
      } else if (e.details && e.details.errors && e.details.errors.length > 0 && e.details.errors[0].message) {
        errorDetails += ` | API Error(s): ${e.details.errors.map(err => err.message).join('; ')}`;
      }
      logErrorToDeveloperSheet(Session.getActiveUser().getEmail(), SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : "N/A", `Classroom Assignment Error for Group ${groupName}`, errorDetails, e.stack);
    }
  });
}

function logErrorToDeveloperSheet(userEmail, targetSpreadsheetId, type, message, details, stack) {
  try {
    if (!WEBAPP_ERROR_LOG_SPREADSHEET_ID || WEBAPP_ERROR_LOG_SPREADSHEET_ID === 'YOUR_DEVELOPER_ERROR_LOG_SPREADSHEET_ID') {
      console.error(`WEBAPP_ERROR_LOG_SPREADSHEET_ID not set. Error: User:${userEmail}, TargetSS:${targetSpreadsheetId}, Type:${type}, Msg:${message}`);
      return;
    }
    const ss = SpreadsheetApp.openById(WEBAPP_ERROR_LOG_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(WEBAPP_ERROR_LOG_SHEET_NAME);
    if (sheet) {
      sheet.appendRow([new Date(), userEmail || Session.getActiveUser().getEmail(), targetSpreadsheetId || 'N/A', type, message, (typeof details === 'string' ? details : JSON.stringify(details)).substring(0, 25000), (typeof stack === 'string' ? stack : '').substring(0, 25000)]);
    } else {
      console.error(`WebApp error log sheet "${WEBAPP_ERROR_LOG_SHEET_NAME}" not found in SS ID "${WEBAPP_ERROR_LOG_SPREADSHEET_ID}". Original error: ${type} - ${message}`);
    }
  } catch (e) {
    console.error(`CRITICAL: Could not log error to developer sheet: ${e.toString()}. Original Error: User:${userEmail}, TargetSS:${targetSpreadsheetId}, Type:${type}, Msg:${message}`);
  }
}