/***********************************************************************
Copyright 2019 Google Inc.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Note that these code samples being shared are not official Google
products and are not formally supported.
************************************************************************/

/**
 * @fileoverview Main script file with menu setup and main functions.
 */

/**
 * @OnlyCurrentDoc
 */

// Global references to the main doc and sheets.
var INPUT_NAME = 'INPUT';
var STRUCTURE_NAME = 'STRUCTURE_AND_DEFAULTS';
var MAPPING_NAME = 'NAME_IDS_MAPPING';
var IO_NAME = 'IO-SDF';
var LI_NAME = 'LI-SDF';
var AG_NAME = 'AdGroup-SDF';
var AD_NAME = 'Ad-SDF';
var SHEET_NAMES = [IO_NAME, LI_NAME, AG_NAME, AD_NAME];

var ss = SpreadsheetApp.getActive();
var inputSheet = ss.getSheetByName(INPUT_NAME);
var structuresSheet = ss.getSheetByName(STRUCTURE_NAME);
var mappingSheet = ss.getSheetByName(MAPPING_NAME);
var numberOfVideos = 0;
var errorMessage, settings;
var structures = {};
var input = {};
var createdGroups = [];

/**
 * Creates and adds the custom menu to the toolbar.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TrueView SDF Generator')
      .addItem('Initial Setup (resets everything!)', 'initialSetup')
      .addSeparator()
      .addItem('Convert Names to IDs', 'convertToIds')
      .addSeparator()
      .addItem('Generate SDF sheets', 'generateSdf')
      .addToUi();
}

/**
 * Reads the INPUT and STRUCTURE sheets and generates the output SDF-compatible
 * sheets.
 */
function generateSdf() {
  numberOfVideos = inputSheet.getLastRow() - 1;
  Logger.log('Number of videos detected: ' + numberOfVideos);
  errorMessage = '';
  structures = populateObject_(structuresSheet);
  input = populateObject_(inputSheet);
  createSheets_();
  createInsertionOrder_();
  for (var i = 0; i < numberOfVideos; i++) {
    // If multiple ads per adgroup, check if ad's LI/AdGroup with
    // that ID has already been created
    if (createdGroups.indexOf(input['ID'][i]) < 0) {
      createLineItems_(i);
      createAdGroups_(i);
      createdGroups.push(input['ID'][i]);
    }
    createAds_(i);
  }
  if (errorMessage.length > 2) {
    Browser.msgBox(errorMessage);
  } else {
    Browser.msgBox('SDF sheets generated without (apparent) errors');
  }
}

/**
 * Creates the (empty) destination sheets for the SDF data.
 * @private
 */
function createSheets_() {
  for (var i = 0; i < SHEET_NAMES.length; i++) {
    var sheetName = SHEET_NAMES[i];
    Logger.log('Creating/Erasing sheet: ' + sheetName);
    var sheet = ss.getSheetByName(sheetName) ?
        ss.getSheetByName(sheetName) :
        ss.insertSheet(sheetName, ss.getNumSheets());
    sheet.clearContents();
  }
}

/**
 * Populates the headers of the specified sheet getting the column names from
 * the STRUCTURE_AND_DEFAULTS sheet.
 * @param {string} sheetName Name of the sheet to set the headers of.
 * @private
 */
function populateHeaders_(sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  sheet.appendRow(structures[sheetName]);
  sheet.getRange(1, 1, 1, 100).setBackground('yellow');
}

/**
 * Replaces every %PLACEHOLDER% in an array of default values with the
 * corresponding values matched in each row the INPUT sheet.
 * @param {?Array<string>} valuesArray Default values (with placeholders).
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @return {?Array<string>} The default values with the replaced placeholders.
 * @private
 */
function replacePlaceholders_(valuesArray, index) {
  for (i in valuesArray) {
    for (col in input) {
      valuesArray[i] =
          valuesArray[i].replace('%' + col + '%', input[col][index]);
    }
    if (valuesArray[i].indexOf('%') >= 0) {
      // Couldn't find the right placeholder
      errorMessage += 'Found a placeholder with no matching value: ' +
          valuesArray[i] + '\\n';
    }
  }
  return valuesArray;
}

/**
 * Generates the Insertion Order data (and populates the corresponding sheet).
 * @private
 */
function createInsertionOrder_() {
  var sheetName = IO_NAME;
  var sheet = ss.getSheetByName(sheetName);
  populateHeaders_(sheetName);
  // Load default values
  var defaultValues = structures[sheetName + '-defaults'].slice();
  var newValues = [[]];
  newValues[0] = defaultValues;
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newValues[0].length)
      .setNumberFormat('@STRING@')
      .setValues(newValues);
}

/**
 * Generates the Line Items data (and populates the corresponding sheet).
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @private
 */
function createLineItems_(index) {
  var sheetName = LI_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (index == 0) {
    populateHeaders_(sheetName);
  }
  // Load default values and replace placeholders
  var defaultValues = structures[sheetName + '-defaults'].slice();
  var newValues = [[]];
  newValues[0] = replacePlaceholders_(defaultValues, index);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newValues[0].length)
      .setNumberFormat('@STRING@')
      .setValues(newValues);
}

/**
 * Generates the AdGroups data (and populates the corresponding sheet).
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @private
 */
function createAdGroups_(index) {
  var sheetName = AG_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (index == 0) {
    populateHeaders_(sheetName);
  }
  // Load default values and replace placeholders
  var defaultValues = structures[sheetName + '-defaults'].slice();
  var newValues = [[]];
  newValues[0] = replacePlaceholders_(defaultValues, index);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newValues[0].length)
      .setNumberFormat('@STRING@')
      .setValues(newValues);
}

/**
 * Generates the Ads data (and populates the corresponding sheet).
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @private
 */
function createAds_(index) {
  var sheetName = AD_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (index == 0) {
    populateHeaders_(sheetName);
  }
  // Load default values and replace placeholders
  var defaultValues = structures[sheetName + '-defaults'].slice();
  var newValues = [[]];
  newValues[0] = replacePlaceholders_(defaultValues, index);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newValues[0].length)
      .setNumberFormat('@STRING@')
      .setValues(newValues);
}

/**
 * Converts human-readable values (e.g. for Geotargeting locations) in the INPUT
 * sheet to the corresponding ID accepted by DV360, using a matching table that
 * the user must provide in the NAME_IDS_MAPPING sheet.
 */
function convertToIds() {
  var lookupObj = populateObject_(mappingSheet);
  var inputs = inputSheet.getDataRange().getValues();
  var numberOfRows = inputs.length;
  for (var lookupIndex = 0; lookupIndex < lookupObj['INPUT_COLUMN'].length;
       lookupIndex++) {
    var inputColumnName = lookupObj['INPUT_COLUMN'][lookupIndex];
    var textToLookFor = lookupObj['FULL_NAME'][lookupIndex];
    var dv360Id = lookupObj['DV360_ID'][lookupIndex];
    for (var inputColIndex = 0; inputColIndex < inputs[0].length;
         inputColIndex++) {
      if (inputs[0][inputColIndex] == inputColumnName) {
        // Found the correct corresponding column in the INPUT sheet
        for (var row = 1; row < numberOfRows; row++) {
          // Looping through the column to replace the names with IDs
          var currentString = inputs[row][inputColIndex];
          var oldString = currentString;
          currentString = currentString.toString().replace(
              textToLookFor, dv360Id);
          if (currentString != oldString) {
            inputSheet.getRange(row + 1, inputColIndex + 1, 1, 1)
                .setNumberFormat('@STRING@')
                .setValue(currentString);
          }
        }
      }
    }
  }
}
