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

var doc = SpreadsheetApp.getActive();
var inputSheet = doc.getSheetByName(INPUT_NAME);
var structuresSheet = doc.getSheetByName(STRUCTURE_NAME);
var mappingSheet = doc.getSheetByName(MAPPING_NAME);
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
  errorMessage = '';
  structures = populateObject_(structuresSheet);
  input = populateObject_(inputSheet);
  createSheets_();
  populateSheet_(IO_NAME, 1);
  for (var i = 0; i < (inputSheet.getLastRow() - 1); i++) {
    // If multiple ads per adgroup, check if ad's LI/AdGroup with
    // that ID has already been created
    if (createdGroups.indexOf(input['ID'][i]) < 0) {
      populateSheet_(LI_NAME, i);
      populateSheet_(AG_NAME, i);
      createdGroups.push(input['ID'][i]);
    }
    populateSheet_(AD_NAME, i);
  }
  if (errorMessage) {
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
    var sheet = doc.getSheetByName(sheetName);
    if (!sheet) {
      sheet = doc.insertSheet(sheetName, doc.getNumSheets());
    }
    sheet.clearContents();
    sheet.appendRow(structures[sheetName]);
    sheet.getRange(1, 1, 1, 100).setBackground('yellow');
  }
}

/**
 * Replaces every %%PLACEHOLDER%% in an array of default values with the
 * corresponding values matched with each row of the INPUT sheet.
 * @param {?Array<string>} valuesArray Default values (with placeholders).
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @return {?Array<string>} The default values with the replaced placeholders.
 * @private
 */
function replacePlaceholders_(valuesArray, index) {
  for (i in valuesArray) {
    for (colName in input) {
      valuesArray[i] =
          valuesArray[i].replace('%%' + colName + '%%', input[colName][index]);
    }
    if (valuesArray[i].indexOf('%%') >= 0) {
      errorMessage += 'Found a placeholder with no matching value: ' +
          valuesArray[i] + '\\n';
    }
  }
  return valuesArray;
}

/**
 * Populates a sheet using the corrisponding content using default
 * data and replacing placeholders for each entry.
 * @param {!string} sheetName Name of the sheet to populate.
 * @param {number} index Index of the currently evaluated entry from INPUT.
 * @private
 */
function populateSheet_(sheetName, index) {
  var sheet = doc.getSheetByName(sheetName);
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
  var mappingObj = populateObject_(mappingSheet);
  var inputs = inputSheet.getDataRange().getValues();
  // Loops through all the entries from the NAME_IDS_MAPPING "match table".
  for (
      var mappingIndex = 0;
      mappingIndex < mappingObj['INPUT_COLUMN'].length;
      mappingIndex++) {
    var inputColumnName = mappingObj['INPUT_COLUMN'][mappingIndex];
    var textToLookFor = mappingObj['FULL_NAME'][mappingIndex];
    var dv360Id = mappingObj['DV360_ID'][mappingIndex];
    // Loops through all the entries from the INPUT list.
    for (
        var inputColIndex = 0;
        inputColIndex < inputs[0].length;
        inputColIndex++) {
      if (inputs[0][inputColIndex] == inputColumnName) {
        // Found the correct corresponding column in the INPUT list,
        // looping now through the INPUT rows to replace names with IDs.
        for (var rowIndex = 1; rowIndex < inputs.length; rowIndex++) {
          var currentString = inputs[rowIndex][inputColIndex];
          var oldString = currentString;
          currentString =
              currentString.toString().replace(textToLookFor, dv360Id);
          if (currentString != oldString) {
            inputSheet.getRange(rowIndex + 1, inputColIndex + 1, 1, 1)
                .setNumberFormat('@STRING@')
                .setValue(currentString);
          }
        }
      }
    }
  }
}
