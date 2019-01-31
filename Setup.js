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
 * @fileoverview Variables and methods for the initial setup of the main sheets.
 */

var INPUT_HEADERS = ['ID', 'NAME', 'YOUTUBE_ID', 'DISPLAY_URL', 'LANDING_URL'];
var STRUCTURE_HEADERS = [
  'IO-SDF',
  'IO-SDF-defaults',
  'LI-SDF',
  'LI-SDF-defaults',
  'AdGroup-SDF',
  'AdGroup-SDF-defaults',
  'Ad-SDF',
  'Ad-SDF-defaults'
];
var MAPPING_HEADERS = ['INPUT_COLUMN', 'FULL_NAME', 'DV360_ID'];

/**
 * Sets up and formats the needed sheets in the Spreadsheet: "INPUT" with the
 * values for the different TV ads, "STRUCTURE_AND_DEFAULTS" for the fields
 * and default values of the SDF objects, "NAMES_IDS_MAPPING" for the match
 * table to convert readable names to DV360-friendly IDs.
 */
function initialSetup() {
  // Creates the MAPPING sheet.
  if (mappingSheet) {
    doc.deleteSheet(mappingSheet);
  }
  doc.insertSheet(MAPPING_NAME,0);
  mappingSheet = doc.getSheetByName(MAPPING_NAME);
  mappingSheet.setTabColor('black');
  mappingSheet
    .getRange(1,1,1,MAPPING_HEADERS.length)
    .setValues([MAPPING_HEADERS])
    .setBackground('lightblue')
    .setFontWeight('bold');
  // Creates the INPUT sheet.
  if (inputSheet) {
    doc.deleteSheet(inputSheet);
  }
  doc.insertSheet(INPUT_NAME,0);
  inputSheet = doc.getSheetByName(INPUT_NAME);
  inputSheet.setTabColor('red');
  inputSheet
    .getRange(1,1,1,INPUT_HEADERS.length)
    .setValues([INPUT_HEADERS])
    .setBackground('lightgreen')
    .setFontWeight('bold');
  // Creates the STRUCTURES_AND_DEFAULT sheet.
  if (structuresSheet) {
    doc.deleteSheet(structuresSheet);
  }
  doc.insertSheet(STRUCTURE_NAME,0);
  structuresSheet = doc.getSheetByName(STRUCTURE_NAME);
  structuresSheet.setTabColor('red');
  structuresSheet
      .getRange(1, 1, structuresSheet.getMaxRows(),
                structuresSheet.getMaxColumns())
      .setNumberFormat("@");
  structuresSheet
    .getRange(1,1,1,STRUCTURE_HEADERS.length)
    .setValues([STRUCTURE_HEADERS])
    .setBackground('orange')
    .setFontWeight('bold');

  for (i in SDF) {
    var fields = SDF[i].fields;
    for (row in fields) {
      var field = fields[row];
      structuresSheet.getRange(+row + 2, i * 2 + 1, 1, 2)
          .setValues([[field[0], field[1]]]);
      if (field[2].length > 2) {
        structuresSheet.getRange(+row + 2, i * 2 + 2)
            .setBackground(field[2]);
      }
    }
    structuresSheet
        .getRange(2, i * 2 + 1, fields.length, 1)
        .setBackground('lightgray')
        .setFontWeight('bold');
  }
  for (c = 1; c <= STRUCTURE_HEADERS.length; c++) {
    structuresSheet.setColumnWidth(c, 280);
  }
  structuresSheet
      .getRange(2, 1, structuresSheet.getMaxRows() - 1,
                structuresSheet.getMaxColumns())
      .setFontSize(9)
      .setNumberFormat("@");
}
