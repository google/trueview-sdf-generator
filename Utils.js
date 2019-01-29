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
 * @fileoverview Contains various utility functions.
 */

/**
 * Given a source sheet, populates an object with the column names as key
 * and an array of the column values as value. Yo.
 * @param {!Object} sheet The source sheet.
 * @return {!Object} The object populated with the source sheet values.
 * @private
 */
function populateObject_(sheet) {
  var object = {};
  var values = sheet.getDataRange().getValues();
  for (col in values[0]) {
    var contentArray = [];
    for (row in values) {
      contentArray.push(values[row][col]);
    }
    object[values[0][col]] = contentArray;
  }
  return object;
}
