/*
   ObjService
 
   Copyright (c) 2011 James Ferreira

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

/**
 * ObjService
 * @author James Ferriera
 * @documentation http://goo.gl/JdEHW
 *
 * Changes an object like e.parameter into a 2D array useful in 
 * writting to a spreadsheet with using the .setValues method
 *
 * @param   {Array}   headers    [header, header, ...] 
 * @param   {Array}   objValues  [{key:value, ...}, ...]
 * @returns {Array}              [[value, value, ...], ...]
 */
function objectToArray(headers, objValues){
  var values = [];  
  var headers = camelArray(headers);  
  for (var j=0; j < objValues.length; j++){
    var rowValues = [];
    for (var i=0; i < headers.length; i++){
      rowValues.push(objValues[j][headers[i]]);
    }  
    values.push(rowValues);
  } 
  return values;
}


/**
 * Changes a range array often returned from .getValues() into an 
 * array of objects with key value pairs.
 * The first element in the array is used as the keys (headers)
 *
 * @param   {Array}   range   [[key, key, ...],[value, value, ...]] 
 * @returns {Array}           [{key:value, ...}, ...] 
 */
function rangeToObjects(range){
  var headers = range[0];
   var values = range;
  var rowObjects = [];
  for (var i = 1; i < values.length; ++i) {
    var row = new Object();
    row.rowNum = i;
    for (var j in headers){
      row[camelString(headers[j])] = values[i][j];
    }
   rowObjects.push(row); 
  }   
  return rowObjects;
}

/**
 * Changes a range array into an array of objects with key value pairs
 *
 * @params  {array}    headers  [key, key, ...]
 * @params  {array}    values    [[value, value, ...], ...]
 * @returns {array}    [{key:value, ...}, ...]  
 */
function splitRangesToObjects(headers, values){
  var rowObjects = [];
  for (var i = 1; i < values.length; ++i) {
    var row = new Object();
    row.rowNum = i;
    for (var j in headers){
      row[camelString(headers[j])] = values[i][j];
    }
   rowObjects.push(row); 
  }   
  return rowObjects;
}



/**
 * Removes special characters from strings in an array
 * Commonly know as a camelCase, 
 * Examples:
 *   "First Name" -> "firstName"
 *   "Market Cap (millions) -> "marketCapMillions
 *   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 * @params  {array} headers   [string, string, ...]
 * @returns {array}           camelCase 
 */
function camelArray(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = camelString(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}


/**
 * Removes special characters from a string
 * Commonly know as a camelCase, 
 * Examples:
 *   "First Name" -> "firstName"
 *   "Market Cap (millions) -> "marketCapMillions
 *   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 * @params  {string}  header   string
 * @returns {string}           camelCase 
 */
function camelString(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}
function isDigit_(char) {
  return char >= '0' && char <= '9';
}