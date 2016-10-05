// Copyright 2016 Martin Hawksey. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/** one time function used to set credentials
/** Stored in Google User Properties

function setupOneTimeOnly(){
  CRM.setAPIBase('[this_bit]'); // api_base [this_bit]/civicrm/extern/rest.php with either http or https
  CRM.setApiKey('YOUR_API_KEY'); // https://wiki.civicrm.org/confluence/display/CRMDOC/REST+interface#RESTinterface-CreatingAPIkeysforusers
  CRM.setKey('YOUR_SITE_SECRET'); // https://wiki.civicrm.org/confluence/display/CRMDOC/Managing+Scheduled+Jobs#ManagingScheduledJobs-ConfiguringyourSiteKey
}

alternatively include the following in each script (this allows credenitals to be shared if the script is being run by someone else other than the script owner ...

CRM.init({api_base:'[this_bit]', // api_base [this_bit]/civicrm/extern/rest.php with either http or https
          api_key: 'YOUR_API_KEY',
          key:     'YOUR_API_KEY',        
});

... if you use script properties to store these you can also add using 

**/

// Config stored in Script Properties (accessed/editted via File > Project properties and Script Properties tab with key/values for api_base, api_key and key)
var config = PropertiesService.getScriptProperties().getProperties();
CRM.init(config);

// Adding some basic UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CiviCRM')
      .addItem('Add events', 'processEvents')
      .addToUi();
}

// function we call to process the events submitted 
function processEvents(){
  var doc = SpreadsheetApp.getActiveSheet();
  var data = doc.getRange(1, 1, doc.getLastRow(), doc.getLastColumn()).getValues();
  var dataObj = rangeToObjects(data);
  for (i in dataObj){
    // if action is 'Add to Civi' and no link then process
    if (dataObj[i].action === "Add to Civi" && dataObj[i].link === ""){
      addEvent(dataObj[i]); // Civi API call to add event
    }
  }
}

// add event to civi with row details
function addEvent(evnt){
  CRM.api3('Event', 'create', {
    "sequential": 1,
    "title": evnt.eventTitle,
    "description": evnt.eventDescription,
    "start_date": combineDateTime(evnt.startDate, evnt.startTime), // As Google Forms have seperate date and time form elements combine these e.g. 201610050930
    "end_date": combineDateTime(evnt.startDate, evnt.endTime),
    "event_type_id": getEventTypeValue(evnt.eventType), // BONUS if our event labels and values don't match to addition lookup
    "is_active": 1,
    "is_public": 0
  }).done(function(result) {
    if(!result.is_error){
      // if event added put url in the sheet
      var doc = SpreadsheetApp.getActiveSheet();
      // assuming Link column is 17th along...
      doc.getRange(parseInt(evnt.rowNum)+1, 17).setValue("https://www.alt.ac.uk/" +
                                                         "civicrm/event/info?reset=1&id=" 
                                                         + result.values[0].id);
    };
  });
}

/**
* combine a date and time into a string of yyyyMMddhhmmss
*
* @param {Date} date.
* @param {Date} time.
* @return {string} of date+time.
*/
function combineDateTime(date, time){
  return Utilities.formatDate(new Date(date.getFullYear(), date.getMonth(), date.getDate(), 
                                       time.getHours(), time.getMinutes(), time.getSeconds()), 
                              Session.getScriptTimeZone(), 
                              "yyyyMMddhhmmss");
}



// BONUS helper function to get event type value from a label
function getEventTypeValue(label){
  var id;
  CRM.api3('OptionValue', 'get', {
    "sequential": 1,
    "return": ["value"],
    "label": label
  }).done(function(result) {
    id = result.values[0].value;
  });
  return id;
}

// BONUS function that writes our event type list to a sheet and the google form
function updateEventTypeList(){
  CRM.api3('OptionValue', 'get', {
    "sequential": 1,
    "option_group_id": "event_type",
    "return": ["label","value"],
    "is_active": 1
  }).done(function(result) {
    if(!result.is_error){      
      var sheet_out = [['label','value']];
      var choice_out =[];
      for (i in result.values){
        sheet_out.push([result.values[i].label, result.values[i].value]); 
        choice_out.push(result.values[i].label);
      }
            
      var doc = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = doc.getSheetByName('EventType');
      sheet.getRange(1, 1, sheet_out.length, 2).setValues(sheet_out);
      
      var form = FormApp.openById('14Y4WKrha0CzjgiITjPjnY1PbR0_urD0elBq_4lQERlo');
      var items = form.getItems();
      var eventTypeItem;
      items.map(function(item, index) {
        if (item.getTitle() == 'Event Type') {
          eventTypeItem = item;
        }
      });
      eventTypeItem.asListItem().setChoiceValues(choice_out);
    }
  });
}

