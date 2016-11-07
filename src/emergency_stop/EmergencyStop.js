/***********************************************************************************************
*
*  MCC-Script - Emergency Stop using Google Spreadsheet
*
*  Copyright 2015 crealytics GmbH
*          
*  Licensed under the Apache License, Version 2.0 (the "License");
*  you may not use this file except in compliance with the License.
*  You may obtain a copy of the License at
*                      
*  http://www.apache.org/licenses/LICENSE-2.0
*                      
*  Unless required by applicable law or agreed to in writing, software
*  distributed under the License is distributed on an "AS IS" BASIS,
*  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
*  See the License for the specific language governing permissions and
*  limitations under the License.
*
*  @author: Alexander Giebelhaus
*  @version: 1.1 (Script-Version)
*
***********************************************************************************************/

// Script Settings
var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1Cad0D7_GWfvUwrQoN_lW0vb8sSTCpNh22r4HkLb6FMw/edit#gid=325670813";
var PAUSED_LABEL = "EmergencyStop";
var EMAIL_SUBJECT = "EmergencyStop";

// Script to sheet mapping
var SHEETCOLS = {
  'ACCOUNT_IDS':0,
  'ACCOUNTNAMES':1,
  'EMAILS':8,
  'PAUSEDDATE':6,
  'ACTIVATEDDATE':7,
  'PAUSED_CAMPAIGNS':3,
  'ESTOP':5,
  'ERRORS':9
};

// Get data from setting tab
var SPREADSHEET = null;
var SHEET = null;
var DATA = null;
var SP_RANGE = null;

try{
   SPREADSHEET = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
   SHEET = SPREADSHEET.getSheetByName("Settings");
   SP_RANGE = SPREADSHEET.getRangeByName("settings");
   DATA = checkSpreadsheedData(SPREADSHEET.getRangeByName("settings").getValues());
}catch(e){
  throw 'CAN NOT ACCESS SPREADSHEET. PLEASE CHECK THE URL AND YOUR PERMISSIONS TO THE SHEET (for the user who runs the script in AdWords)!';
}

var actionEnum = {
  YES: 'yes',
  NO: 'no'
};

/**
  * @desc This Function will be called from google to start the script
*/
function main() {  
  if(SPREADSHEET_URL == 'https://docs.google.com/spreadsheets/d/1Cad0D7_GWfvUwrQoN_lW0vb8sSTCpNh22r4HkLb6FMw/edit#gid=325670813'){
        throw 'SPREADSHEET_URL WAS NOT REPLACED TO YOUR OWN SPREADSHEET!';
  }
  
  if (DATA && DATA.length > 0) {
    // Execute all accounts in parallel
    MccApp.accounts().withIds(getColumnFromRange(DATA,SHEETCOLS.ACCOUNT_IDS)).executeInParallel("execInParallel","finalExecution");
  }
}

/**
  * @desc Executes an account in parallel
  * @return dataObject
*/
function execInParallel(){
  var account = AdWordsApp.currentAccount();  
  var action = null; // Contains either the action, 'yes' to stop, 'no' in case of reactivation or null if no action is going to happen.
  var campaigns = null;
  var shoppingCampaigns = null;

  // Initialize a default data object
  var dataObject = {
    accountid: 0,
    accountname: '',
    activated_count: 0,
    paused_count: 0,
    paused_campaigns: 0,
    paused_shopping_campaigns: 0,
    sheet_current_row_index: 0,
    emails: null,
    errors: null
  };

  // Set the actual values for the dataObject
  dataObject.accountid = account.getCustomerId();
  dataObject.sheet_current_row_index = getColumnFromRange(DATA, SHEETCOLS.ACCOUNT_IDS).indexOf(dataObject.accountid);
  dataObject.accountname = account.getName();

  // Get all informations of the sheet row based on the current accountId e.g. email, scheduling data
  var account_data = getRowInRangeById(DATA, dataObject.accountid);

  // Get the email address(es)
  var emails = extractAndCheckEmail(account_data);
  if(emails == -1){
    dataObject.errors = "invalid email address";
  }else{
    dataObject.emails = emails;    
  }

  // If a schedule date is set, calculate the date range and set action to 'yes' when current date is in range ('no' if the date is after the activation date)
  // else read only the value of the EmergencyStop cell and set it as action value.
  if(account_data[SHEETCOLS.PAUSEDDATE] !== "" && account_data[SHEETCOLS.ACTIVATEDDATE] !== ""){
      var now = new Date();
      var pausedDate = new Date(account_data[SHEETCOLS.PAUSEDDATE]);
      var activationDate = new Date(account_data[SHEETCOLS.ACTIVATEDDATE]);

      if(Number(pausedDate) <= Number(now) && Number(activationDate) > Number(now)){
        action = actionEnum.YES;
      }else if(Number(activationDate) <= Number(now)){
        action = actionEnum.NO;
      }
     // Change EmergencyStop value in the spreadsheet based on the current action
     SP_RANGE.getCell(dataObject.sheet_current_row_index+1,SHEETCOLS.ESTOP+1).setValue(action);
  }else{
    action = account_data[5];
  }

  // Create label if it does not exist
  createLabelIfNotExist();

  // Get the campaigns to iterate over (depending on the action and type of campaign)
  if(action == actionEnum.YES){
    campaigns = AdWordsApp.campaigns().withCondition("Status = ENABLED").get();
    shoppingCampaigns = AdWordsApp.shoppingCampaigns().withCondition("Status = ENABLED").get();
  }else if(action == actionEnum.NO){
    campaigns = AdWordsApp.campaigns().withCondition("LabelNames CONTAINS_ANY ['"+PAUSED_LABEL+"']").withCondition("Status = PAUSED").get();
    shoppingCampaigns = AdWordsApp.shoppingCampaigns().withCondition("LabelNames CONTAINS_ANY ['"+PAUSED_LABEL+"']").withCondition("Status = PAUSED").get();
  }

  // Analyse the campaigns (normal ones and shopping) and update the data object, if an action is set and no error occured
  if(action !== null && dataObject.errors === null){
    if (campaigns !== null && campaigns.totalNumEntities() >  0) {
      dataObject = analyseCampaigns(campaigns,action,dataObject);
    }
    if (shoppingCampaigns !== null && shoppingCampaigns.totalNumEntities() >  0) {
      dataObject = analyseCampaigns(shoppingCampaigns,action,dataObject);
    }
  }

  // Get and set the total amount of 'normal' paused campaigns with label
  dataObject.paused_campaigns = (AdWordsApp.campaigns().withCondition("LabelNames CONTAINS_ANY ['"+PAUSED_LABEL+"']").withCondition("Status = PAUSED").get().totalNumEntities());

  // Get and set the total amount of shopping campaigns which are paused
  dataObject.paused_shopping_campaigns = (AdWordsApp.shoppingCampaigns().withCondition("LabelNames CONTAINS_ANY ['"+PAUSED_LABEL+"']").withCondition("Status = PAUSED").get().totalNumEntities());

  // Return the data object as String to the callback function (once all the accounts have been processed).
  return JSON.stringify(dataObject);
}

/**
  * @desc Analyzes the campaigns and pauses or reactivates a campaign.
  * @param AdWordsApp.CampaignIterator campaigns, string action, dataObject dataObject 
  * @return dataObject - updated data object
*/
function analyseCampaigns(campaigns, action, dataObject){
   var paused = false;
   var activated = false;

    while (campaigns.hasNext()) {

      var campaign = campaigns.next();
      var campaign_name = campaign.getName();

      Logger.log("analysing campaign: "+campaign_name+" in account: "+dataObject.accountname);

      if(action == actionEnum.NO){
        activated = activateCampaign(campaign);
        if(activated){
          dataObject.activated_count++;
          Logger.log("--> activated");
        }
      }
      else if(action == actionEnum.YES){
        paused = pauseCampaign(campaign);
        if(paused){
          dataObject.paused_count++;
          Logger.log("--> paused");
        }
      }    
  }
  return dataObject;
}

/**
  * @desc Checks if an given email is valid
  * @param string email
  * @return boolean - true if the email is valid else false
*/
function validateEmail(email){
    var re = /^(([^<>()[\]\.,;:\s@\"]+(\.[^<>()[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})$/i;
    return re.test(email);
}

/**
  * @desc Send an Email to the given addresses 
  * @param Array/String email, string subject, string text 
*/
function sendEmail(email,subject,text)
{
  var emailaddys;

  // Check if the email is an array else convert it to an array
  if(Array.isArray(email)){
    emailaddys=email;
  }else{    
    emailaddys = new Array(email);
  }

  // Send email to the given addresses
  for(var i = 0; i < emailaddys.length; i++){
    if(emailaddys[i] !== null && validateEmail(emailaddys[i].trim())){
     if (!AdWordsApp.getExecutionInfo().isPreview()) {
        MailApp.sendEmail({
          to: emailaddys[i].trim(),
          subject: subject,
          htmlBody: text
        });
      }
    }
  }
}

/**
  * @desc Extracts the emails from the spreadsheet account data
  * @param account_data
  * @return array (but -1 if an error occured and null if no email was given)
*/
function extractAndCheckEmail(account_data){
    var emails=null;
    var valid=true;

    // Get the email address(es) from the corresponding sheet cell
    if(account_data[SHEETCOLS.EMAILS] !== ''){
      if(account_data[SHEETCOLS.EMAILS].indexOf(',') != -1){
        emails = account_data[SHEETCOLS.EMAILS].split(",");
      }else{      
        emails = new Array(account_data[SHEETCOLS.EMAILS]);
      }
    }

    // Check if the email address is valid
    if(emails !== null){
      for(var i = 0; i < emails.length; i++){
        if(emails[i] !== null && emails[i].length > 1){
            if (!validateEmail(emails[i].trim())) {
                valid = false;
                Logger.log("invalid email " + emails[i].trim());
            }
        }
      }
    }

    if(!valid) emails = -1;

    return emails;
  }

/**
  * @desc Once all the accounts have been processed, the callback function is executed once.
  * @param MccApp.​ExecutionResult results
*/
function finalExecution(results)
{

   var offset = 7;

   for (var i = 0; i < results.length; i++) {
    // turn the returned value back into a JavaScript object.
     var resultsObject = JSON.parse(results[i].getReturnValue()); 
     var scriptErrors = '';
     
     // if a return value was returned and an error message found, set it as error
     if(resultsObject && resultsObject.errors !== null){
      scriptErrors =  scriptErrors+resultsObject.errors;
     }    

     // if any other kind of script error occured (overrides the lower prioritized errors!)
     if(results[i].getError() !== null){
       scriptErrors = scriptErrors+results[i].getError();
     }

     if(resultsObject !== null){
       Logger.log("new reactivated: "+resultsObject.activated_count);
       Logger.log("new paused: "+resultsObject.paused_count);
       Logger.log("paused campaigns by script: "+resultsObject.paused_campaigns);
       Logger.log("paused shopping campaigns by script: "+resultsObject.paused_shopping_campaigns);
       Logger.log("index: "+resultsObject.sheet_current_row_index);
       
       // Write number of paused campaigns by script to sheet
       SHEET.getRange('F'+(resultsObject.sheet_current_row_index+offset)).setValue(resultsObject.paused_campaigns+resultsObject.paused_shopping_campaigns);
       
       // Write number of paused campaigns by script to sheet
       SHEET.getRange('F'+(resultsObject.sheet_current_row_index+offset)).setValue(resultsObject.paused_campaigns+resultsObject.paused_shopping_campaigns);
       if(scriptErrors != ''){      
        // Write errors based on an account to sheet
         SHEET.getRange('L'+(resultsObject.sheet_current_row_index+offset)).setValue(scriptErrors); 
       }else{
        // Remove errors from sheet
         SHEET.getRange('L'+(resultsObject.sheet_current_row_index+offset)).setValue('');  
       }
       
       // Send an email if addresses were found and changes were made or an error occured
       if(resultsObject.emails !== null && (resultsObject.activated_count + resultsObject.paused_count > 0) || scriptErrors !== ''){
         if(scriptErrors === ''){ scriptErrors = null; }
         // Generate the eMail-Template
         var emailcontent = generateEmailContent(resultsObject, scriptErrors);
         // Send eMail
         sendEmail(resultsObject.emails, EMAIL_SUBJECT, emailcontent);
       }       
     }else{
         // Write script errors to sheet
         if(scriptErrors !== ''){
             SHEET.getRange('L'+offset).setValue(scriptErrors);
         }
     }
   }
}

/**
  * @desc Generate eMail template
  * @param dataObject dataObj, string errors
  * @return string - generated template
*/
function generateEmailContent(dataObj, errors){
  var strVar="";
      strVar += "<style>";
      strVar += " td,th{";
      strVar += "   padding: 10px;";
      strVar += " }";
      strVar += "<\/style>";
      strVar += "";
      strVar += "<div>";
      strVar += " <p align=\"right\" style=\"text-align:right\"><i><span style=\"font-size:10.0pt;color:#666666\">Powered by crealytics<\/span><\/i><\/p>";
      strVar += "<\/div>";
      strVar += "";
      strVar += "<div style=\"background-color:#3C78D8; padding:5px;\"> ";
      strVar += "<p><span style=\"font-size:18.0pt;font-family:Verdana;color:white;\">Emergency Stop<\/span><\/p>";
      strVar += "<p><span style=\"font-family:Verdana;color:white;\">for "+dataObj.accountname+" ( "+dataObj.accountid+" )<\/span><\/p>";
      strVar += "<\/div>";
      strVar += "<br\/><br\/>";   
      strVar += "";   
      strVar += "<div style=\"padding:5px;font-family:Sans-Serif; background-color:#FF9900;\">";
      strVar += "<span style='color:#ffffff;font-family:Sans-Serif'><b>Following changes were made:<\/b><\/span><br\/>";
      strVar += "<\/div>";
      strVar += "";
      strVar += "<div style=\"background-color:#EEEEEE; color:#000000; padding:5px; font-size:14.0pt; color:#444;\">";
      strVar += "new paused campaigns: "+dataObj.paused_count+"<br/>";
      strVar += "new reactivated campaigns: "+dataObj.activated_count+"<br/>";
      strVar += "paused campaigns by script: "+(dataObj.paused_campaigns+dataObj.paused_shopping_campaigns)+"<br/>";
      strVar += "therefrom shopping campaigns: "+dataObj.paused_shopping_campaigns;
      strVar += "<\/div>";
      strVar += "<br\/><br\/>";   
      strVar += "";   

      // if an error message was handed over
      if(errors !== null){
       strVar += "<div style=\"padding:5px;font-family:Sans-Serif; background-color:red;\">";
       strVar += "<span style='color:#ffffff;font-family:Sans-Serif'><b>errors occurs: "+errors+"<\/b><\/span><br\/>";
       strVar += "<\/div>";
      }

      strVar += "<br\/><br\/>";
      strVar += "";

      return strVar;
}

/**
  * @desc Remove any rows without an account id
  * @param Object[][] data - the rectangular grid of values for this range
  * @return Object[][] - grid of values for this range
*/  
function checkSpreadsheedData(data){
  var cleanedData = [];
  for(var i=0; i<data.length; i++){
    if(data[i][SHEETCOLS.ACCOUNT_IDS].length > 0 && data[i][SHEETCOLS.ACCOUNTNAMES].length > 0){
      cleanedData.push(data[i]);
    }
  }
  return cleanedData;
}

/**
  * @desc Selects a specific column from a range
  * @param Object[][] data - the rectangular grid of values for this range, int colnumber - which column to select
  * @return Object[] - values for this column
*/  
function getColumnFromRange(data, colnumber){
  var column = [];
  for(var i=0; i<data.length; i++){
    column.push(data[i][colnumber]);
  }
  return column;
}

/**
  * @desc Selects a row of a range by an account id
  * @param Object[][] data - the rectangular grid of values for this range, string id
  * @return Object[] - values for this row
*/  
function getRowInRangeById(data, id){
  var ids = getColumnFromRange(data, SHEETCOLS.ACCOUNT_IDS);
  if(ids.indexOf(id) > -1){
    return data[ids.indexOf(id)];
  }
}

/**
  * @desc Checks if the paused_label is attached to a campaign
  * @param AdWordsApp.Campaign campaign
  * @return boolean
*/  
function campaignHasPausedLabel(campaign){
  var hasLabel = false;
  var labelsSelector = campaign.labels();
  var labelIterator = labelsSelector.get();
  while (labelIterator.hasNext()) {
   var label = labelIterator.next();
    if(label.getName() == PAUSED_LABEL){
      hasLabel=true;
    }
  }
  return hasLabel;
}

/**
  * @desc Create a label
*/  
function createLabelIfNotExist(){
  if(!pauseLabelExists(AdWordsApp.labels().get())){    
    AdWordsApp.createLabel(PAUSED_LABEL);
    Logger.log("Label "+PAUSED_LABEL+" created");
  }
}

/**
  * @desc Checks if the label exists
  * @param AdWordsApp.LabelIterator labelsIter
  * @return boolean
*/  
function pauseLabelExists(labelsIter) {
  var hasLabel = false;
  var labelIterator = labelsIter;

  while (labelIterator.hasNext()) {
    var label = labelIterator.next();
    if(label.getName() == PAUSED_LABEL){
      hasLabel=true;
    }
  }
  return hasLabel;
}

/**
  * @desc Pauses a campaign
  * @param AdWordsApp.Campaign campaign
  * @return boolean - true if succcessfull
*/ 
function pauseCampaign(campaign){
  if(campaign.isEnabled()){      
          if(! campaignHasPausedLabel(campaign)){
            campaign.pause();
            campaign.applyLabel(PAUSED_LABEL);
            return true;
          }
  }
  return false;
}

/**
  * @desc Reactivates a campaign
  * @param AdWordsApp.Campaign campaign
  * @return boolean - true if succcessfull
*/ 
function activateCampaign(campaign){
 if(campaign.isPaused()){
          if(campaignHasPausedLabel(campaign)){
            campaign.enable();
            campaign.removeLabel(PAUSED_LABEL);
            return true;
          }
        }
  return false;
}
