// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet
// View the README at the original source repo for installation instructions. 
// Don't forget to add the required slack information below.

// Source: https://github.com/markfguerra/google-forms-to-slack
// Adapted by Andy Meneely

/////////////////////////
// Begin customization //
/////////////////////////

// Alter this to match the incoming webhook url provided by Slack
var slackIncomingWebhookUrl = '';

// Include # for public channels, omit it for private channels
var postChannel = '#kiosk';
var postIcon = ":kiosky:"; //note that I made a custom emoji in the workspace... for funsies
var postUser = "Kiosky McBotFace";
var postColor = "#0000DD";

var messageFallback = "Someone is here for you!";
var messagePretext = "";

///////////////////////
// End customization //
///////////////////////

// In the Script Editor, run initialize() at least once to make your code execute on form submit
function initialize() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("submitValuesToSlack")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

// Running the code in initialize() will cause this function to be triggered this on every Form Submit
function submitValuesToSlack(e) {
  // Test code. uncomment to debug in Google Script editor
  // if this is called in the Google Script editor it SHOULD send messages to Slack
  if (typeof e === "undefined") {
    e = {
      values: [
      'date here',
      "Debuggy",
      "Debuggyson",
      "Dr. Meneely, Undergrad Director",
      'In Person',
      '',
      '',
      '',
      '',
      '',
      '',
      'In Person'
      ]
    };  
    messagePretext = "Debugging our Sheets to Slack integration";
  }

  var attachments = constructAttachments(e.values);
  //e.values corresponds to the columns in the spreadsheet - yours may vary
  const firstName = e.values[1];
  const lastName = e.values[2];
  const toSee = e.values[3];
  const modality = (e.values[11] == 'In Person') ? 'in person' : 'via zoom';

  // UPDATE THESE WITH YOUR USERNAMES
  const slackUsernames = {
    "Dr. Meneely, Undergrad Director" : 'axmvse',
    'Megan Lehman, Academic Advisor' : 'melics',
    'Carrie Koneski, Academic Advisor' : 'caksma',
    'Sarah Mittiga, Academic Advisor' : 'sxmrgr',
    'Dr. Hawker, SE-MS Program Director' : 'jshvse',
    'Dr. Sharma, Dept Chair' : 'nxsvse',
  }  

  var payload = {
    "channel": postChannel,
    "text" : `${firstName} ${lastName} is here to see @${slackUsernames[toSee]} ${modality}`,
    "username": postUser,
    "icon_emoji": postIcon,
    "link_names": 1,
    // we commented this out to keep the messages short, but it might be handy
    // "attachments": attachments
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}

// Creates Slack message attachments which contain the data from the Google Form
// submission, which is passed in as a parameter
// https://api.slack.com/docs/message-attachments
var constructAttachments = function(values) {
  var fields = makeFields(values);
  var attachments = [{
    "fallback" : messageFallback,
    "pretext" : messagePretext,
    "mrkdwn_in" : ["pretext"],
    "color" : postColor,
    "fields" : fields
  }]

  return attachments;
}

// Creates an array of Slack fields containing the questions and answers
var makeFields = function(values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    var val = values[i];
    fields.push(makeField(colName, val));
  }

  return fields;
}

// Creates a Slack field for your message
// https://api.slack.com/docs/message-attachments#fields
var makeField = function(question, answer) {
  var field = {
    "title" : question,
    "value" : answer,
    "short" : false
  };
  return field;
}

// Extracts the column names from the first row of the spreadsheet
var getColumnNames = function() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get the header row using A1 notation
  var headerRow = sheet.getRange("1:1");

  // Extract the values from it
  var headerRowValues = headerRow.getValues()[0];

  return headerRowValues;
}
