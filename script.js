/**
ALL YOU NEED TO DO IS REPLACE docID
 */
var docId = '<DOC ID>';

/**
Leave these alone
 */
var ssheet = SpreadsheetApp.openById(
  '<SHEET ID>'
);
var doc = DocumentApp.openById(docId);
var docBody = doc.getBody();
var docName = doc.getName()
var docURL = doc.getUrl()

var currentDateinMonthDayYear = currentDateinMonthDayYear()

var industries = ['landscap', 'janitor'];

var replacedLinks = {};

function currentDateinMonthDayYear() {
  const currentDate = new Date();
  const formattedDate = (currentDate.getMonth() + 1).toString().padStart(2, '0') + '/' + currentDate.getDate().toString().padStart(2, '0') + '/' + currentDate.getFullYear();
  return formattedDate
}

/**
Script Wrapper
 */
function runScript() {
  // Returns string from industries variable
  let [industryObject, industryString] = whichIndustryIsContentAbout();

  // Create JSON Object From Sheet Data
  obj = createJSONFromSheet(industryObject, industryString);

  // Loop through Created JSON Object
  Object.entries(obj).forEach(([key, value]) =>
    replaceTextWithLink(value, key)
  );

  if (ScriptApp.getUserTriggers(doc).length == 0) {
    createDocOpenTrigger()
  }

  // Log replaced links in Console
  var prettyPrintreplacedLinks = JSON.stringify(replacedLinks).replace(/\,/g, '\n')
  var replacedLinkCount = Object.values(replacedLinks).reduce((acc, val) => acc + val.length, 0)
  var scriptUser = Session.getActiveUser().getEmail()

  Logger.log(docName)
  Logger.log(docURL)
  Logger.log(prettyPrintreplacedLinks);
  Logger.log(`${replacedLinkCount} links added`);
  Logger.log(`Script ran by: ${scriptUser}`);

  if (docName.includes('Copy of') && scriptUser.includes('fjones')) Logger.log('*TESTING')

  if (!scriptUser.includes('fjones')) {

    // Log replaced links in Google Sheet Script Log
    addReplacedLinksToScriptLog(replacedLinks, scriptUser, replacedLinkCount)

    // Send email
    sendEmailMessage(prettyPrintreplacedLinks, replacedLinkCount, scriptUser)
  }
}

function addReplacedLinksToScriptLog(replacedLinks, scriptUser, replacedLinkCount) {
  for (let x = 0; x < Object.values(replacedLinks).length; x++) {
    const url = Object.keys(replacedLinks)[x];
    const anchorTextArray = Object.values(replacedLinks)[x];
    for (let x = 0; x < anchorTextArray.length; x++) {
      const anchor = anchorTextArray[x];
      ssheet.getSheetByName('Script Log').appendRow([currentDateinMonthDayYear,docName, docURL, scriptUser, replacedLinkCount, url, anchor])
      ssheet.getSheetByName('Script Log').getRange(ssheet.getSheetByName('Script Log').getLastRow(), 1, 1, ssheet.getSheetByName('Script Log').getLastColumn()).clearFormat();
    }
  }
  removeDuplicatesFromScriptLog()
}

function removeDuplicatesFromScriptLog() {
  // Remove rows which have duplicate values in both columns B and D in Google Sheet Script Log.
  removeExtraRows()
  var range = ssheet.getSheetByName('Script Log').getRange("A2:F");
  range.removeDuplicates([2, 5, 6]);
  removeExtraRows()
}

//Remove All Extra Rows in Google Sheet Script Log
function removeExtraRows() {
  var maxRows = ssheet.getSheetByName('Script Log').getMaxRows();
  var lastRow = ssheet.getSheetByName('Script Log').getLastRow();
  if (maxRows - lastRow != 0) {
    ssheet.getSheetByName('Script Log').deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function sendEmailMessage(prettyPrintreplacedLinks, replacedLinkCount, scriptUser) {
  var message = {
    to: "fjones@youraspire.com",
    subject: `Add Internal Links Ran on ${docName}`,
    body: `The add internal links script was just ran on ${docURL}
    
${replacedLinkCount} links added

${prettyPrintreplacedLinks}
    
Script ran by: ${scriptUser}`,
    name: "Add Internal Links Apps Script"
  }
  MailApp.sendEmail(message);
}

function createDocOpenTrigger() {
  ScriptApp.newTrigger('addToDocUI')
    .forDocument(docId)
    .onOpen()
    .create()
}

function addToDocUI() {
  DocumentApp.getUi()
    .createMenu('User Actions')
    .addItem('Remove Highlights', 'deleteHighlight')
    .addToUi();
}

function deleteHighlight() {
  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  var body = DocumentApp.getActiveDocument().getBody()
  var paragraphs = body.getParagraphs();
  paragraphs.forEach((para) => {
    Logger.log(para)
    para.setAttributes(highlightStyle)
  })
}

function createJSONFromSheet(industryObj, industryStr) {
  // Input sheet name
  let sheet = ssheet.getSheetByName('Internal Links');

  // Range to iterate over
  let values = sheet.getRange('D2:E').getValues();

  // Logger.log(values)
  var jsonObject = {};


  // Loops through sheet rows
  for (var i = 1; i < values.length; i++) {
    var anchor = values[i][0];
    var url = values[i][1];
    if (url in jsonObject) {
      jsonObject[url].push(anchor);
    } else {
      jsonObject[url] = [anchor];
    }
  }

  // Logger.log(JSON.stringify(jsonObject));

  // Everything that doesn't contain any industry string is a industry agnostic url
  var industryAgnosticObj = createJSONObjwIndustryAgnosticURLS(jsonObject, industryObj)

  // Everything that only contains the industry string; see whichIndustryIsContentAbout() function to understand industry string
  var industrySpecificObj = createJSONObjwIndustrySpecificURLS(jsonObject, industryObj, industryStr)

  var combinedFilteredJSONObj = combineTwoJSONObj(industryAgnosticObj, industrySpecificObj)
  return combinedFilteredJSONObj
}

// This combines the Industry Agnostic URLs with the Industry Specific URLS
function combineTwoJSONObj(jsonObj1, jsonObj2) {
  return {
    ...jsonObj1,
    ...jsonObj2
  };
}

function createJSONObjwIndustryAgnosticURLS(obj, industryObj) {
  // Convert `obj` to a key/value array
  const asArray = Object.entries(obj);

  let arrays = [];

  for (let key in industryObj) {
    arrays = arrays.concat(industryObj[key]);
  }

  let pattern = arrays.join("|");
  let regexFromMyArray = new RegExp(pattern, 'gi');

  const filtered = asArray.filter(([key, value]) => !key.match(regexFromMyArray));

  // Convert the key/value array back to an object:
  const industryAgnosticURLs = Object.fromEntries(filtered);

  return industryAgnosticURLs
}

function createJSONObjwIndustrySpecificURLS(obj, industryObj, industryStr) {
  // Convert `obj` to a key/value array
  const asArray = Object.entries(obj);

  let pattern = industryObj[industryStr].join("|");
  let regexFromMyArray = new RegExp(pattern, 'gi');

  const filtered = asArray.filter(([key, value]) => key.match(regexFromMyArray));

  // Convert the key/value array back to an object:
  const industrySpecificURLsObj = Object.fromEntries(filtered);

  return industrySpecificURLsObj
}

function replaceTextWithLink(searchText, linkUrl) {
  // Look for longer anchor text first
  var searchTextSorted = searchText.sort((a, b) => b.length - a.length);

  // successfulSearches variable is used to keep track of the number of successful searches performed during the loop. 
  // IF reinitialized every time the loop runs it will allow for the possibility of MULTIPLE links being added with DIFFERENT anchor texts linking to the same page.
  // IF initialized before the loop it will result it only a SINGLE link with a SINGLE anchor text being added to the Google Doc if it finds the anchor text phrase.
  var successfulSearches = 0;

  // Loop through the searchText array
  for (let i = 0; i < searchTextSorted.length; i++) {
    // Search for the searchText in the document


    // For logging purposes
    var trimmedSearchTextForReplacedLinks = `${searchTextSorted[i].trim()}`;

    // Creates regex pattern to search document body
    var trimmedSearchText = `${searchTextSorted[i].trim()}[a-z'’]{0,}`;

    // Actual search result from first search
    var searchResult = docBody.findText(trimmedSearchText);

    // Loops over all search results.
    // successfulSearches in if condition controls how many instances of the SAME anchor text will be linked to if multiple matches are found. 
    // default is 1. if == 0 that means 1 link. if == 1 that means 2.
    while (searchResult !== null && successfulSearches == 0) {
      // Iterates successfulSearches
      successfulSearches++;

      // Utility to get parent element
      var parentElement = searchResult.getElement().getParent();

      // Utlity to get parent elements parent
      var parentElementsParent = parentElement.getParent();

      // Utlity to get BOLD attribute of searchResult element
      var isBold = searchResult.getElement().getAttributes().BOLD;

      // Utility to get HEADING attribute of parentElement, also converts the heading to a string
      var parentElementHeadingIsNormal = String(parentElement.getAttributes().HEADING);

      // When Content Outline, this regex prevents adding links in MetaData tables and other places where links shouldn't be added
      let regex =
        /Page title|Meta description|Blog Title|URL|Airtable|Wireframe|SEO|Product|Notes|Kicker|Heading|button|Featured|Button|Question|Relevant product(s)|Title|Body \(220 characters\)|Q[0-9]|Client:|Assignment:|Headline/gm;
      // check if table row and run regex test using string above
      if (
        parentElementsParent.getParent() == 'TableRow' &&
        regex.test(parentElementsParent.getParent().getText())
      ) {
        break;
      } else {
        // Replace searched and matched text
        if (
          (parentElement == 'Paragraph' || parentElement == 'ListItem') &&
          isBold == null &&
          parentElementHeadingIsNormal == 'NORMAL'
        ) {
          var startIndex = searchResult.getStartOffset();
          var endIndex = searchResult.getEndOffsetInclusive();
          var text = searchResult.getElement().asText();
          var highlightStyle = {};
          highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFF00';
          text.setAttributes(highlightStyle);
          text.setLinkUrl(startIndex, endIndex, linkUrl);
          searchResult = docBody.findText(trimmedSearchText, searchResult);

          // Utlity for logging
          if (linkUrl in replacedLinks) {
            replacedLinks[linkUrl].push(trimmedSearchTextForReplacedLinks);
          } else {
            replacedLinks[linkUrl] = [trimmedSearchTextForReplacedLinks];
          }
        }
      }
    }
  }
}

function whichIndustryIsContentAbout() {
  var text = docBody.getText();
  industryWordCount = {};
  for (let i = 0; i < industries.length; i++) {
    const industry = industries[i];
    let count = 0;
    const match = text.match(new RegExp(industry, 'gmi'));
    // Logger.log(industry)
    if (match) {
      count = match.length;
    }
    industryWordCount[industry] = count;
  }

  let industryString = Object.keys(industryWordCount).reduce((a, b) =>
    industryWordCount[a] > industryWordCount[b] ? a : b
  );

  const industryDictionary = {};
  industries.forEach(val => industryDictionary[val] = [val]);
  // Add related words to search for in urls
  industryDictionary['janitor'].push('cleaning')

  return [industryDictionary, industryString]

}
