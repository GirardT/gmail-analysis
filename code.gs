// Define global variables
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Analysis");
var domainCounts = [];
var logString = "{";
var emailLimit = 1000;

// Function to create menu in Google Sheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Count Senders')
      .addItem('Get Search Patterns', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('List Emails with Count', 'menuItem2'))
      .addToUi();
}

// Function to fetch emails and get search patterns
function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.
    listEmailsBySenderDomain();
    getSearchPatterns();
}

// Function to fetch emails and list emails with counts
function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.
    listEmailsBySenderDomain();
    listEmailsWithCount()
}

// Function to fetch emails and count sender domains
function listEmailsBySenderDomain() {
  var senderDomains = {};
  var emailCount = 0;

  // Fetch emails from the inbox
  var threads = GmailApp.getInboxThreads();

  // Iterate through each thread
  for (var i = 0; i < threads.length && emailCount < emailLimit; i++) {
    var thread = threads[i];

    // Fetch messages from the thread
    var messages = thread.getMessages();

    // Iterate through each message in the thread
    for (var j = 0; j < messages.length && emailCount < emailLimit; j++) {
      var message = messages[j];
      emailCount++;

      // Get the sender's email address
      var senderEmail = message.getFrom();

      // Extract domain from sender's email
      var domain = extractDomain(senderEmail);

      // Update senderDomains object with domain and count
      if (domain in senderDomains) {
        senderDomains[domain]++;
      } else {
        senderDomains[domain] = 1;
      }
    }
  }

  // If sheet doesn't exist, create it
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Email Analysis");
  }

  // Convert senderDomains object to array of objects [{domain: count}, ...]
  for (var domain in senderDomains) {
    domainCounts.push({ domain: domain, count: senderDomains[domain] });
  }

  // Sort domainCounts array by count in descending order
  domainCounts.sort(function(a, b) {
    return b.count - a.count;
  });

  // Construct logString for search patterns
  domainCounts.forEach(function(domainObj) {
    logString += "from:" + domainObj.domain + " ";
  });
  logString += "}";
}

// Function to get search patterns
function getSearchPatterns() {
  // Split log string into chunks of up to 1000 characters each
  var chunkSize = 1000;
  for (var i = 0; i < logString.length; i += chunkSize) {
    var chunk = logString.substring(i, i + chunkSize);
    sheet.appendRow([chunk])
  }
}

// Function to list sender domains and their counts
function listEmailsWithCount() {
  // Add header row
  sheet.appendRow(["Domain", "Email Count"]);
  // Write data to the sheet
  domainCounts.forEach(function(domainObj) {
    sheet.appendRow([domainObj.domain, domainObj.count]);
  });
}

// Function to extract domain from an email address
function extractDomain(email) {
  // Extracting email address from the sender field
  var emailRegex = /<([^>]+)>/;
  var matches = email.match(emailRegex);
  if (matches && matches.length > 1) {
    email = matches[1];
  }

  // Extracting domain from the email address
  var atIndex = email.indexOf("@");
  var domain = email.slice(atIndex + 1);

  // Remove any non-alphanumeric characters (including trailing ">")
  domain = domain.replace(/[^a-zA-Z0-9.-]/g, '');
  return domain;
}
