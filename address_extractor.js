function all() { 
  ss = SpreadsheetApp.getActiveSpreadsheet();
  error = false;
  GetAddresses();
  if(error == false){
  Duplicates();
  }
}

function GetAddresses ()
{
  // message to tell user the process is going on 
  var userInput = ss.getActiveSheet().getRange(1,1,1,9).getCell(1,9).getValues();
  console.log("✅userInput passed into getaddress: "+userInput)
  ss.toast("Fetching data......")

  // Label to search  
  var labelName ="Inbox"
  var sheetName = "Label: " + labelName;
  var sheet = ss.getSheetByName (sheetName) || ss.insertSheet (sheetName, ss.getSheets().length);
  
  // get all messageData in an array
  var addressesOnly = [];
  var messageData = [];

  // allow users to input a date: retrieve email after this date 
  var date = "";
  var selectedDate = sheet.getRange(1,1,1,6).getCell(1,6).getValue();
  if(selectedDate == ""){
    date = new Date();
  }
  else{
    date = selectedDate
  }
  var td = Utilities.formatDate(date, 'GMT+08:00', "yyyy/MM/dd");
  var queryString = `label: ${labelName} after: ${td}`;
  var threads = GmailApp.search(queryString);

  //if number of threads exceeded 500,loop it to get more threads.
  if(threads.length==500){
    var nextDate = threads.slice(-1)[0].getLastMessageDate();
    while(nextDate >= selectedDate){
        td = Utilities.formatDate(nextDate, 'GMT+08:00', "yyyy/MM/dd");
        queryString = `label: ${labelName} before: ${td}`;
        threads = threads.concat(GmailApp.search(queryString));
        nextDate = threads.slice(-1)[0].getLastMessageDate();
    }
    console.log("out of the loop, resulting thread length is" + threads.length)
  }
  for(var i = 0; i < threads.length; i++){
    var thread = threads[i];
    var msg = thread.getMessages()[0]; //get first message
    
    //get only the email address from the getFrom()
    var mailFrom = msg.getFrom();    
    var addressonly = mailFrom.replace(/^.+<([^>]+)>$/, "$1")
    var userEmail = Session.getActiveUser().getEmail();

    //check if the first msg is sent from repetively received email address. if yes, skip to next thread    
    //if is filtered, skip the thread.
    if (userInput.indexOf(addressonly) >= 0) {
      console.log("skip this thread: " + addressonly);
      continue;
    }

    //check if is sent from internal people, if yes, read next email to get from
    j=0
    if(addressonly.match(/@cheungwofood.com.sg$/) || addressonly.match(userEmail) || addressonly.match('hello@heychips.com'))
    {
      console.log("this is from my internal email: "+ mailFrom)
      var repeat = true;
      while(repeat){
        try {
        j=j+1
        mailFrom = thread.getMessages()[j].getFrom()
        addressonly = mailFrom.replace(/^.+<([^>]+)>$/, "$1")
        }
        catch(err) {
          console.log("mailfrom is undefined,skip to next thread.")
          break;
          repeat = false;
          }
        }
      if(repeat){
        continue;
      }
    }
    var mailDate = thread.getLastMessageDate();
    //var mailDate = msg.getDate ();
    // mailFrom format may be either one of these:
    // name@domain.com
    // any text <name@domain.com>
    // "any text" <name@domain.com>
    var name = "";
    var email = "";
    var matches = mailFrom.match (/\s*"?([^"]*)"?\s+<(.+)>/);
    if (matches)
    {
      email = matches[2];
      name = matches[1]
    }
    else
    {
      email = mailFrom;
      //name = mailFrom.match(/^[a-zA-Z]+?(?=@)/)
    }

    // Check if (and where) we have this already
    var index = addressesOnly.indexOf (mailFrom);
    if (index > -1)
    {
      continue;
    }
    // Add the data
    addressesOnly.push (mailFrom);
    messageData.push ([name, email, mailDate]);
  }
   // Add data to corresponding sheet
  try{
    sheet.getRange (sheet.getLastRow() + 1, 1, messageData.length, 3).setValues (messageData);
    }
  catch(err){
      ss.toast("Uh oh,no data was found. Check your date format", "⚠️ Error",15);
      error = true;
  }
}


//
// Adds a menu to easily call the script
//
function onOpen ()
{ 
    ss = SpreadsheetApp.getActiveSpreadsheet(); 
    sheet = ss.getActiveSheet();
    SpreadsheetApp.getUi().createMenu("📧 EmailAddress extrator")
    .addItem("Extract Addresses", "all")
    .addToUi();

    var date = new Date();
    var data = Utilities.formatDate(date, 'GMT+08:00', "yyyy/MM/dd")
    const preset = [["Sender","Email Address","Date of Last Msg"," ",
    "Retrieving email addresses after this date:(yyyy/mm/dd):",date," ","Filter Email List(in comma separated form):"]]
    sheet.getRange(1,1,1,8).setValues(preset);

}

function Duplicates()
{ 
  var sheet = ss.getActiveSheet();
  const data = sheet.getRange(2,1,sheet.getLastRow(),3).sort({column: 3,ascending: false}).getValues();
  let uniqueSs = [];
  let uniqueRows = [];
  data.forEach((row) => {
    let emailAddress = row[0];
    //if the address is not in the unique lists, add it.
    if (!uniqueSs.includes(emailAddress)) {
    uniqueSs.push(emailAddress);
    uniqueRows.push(row);
    }
  });
  //clear the sheet with the newly input
  sheet.getRange(2,1,sheet.getLastRow(),3).clear()
  //input the unique rows
  sheet.getRange(2, 1, uniqueRows.length, 3).setValues(uniqueRows);

  ss.toast("Extraction process is completed! =)")
  }


