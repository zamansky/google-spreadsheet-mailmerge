// Mailmerge
//
// Have a sheet named "Document" with in the following format:
// Cell        Value
// A1          Subject
// A2          The actual subject
// C1          Message
// C2          The actual message
//
// That is, A1 and C1 are the words "Subject" and "Message"
// and A2 and C2 the actual values that you will use for subject and message
//
// Embed newlines in the message (CTL-Enter)
// Things demarkated by % will be substituted from the other sheeet, so %First% will get values from the column First in the other sheeet etc (see below).
//
// Have another sheet named "Recipients" where the first row represent field names and the rest of each column are the values.
//
// For example if you have First in cell A1, Last in cell A2 and Email in cell A3, in your message occurances of %First% will be replaced by first name.
//
// Column  C MUST be the Email column but that's easily changable in the script
//
//
// This script is made available as is. It is fragile but works for me. Use at your own risk.
 
function onOpen( ){
// This line calls the SpreadsheetApp and gets its UI   
// Or DocumentApp or FormApp.
  var ui = SpreadsheetApp.getUi();
 
//These lines create the menu items and 
// tie them to functions we will write in Apps Script
  
 ui.createMenu('Custom')
      .addItem('Mailmerge', 'mailmerge')
      .addToUi();
}

function mailmerge(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Document");
  var r = sheet.getDataRange();
  var subject = r.getValues()[1][0];
  var raw_email = r.getValues()[1][2];
  
  // Logger.log(email);
                    
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recipients");
  
  var r  = sheet.getDataRange();
  
  var subs = {};
  
  values = r.getValues();
  var keys = values[0];
  var vals = values[1];
  
  for (i in keys){
   subs["%"+keys[i]+"%"] = i; 
  }
  
  //Logger.log(subs);
    
  
  // now loop over the rest of the lines to make each message
  
  for (var i in values){
    if (i==0) continue;
    var email = raw_email;
    var emailaddress=values[i][2];
    for (var j in values[0]){
     var k = "%"+values[0][j]+"%" 
     email = email.replace(new RegExp(k,'g'),values[i][j])
     
    }
    //Logger.log(emailaddress+":"+subject+":"+email);
    MailApp.sendEmail(emailaddress,subject,email,{noReply:true});
  }
    

  
}

