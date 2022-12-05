
// Including a menu where I can run the spreadsheet functions
function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addItem('Sync Data', 'importrealtimeattachments',
      'Sync Run Data', 'import_reportrun')
      .addToUi();

 ui.createMenu("test").addItem().addItem
}


function importrealtimeattachments() {

  

 // Email threads containing csv attachments for real-time SLA reports

  // The emails are divided into two categories: hourly and daily. The Daily Data is used to calculate an accurate average handling time from      Freshdesk.

 emailshour = ["Conversation_Team_Hour_Realtime","Social Team_Hour_Realtime","Email Team_Realtime","Retention Team_Realtime","OB Team_Realtime"];
 emailsday = ["Conversation Team_Day - RealTime","Social Team by Day - RealTime"]




  //  Going through the names of the email threads mentioned above in a loop
  for(let j = 0; j < emailshour.length; j++) {


   // Finding The desired thread
   var threads = GmailApp.search(emailshour[j]);

   // Finding The thread's most recently received message
   messagcount = threads[0].getMessageCount()
   var message = threads[0].getMessages()[messagcount-1];

   // obtaining attachments in the message
   var attachment = message.getAttachments()[0];

    // Setting the Attachment Content Type based on the File Extension
   attachment.setContentTypeFromExtension();

   // Checking to see if the content is text or Csv
   if (attachment.getContentType() === "text/csv") {
      // Obtaining Active Spreadsheet
     var realtimesheet = SpreadsheetApp.getActiveSpreadsheet();

     // Obtaining The target worksheet by name 
     var sheet = realtimesheet.getSheetByName("Syncdata");

     // The CSV Content is Parsed and Converted to a String 
     var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");

     //  Specifying the desired range for pasting the attachment data into the spreadsheet
     sheet.getRange(1, 1+(j*9), csvData.length, csvData[0].length).setValues(csvData)

    } 

  }



  // Repeating the same process for the Daily data
  for(let i = 0; i < emailsday.length; i++) {



   var threads = GmailApp.search(emailsday[i]);
   messagcount = threads[0].getMessageCount()
   Logger.log(emailsday[i])
   Logger.log(threads[0].getId())
   var message = threads[0].getMessages()[messagcount - 1];
   var attachment = message.getAttachments()[0];

   attachment.setContentTypeFromExtension();

   if (attachment.getContentType() === "text/csv") {

      
     csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
     sheet.getRange(28, 2+(i*9), csvData.length, csvData[0].length).setValues(csvData)

    }

  }


}  

// Function that sleeps for have a minute before clearing the content of the sheet
function clearcontent(sheet)
{
  Utilities.sleep(300000)
   sheet.clearcontent
}








