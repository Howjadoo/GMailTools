// Gmail Emails to Google Drive PDFs
// Author: Daniel Shanklin
// ... with some code help from the internet
//
// 1.) Label any GMail conversation as "ConvertToPDF"
// 2.) Run this script (can be automated to run in the background via trigger)
// 3.) The selected emails are converted to PDF and saved to your "My GMail Attachments" folder in Google Drive
// 4.) The label is changed from "ConvertToPDF" to "ConvertToPDF_Done"
// 5.) Runs five at a time.  If you set up a trigger to run the script every 60 seconds, you can process up to 300 emails per hour

function saveGmailAsPDF() {

  var gmailLabels  = "ConvertToPDF";
  // var driveFolder  = "My Gmail";
  var driveFolder  = "My Gmail Attachments";
  var MyTimeZone = "US/Eastern";

  var threads = GmailApp.search("in:" + gmailLabels, 0, 5);

  console.log('Number to process: ' + threads.length);
  if (threads.length > 0) {

    /* Google Drive folder where the Files would be saved */
    // var folders = DriveApp.getFoldersByName(driveFolder);
    // var folder = folders.hasNext() ?
    //    folders.next() : DriveApp.createFolder(driveFolder);
    // var folder = getDriveFolder(driveFolder);

    /* Gmail Label that contains the queue */
    var label = GmailApp.getUserLabelByName(gmailLabels) ?
        GmailApp.getUserLabelByName(gmailLabels) : GmailApp.createLabel(driveFolder);
    var labelDone = GmailApp.getUserLabelByName("ConvertToPDF_Done");
    var labelInProcess = GmailApp.getUserLabelByName("ConvertToPDF_InProcess");

    console.log('Number to process: ' + threads.length);
    for (var t=0; t<threads.length; t++) {

      // threads[t].removeLabel(label);
      threads[t].addLabel(labelInProcess);
      var msgs = threads[t].getMessages();

      var html = "";
      var attachments = [];
      
      var fromEmail = msgs[0].getFrom().replace(/<|>|"/g, '');
      var folder = getDriveFolder(driveFolder + "/" + Utilities.formatDate(msgs[0].getDate(), MyTimeZone, "yyyy") + "/" + fromEmail); 

      var subject = threads[t].getFirstMessageSubject();
      
      var firstDateInThread = Utilities.formatDate(msgs[0].getDate(), MyTimeZone, "yyyy-MM-dd");

      /* Append all the threads in a message in an HTML document */
      for (var m=0; m<msgs.length; m++) {

        var msg = msgs[m];
        var curDate = Utilities.formatDate(msg.getDate(), MyTimeZone, "yyyy-MM-dd");

        html += "From: " + msg.getFrom() + "<br />";
        html += "To: " + msg.getTo() + "<br />";
        html += "Date: " + msg.getDate() + "<br />";
        html += "Subject: " + msg.getSubject() + "<br />";
        html += "<hr />";
        html += msg.getBody().replace(/<img[^>]*>/g,"");
        html += "<hr />";

        var atts = msg.getAttachments();
        for (var a=0; a<atts.length; a++) {
          attachments.push(atts[a]);
        }
      }

      /* Save the attachment files and create links in the document's footer */
      if (attachments.length > 0) {
        var footer = "<strong>Attachments:</strong><ul>";
        for (var z=0; z<attachments.length; z++) {
          var file = folder.createFile(attachments[z]);
          footer += "<li><a href='" + file.getUrl() + "'>" + file.getName() + "</a></li>";
        }
        html += footer + "</ul>";
      }

      /* Conver the Email Thread into a PDF File */
      var tempFile = DriveApp.createFile("temp.html", html, "text/html");
      folder.createFile(tempFile.getAs("application/pdf")).setName(firstDateInThread + " - " + subject + ".pdf");
      tempFile.setTrashed(true);
      
      // Only remove label when completely done!
      threads[t].addLabel(labelDone);
      threads[t].removeLabel(labelInProcess);
      threads[t].removeLabel(label);
      

    }
  }
}

function getDriveFolder(path) {

  var name, folder, search, fullpath;

  // Remove extra slashes and trim the path
  fullpath = path.replace(/^\/*|\/*$/g, '').replace(/^\s*|\s*$/g, '').split("/");

  // Always start with the main Drive folder
  folder = DriveApp.getRootFolder();

  for (var subfolder in fullpath) {

    name = fullpath[subfolder];
    search = folder.getFoldersByName(name);

    // If folder does not exit, create it in the current level
    folder = search.hasNext() ? search.next() : folder.createFolder(name);

  }

  return folder;

}
