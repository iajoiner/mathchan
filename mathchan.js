//Clean sender
function cleanSender(sender) {
  var pos = sender.search("<");
  return sender.slice(0, pos - 1);
}
//Num to string, padding included
function numToString(num) {
  if(num < 10) {
    return '0' + num.toString();
  }
  else {
    return num.toString();
  }
}
//date to time stamp
function genTimeStamp(date) {
  if (!(date instanceof Date)) {
    console.log('Type Error!');
    return undefined;
  }
  var hours = numToString(date.getHours());
  var minutes = numToString(date.getMinutes());
  var seconds = numToString(date.getSeconds());
  var month = numToString(date.getMonth() + 1);
  var day = numToString(date.getDate());
  var year = numToString(date.getFullYear());
  var timeStamp = month+day+year+hours+minutes+seconds;
  return timeStamp;
}
function spaceEscape(path) {
  return path.split(' ').join('_');
}
//Use file extension to determine whether a file is an image
function isImage(file) {
  return file.endsWith(".jpg") || file.endsWith(".png") || file.endsWith(".JPG") || file.endsWith(".PNG") || file.endsWith(".jpeg") || file.endsWith(".JPEG") || file.endsWith(".jpeg") || file.endsWith(".JPEG") || file.endsWith(".tiff") || file.endsWith(".TIFF");
}
function run() {
  var mail = Application('Mail');
  var finder = Application('Finder');
  var app = Application.currentApplication();
  var keyword = "HW";
  var topFolderPath = "/Users/CatLover/Documents/HWBox";
  var topFolder = finder.startupDisk.folders.byName("Users").folders.byName("CatLover").folders.byName("Documents").folders.byName("HWBox");
  mail.includeStandardAdditions = true;
  finder.includeStandardAdditions = true;
  app.includeStandardAdditions = true;
  var messages = mail.inbox.messages;
  var messagesLength = messages.length;
  for(let i = 0; i < messagesLength; i++) {
    let message = messages[i];
    if (message.subject().includes(keyword) && message.mailAttachments().length != 0) {
	  var sender = spaceEscape(cleanSender(message.sender()));//Name only
	  var timeStamp = genTimeStamp(message.dateReceived());//Folder name
	  var attachments = message.mailAttachments();
	  var individualPath = topFolderPath + '/' + sender;
	  if (!finder.exists(Path(individualPath))) {
	  //No folder!
	    finder.make({new: "folder", at: Path(topFolderPath), withProperties:{name: sender}});
	  }
	  var messagePath = individualPath + '/' + timeStamp;
	  if (!finder.exists(Path(messagePath))) {
	  //No folder!
	    finder.make({new: "folder", at: Path(individualPath), withProperties:{name: timeStamp}});
		var command = "/Library/Frameworks/Python.framework/Versions/3.7/bin/img2pdf ";
		var attachmentsLength = attachments.length;
		var hasPics = false;
		for(let j = 0; j < attachmentsLength; j++) {
	      let attachment = attachments[j];
	      var fileName = spaceEscape(attachment.name());
		  var filePath = messagePath + '/' + fileName;
		  mail.save(attachment, {in: Path(filePath)});
		  if (isImage(fileName)) {
		    command = command + filePath + ' ';
			hasPics = true;
		  }
	    }
		if (hasPics) {
		  //Pics
		  var pdfName = spaceEscape(sender + timeStamp);
		  command = command + '-o ' + messagePath + '/' + pdfName + '.pdf';
		  //console.log(command);
		  app.doShellScript(command);
		}
	  }
	  else {
	    continue;//Non-spammers aren't going to send two emails at the same time down to the same second. No need to process an already processed email.
	  }
	}
  }
}
