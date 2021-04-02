


function sendBirthdayMail() {
 
 // return last row for
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

// This logs the value in the very last cell of this sheet


  // before html bodies i used that variables, now they are dummy but still.
  var subject2= "Reminder";
  var subject = "Happy Birthday";
  var birthdaymessage = "Happy Birthday to you";
  var monthmessage="One Month left ";
  var onedaymessage="One day left "

  // html bodies

  
  var reminder1 = HtmlService.createHtmlOutputFromFile('reminderMail').getContent();
  var reminder2 = HtmlService.createHtmlOutputFromFile('reminderMail1').getContent();
  var greenTeen = HtmlService.createHtmlOutputFromFile('GreenMailTeen').getContent();
  var greenAdult = HtmlService.createHtmlOutputFromFile('GreenMailAdult').getContent();
  var blueAdult = HtmlService.createHtmlOutputFromFile('BlueMailAdult').getContent();
  var blueTeen = HtmlService.createHtmlOutputFromFile('BlueMailTeen').getContent();
  var standart = HtmlService.createHtmlOutputFromFile('standart').getContent(); 
  var standart1 = HtmlService.createHtmlOutputFromFile('standart1').getContent();
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;
  
  // datarange to fetch data correctly
  
  var dataRange = sheet.getRange(startRow, 1, lastRow, 10);
  var data = dataRange.getValues();
  
  for (var i = 1; i < data.length; ++i) {
    // get informations from excel to use
    var row = data[i];
    var emailAddress = row[2];
    
    // get current date , and customer entered date to calculate 
    var bDate = sheet.getRange(i+1,2).getValue(); // Date from SpreadSheet
    
    var cDate=new Date();
    var cMonth = cDate.getMonth(),
    bMonth = bDate.getMonth();
    var cDay = cDate.getDate(),
    bDay = bDate.getDate();
    // calculating month and day difference for answers.
    var MonthDifferent=cMonth-bMonth;
    var DayDifferent=cDay-bDay;
    
    //picking color
    var color= row[3];
    
    // giving rows to ''Message Sent' to prevent sending same mail twice.
    var birthdayMessageChecker=row[5];
    var monthlyMessageChecker=row[6];
    var dailyMessageChecker=row[7];
    var isSent = 'Message Sent';
   
    // calculate age for choosing which html to send
    var CustomerAge = cDate.getFullYear() - bDate.getFullYear();
    
    
    //  Month Check- Send mail-Check if its already sent
    if(MonthDifferent==-1 && DayDifferent==0 && monthlyMessageChecker !== isSent ){
    
      var monthlyMessageChecker="Message Sent";
console.log(monthmessage,"Message sent to this Address",emailAddress);
  MailApp.sendEmail(emailAddress, subject2, monthmessage,{htmlBody:reminder2});
  // creating setvalue for specified row with message sent to prevent .
        sheet.getRange(startRow + i, 7).setValue(monthlyMessageChecker);
     
    
    }

    // 1  Day Check- Send mail
     if(MonthDifferent==0 && DayDifferent==-1 && dailyMessageChecker !== isSent ){
      var dailyMessageChecker="Message Sent";
      console.log(onedaymessage,"Message sent to this Address",emailAddress);
     MailApp.sendEmail(emailAddress, subject2, monthmessage,{htmlBody:reminder1});
      sheet.getRange(startRow + i, 8).setValue(dailyMessageChecker);
         
    
    }
   
   //  Birthday Check - Send mail
     
    if(MonthDifferent==0 && DayDifferent==0 && birthdayMessageChecker !== isSent){

      // check which color did customer pick
       if(color=='Blue'){
      
      //check if customer >18 or not (to change type of html)
        if(CustomerAge>18){
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:blueAdult});
         sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);
       }
       //check if customer >18 or not (to change type of html)
       else {
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:blueTeen});
        sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);

    }
    }
    // check which color did customer pick
     if(color=='Green'){

      
      //check if customer >18 or not (to change type of html)
        if(CustomerAge>18){
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:greenAdult});
         sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);
       }
       
       else {
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:greenTeen});
        sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);

    }
    }
     if(color=='') {
      //check if customer >18 or not (to change type of html)
        if(CustomerAge>18){
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:standart});
         sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);
       }
       //check if customer >18 or not (to change type of html)
       else {
         console.log(birthdaymessage,"Message sent to this Address",emailAddress);
         var birthdayMessageChecker="Message Sent";
        MailApp.sendEmail(emailAddress, subject2, birthdaymessage,{htmlBody:standart1});
        sheet.getRange(startRow + i, 6).setValue(birthdayMessageChecker);

    }
    }
    }
    
      
   SpreadsheetApp.flush();
    }

     
  
}