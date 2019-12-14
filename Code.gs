function onOpen() {
  //function runs by default when the google sheet is open  
  
  // add button on user interface 
  var user_interface = SpreadsheetApp.getUi();
  
  user_interface.createMenu("Send Email").addItem("Send", 'sendEmailUI').addToUi();
  
}
function sendEmailUI(){
  
  // add button on user interface 
  var user_interface = SpreadsheetApp.getUi();
  
  //user_interface.alert("Send Email");
  
  // get active sheet
  var active_sheet= SpreadsheetApp.getActiveSheet();
  
  // get selected cell by the user
  var selected_cell= active_sheet.getActiveCell();
  
  // active row number
  var row = selected_cell.getRow();
  
  var range= active_sheet.getRange(row,1,1,active_sheet.getLastColumn());
  
  // get data values
  var data= range.getValues()[0];
  
  
  var user_data={
    'First_Name':data[0],
    'Last_Name': data[1],
    'Subject': data[2],
    'Message': data[3],
    'Email': data[4],
    'Sent': data[5],
  }
  
  // if the email is not sent already
  
  if (!user_data.Sent){
    
    // messageon alert box
    var message= "Do you want to send email to "+ user_data.First_Name + " " + user_data.Last_Name +" ?";
    
    var result= user_interface.alert(message,user_interface.ButtonSet.OK_CANCEL);
    
    
    // Based on ok or cancel
    
    if (result == user_interface.Button.OK){
      
      Logger.log(user_data);
      sendEmail(user_data);
      active_sheet.getRange(row,6).setValue('sent');
      user_interface.alert("Sent");
    
    }
    
    else {
      
      
      //user_interface.alert("Sending Cancelled");
      
      }
      
      }
  
     else {
      
      
      user_interface.alert("Already Sent");
      
      }
    
    
  
  

  Logger.log(user_data)
  
  console.log("Successfull");
  
}






// send email function
function sendEmail(user_data){
  
  // final email body
  html = templateEmail(user_data);
  
  
  // send email function
  MailApp.sendEmail({to:user_data.Email,subject:user_data.Subject,htmlBody:html})

}



// email template and user data input
function templateEmail(user_data){
  
  // load template email
  var main = HtmlService.createTemplateFromFile('email_template');
  
  // add user relatd values
  main.user_data = user_data;
  
  // final email 
  var html = main.evaluate().getContent();
  
  return html
  
}

