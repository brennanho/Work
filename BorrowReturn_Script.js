//Created on 11/01/2017 by Brennan Ho (Summer and Fall 2017 Intern)
//Version 1.0

/*For Future Developer(s):

If you require to add a new item to the form e.g. "Mag 3 Charger", you DO NOT need to modify anything in this script. In other words, if a new desired item belongs in one of the already existing question categories,
you can simply add it on the form end and not worry about making any modificatiosn to this script. If you need to create a new question, you will need to make modifications on the script. Note that adding a new 
question will not "break" the script, but rather no updates, emails, or checks will be done on the cell/column of the new item. 

If you do need to add a new question to the form, you will need to make the following updates to this script:

  1. In returnBorrowMain(), make a new variable corresponding to the column of the newly added question (lines 47-56)
  2. In status == "Borrow" block, add a new if statement corresponding to the variable you just created
  3. On line 75, add (concatenate) the new variable to the message variable
  4. In the beginning of the else (status == "return") block, declare a variable corresponding to the new question's column
  5. In the if block (line 97-111), add another if statement corresponding to the new question. Follow the template of the other if statements.
  6. In the if statement (line 112), add a hasReturned("NEW ITEM CELL") as another condition to check i.e. "&& hasReturned(ss.getRange(i,8))"
  7. On line 147, append the value ss.getRange(lastActiveRow,4).getValue() to the variable borrowItemsAttempt
  8. In the for loop of the helper function, mulitpleBorrows, declare a new variable corresponding to the new cell/column
  9. On lines 162-167, add a new if statement with a push method inside the block corresponding to the new variable
  10. In the sendEmails function, declare a new variable corresponding to the name of the new item as an empty string
  11. On lines 222-227, add a new if statement following the template of the previous 3 if statements corresponding to the new question/item
  12. On line 233, add (concatenate) the new variable to the message variable
   
  DONE!

*/


//Below is triggered once a day. This sends out emails to users who have outstanding tech items, you must be on notify@yorkhouse.ca to view currently set triggers
function emailSendMain() 
{
  
  var ss = SpreadsheetApp.getActiveSheet();
  var lastActiveRow = ss.getLastRow();
  var lastActiveCol = ss.getLastColumn();
  var today = new Date();
  
  sendEmails(ss,today,lastActiveRow);
  
}

//Below is triggered on form submit, Borrow and Return sends a notification email to tech.support@yorkhouse.ca
function returnBorrowMain()
{
  var ss = SpreadsheetApp.getActiveSheet();
  var lastActiveRow = ss.getLastRow();
  var timeStamp = ss.getRange(lastActiveRow,1).getValue();
  var lastUsername = ss.getRange(lastActiveRow,2).getValue();
  var lastLaptopCharger = ss.getRange(lastActiveRow,3).getValue();
  var lastPhoneCharger = ss.getRange(lastActiveRow,4).getValue();
  var lastOtherEquipment = ss.getRange(lastActiveRow,5).getValue();
  var status = ss.getRange(lastActiveRow,6).getValue();
  var emailNerds = "tech.support@yorkhouse.ca";
  var emailIntern = "it.intern@yorkhouse.ca";
  
  //User is borrowing 1 or more items
  if (status == "Borrow")
  {
    
    //Checks to see if user has borrowed more than one of the same item before first returning. This sends a different email to the user if true.
    if (multipleBorrows(ss,lastActiveRow) == true)
      return;
  
    //Checks to see if user has indeed borrowed an item in a category (empty string implies not borrowed). If borrowed, prepend the category for the email message.
    if (lastLaptopCharger != "")
      lastLaptopCharger = "Laptop Charger: ".bold() + ss.getRange(lastActiveRow,3).getValue() + "<br/>";
    if (lastPhoneCharger != "")
      lastPhoneCharger = "Phone Charger: ".bold() + ss.getRange(lastActiveRow,4).getValue() + "<br/>";
    if (lastOtherEquipment != "")
      lastOtherEquipment = "Other Equipment: ".bold() + ss.getRange(lastActiveRow,5).getValue() + "<br/>";
      
    subject = "Tech Office Equipment Borrow Form: user has borrowed a tech item";
    message = lastUsername + " has borrowed the following item(s):<br/><br/>" + lastLaptopCharger + lastPhoneCharger + lastOtherEquipment + "<br/>" + "Thank you!";
    MailApp.sendEmail({to: emailNerds, subject: subject, htmlBody: message});
    MailApp.sendEmail({to: emailIntern, subject: subject, htmlBody: message});
    ss.getRange(lastActiveRow,7).setValue("No");
  }
    
  else // User is returning i.e. status == "Return"
  { 
    
    //Each itteration is a different user on the spreadsheet sorted by date
    for (var i = 2; i < lastActiveRow; i++)
    {
      var status = ss.getRange(i,6).getValue();
      var username = ss.getRange(i,2).getValue();
      var laptopCharger = ss.getRange(i,3).getValue();
      var phoneCharger = ss.getRange(i,4).getValue();
      var otherEquipment = ss.getRange(i,5).getValue();
      var returnedItems = [];
      if (username.toLowerCase() == lastUsername.toLowerCase() && ss.getRange(i,2).getFontLine() != "line-through") //Checks to see if user has not already returned an item
      {
        
        //Checks to see if intended return matches existing borrow and also the cell is not an empty string
        if (laptopCharger == lastLaptopCharger && laptopCharger != "")
        {
          ss.getRange(i,3).setFontLine(['line-through']);
          returnedItems.push("Laptop Charger: ".bold() + laptopCharger);
        }
        if (phoneCharger == lastPhoneCharger && phoneCharger != "")
        {
          ss.getRange(i,4).setFontLine(['line-through']);
          returnedItems.push("Phone Charger: ".bold() + phoneCharger);
        }
        if (otherEquipment == lastOtherEquipment && otherEquipment != "")
        {
          ss.getRange(i,5).setFontLine(['line-through']);
          returnedItems.push("Other Equipment: ".bold() + otherEquipment);
        }
        if (hasReturnedItem(ss.getRange(i,3)) && hasReturnedItem(ss.getRange(i,4)) && hasReturnedItem(ss.getRange(i,5))) // If user has returned all their borrowed items, strike through their username and change their status to returned
        {
          ss.getRange(i,2).setFontLine(['line-through']);
          ss.getRange(i,6).setValue("Returned on " + timeStamp);
          ss.getRange(i,7).setValue("No");
        }
        
        returnedItems.unshift(timeStamp); // Prepend the timestamp for emailing out
        break;
      }
    }
      
    if (returnedItems.length <= 1) // Do not email out a succesful return user has already returned or attempted to return an item that was not previously borrowed
      return;
    
    subject = "Tech Office Equipment Borrow Form: user has returned their borrowed tech item";
    message = "Tech equipment has been returned on " + returnedItems[0] + " by user: "+ username.bold() + "<br/><br/>";
    
    for (var j = 1; j < returnedItems.length; j++)
          message += returnedItems[j] + "<br/>";
         
    message += "<br/>Thank you!";
    MailApp.sendEmail({to: emailIntern, subject: subject, htmlBody: message});
    MailApp.sendEmail({to: emailNerds, subject: subject, htmlBody: message});
    ss.deleteRow(lastActiveRow);
  }
}

//*----------------------------*// 
//--BELOW ARE HELPER FUNCTIONS--//
//*----------------------------*//

//Checks to see if user is borrowing the same item without returning the previous
function multipleBorrows(ss,lastActiveRow)
{
  var userName = ss.getRange(lastActiveRow,2).getValue();
  var borrowItemsAttempt = [ss.getRange(lastActiveRow,3).getValue(),ss.getRange(lastActiveRow,4).getValue(),ss.getRange(lastActiveRow,5).getValue()];
  var multipleItems = []; // Keeps track of which items have been borrowed more than once before returning
  
  for (var i = 2; i < lastActiveRow; i++)
  {
    var checkTimeStamp = ss.getRange(i,1).getValue();
    var checkUsername = ss.getRange(i,2).getValue();
    var checkLaptopCharger = ss.getRange(i,3).getValue();
    var checkPhoneCharger = ss.getRange(i,4).getValue();
    var checkOtherEquipment = ss.getRange(i,5).getValue();
    
    if (userName == checkUsername && ss.getRange(i,2).getFontLine() != "line-through")
    {
      
      //Checks attempted borrow item with previously borrowed items before returning
      if (borrowItemsAttempt[0] == checkLaptopCharger && checkLaptopCharger != "")
        multipleItems.push("Laptop Charger: ".bold() + borrowItemsAttempt[0]);
      if (borrowItemsAttempt[1] == checkPhoneCharger && checkPhoneCharger != "")
        multipleItems.push("Phone Charger: ".bold() + borrowItemsAttempt[1]);
      if (borrowItemsAttempt[2] == checkOtherEquipment && checkOtherEquipment != "")
        multipleItems.push("Other Equipment: ".bold() + borrowItemsAttempt[2]);
        
      //If user has borrowed more than one of the same item before returning, email out a multiple borrow email to Nerds and intern   
      if (multipleItems.length > 0)
      {
        multipleItems.unshift(checkTimeStamp);
        emailIntern = "it.intern@yorkhouse.ca";
        emailNerds = "tech.support@yorkhouse.ca";
        subject = "Tech Office Equipment Borrow Form: user has borrowed multiples of the same item";
        message = "Item(s) borrowed more than once without first returning at " + multipleItems[0] + " by user: "+ userName.bold() + "<br/><br/>";
        for (var j = 1; j < multipleItems.length; j++)
        {
          message += multipleItems[j] + "<br/>";
        }
        message += "<br/>Thank you!";
        ss.getRange(lastActiveRow,7).setValue("Yes");
        MailApp.sendEmail({to: emailIntern, subject: subject, htmlBody: message});
        MailApp.sendEmail({to: emailNerds, subject: subject, htmlBody: message});
        return true;
      }
    }
  }
  Logger.log("User has not attempted to borrow the same item twice");
  return false;
  
}

//Checks and returns Boolean for whether a cell has a strikethrough
function hasReturnedItem(cell)
{
  if (cell.getFontLine() == "line-through" || cell.getValue() == "")
    return true;
    
  return false;
}

//Checks and returns Boolean for whether item has been borrowed for more than 3 days
function passedDue(today, dateCell)
{
  return (((today.valueOf() - dateCell.getValue().valueOf())/(1000*60*60*24)) > 3);
}

//Itterates through spreadsheet and sends emails out to students/staff that do not have their email striked through and borrowed time is greater than 3 days
function sendEmails(ss, today, lastRow)
{
  //Each itteration is a different user on the spreadsheet sorted by date
  for (var i = 2; i <= lastRow; i++)
  {
       var borrowedDateCell = ss.getRange(i,1);
       var usernameCell = ss.getRange(i,2);
       var laptopCharger = "";
       var phoneCharger = "";
       var otherEquipment = "";
       
       //Check if user has returned their tech items. If not, prepend the category of item in message
       if (hasReturnedItem(ss.getRange(i,3)) == false)
         laptopCharger = "Laptop Charger: ".bold() + ss.getRange(i,3).getValue() + "<br/>";
       if (hasReturnedItem(ss.getRange(i,4)) == false)
         phoneCharger = "Phone Charger: ".bold() + ss.getRange(i,4).getValue() + "<br/>";
       if (hasReturnedItem(ss.getRange(i,5)) == false)
         otherEquipment = "Other Equipment: ".bold() + ss.getRange(i,5).getValue() + "<br/>";
       
       if (hasReturnedItem(usernameCell) == false && passedDue(today, borrowedDateCell) == true)
       {
          email = usernameCell.getValue() + "@yorkhouse.ca";
          subject = "Tech Office Equipment Borrow Form: overdue borrowed item(s)";
          message = "Please return the following items back to the tech office at your earliest convenience: <br/><br/>" + laptopCharger + phoneCharger + otherEquipment + "<br/><br/>" + "Thank you!";
          MailApp.sendEmail({to: email, subject: subject, htmlBody: message});
       }
  }
}