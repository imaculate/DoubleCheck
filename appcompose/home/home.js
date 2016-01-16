(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  var EnglishWords = ["you", "thanks",  "congratulations", "congrats","greetings", "afternoon", "morning", "evening", "hi","good" ,"day", "goodnews", "news", "hello", "dear", "one" ]
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#get-subject').click(getSubject);
      jQuery('#check-recipients').click(checkRecipientsMessage);
    });
  };

  function setSubject(){
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync('Hello world!');
  }

  function getSubject(){
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function(result){
      app.showNotification('The current subject is', result.value);
    });
  }

  function addToRecipients(){
    var item = Office.context.mailbox.item;
    var addressToAdd = {
      displayName: Office.context.mailbox.userProfile.displayName,
      emailAddress: Office.context.mailbox.userProfile.emailAddress
    };

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
  }

  function checkRecipientsMessage(){
    jQuery("notification-message").empty();
    var item = Office.context.mailbox.item;
    var toRecipients, ccRecipients, bccRecipients;
    
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;

       // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients.
            if(asyncResult.value){
              checkAddresses(asyncResult);
            }
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            if(asyncResult.value){
              checkAddresses(asyncResult);
            }
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            if(asyncResult.value){
              checkAddresses(asyncResult);
            }
        }
                        
        }); // End getAsync for bcc-recipients.
     }

    
  }

  function checkAddresses (asyncResult) {
    var myAddress = Office.context.mailbox.userProfile.emailAddress;
     var salutation = getFromBody(); 
    if(asyncResult.value.length == 0){
      write("Your email has no recipients.");
      return;

    }else if(asyncResult.value.length == 1 ){
       
        if(salutation == ""){
          write("Your email has no salutation, please add that in and press this button again");
          //check that you are not addressing yourself - 

        }else{
        possibleName = salutation[salutation.length-1].replace(/[^A-Za-z0-9]/g, '');
        if(asyncResult.value[0].displayName.substring(possibleName) < 0){
          //it could be that its not a name , check from dictionary
          if(EnglishWords.indexOf(possibleName.toLowerCase().replace(/[^A-Za-z0-9]/g, '')) != -1){//its an english word ,thanks, you , congratulations , hi , etc
              write("We think you haven't addressed the recipient , we might be wrong");
          }
          write("You are addressing "+ possibleName+ " and are sending to "+ asyncResult.value[0].displayName);

        }

        //now check that I haven't addressed myself.
        if(asyncResult.value[0].emailAddress == myAddress){
          write("You are sending this email to yourself, is this intentional?");
        }
      }
    }else{ 
      write("There are many recipients to this email, is it intentional?")
  }
}

function getFromBody(){
  var salutation ="";
  Office.context.mailbox.item.body.getAsync("text",
        function callback(asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else { 
              message = asyncResult.value;
              write("Full message"+ message);
              lines = message.split('\n'); 
              //get the first line with words, its most likely the salutation. 
              i = 0;
              var firstLine = "";
              while(firstLine.length<0 && i<lines.length){
                  firstLine = lines[i];
                  i++;
              }
              if(i!= lines.length){
                salutation = firstLine.split(' ');
              }
              

            }
        });
      return salutation;
}

  function write(message){
    document.getElementById('notification-message').innerText += message; 
}

})();
