(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  var EnglishWords = ["you", "thanks",  "congratulations", "congrats","greetings", "afternoon", "morning", "evening", "hi","good" ,"day", "goodnews", "news", "hello", "dear", "one" ]
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      
      jQuery('#check-reply-self').click({callback: checkReplySelf}, checkRecipients);
      jQuery('#check-addressee-recipient').click({callback: matchAddressee} , checkRecipients );
      jQuery('#check-reply-all').click({callback: checkReplyAll}, checkRecipients);
      jQuery('#check-all').click({callback: checkAddresses}, checkRecipients);
    });
  };




  function checkRecipients(event){
    
    app.showNotification('The status of recipients','');
    var item = Office.context.mailbox.item;
    var toRecipients, ccRecipients, bccRecipients;
    var rcpts = [];
    
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
             rcpts = rcpts.concat(asyncResult.value);
              

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
              rcpts = rcpts.concat(asyncResult.value)

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
              rcpts = rcpts.concat(asyncResult.value)
            
        }
         event.data.callback(rcpts);
                        
        }); // End getAsync for bcc-recipients.
     }

    
  }


  function checkReplySelf(emails){
    var myAddress = Office.context.mailbox.userProfile.emailAddress;

      for(var i =0; i<emails.length; i++){
        if(emails[i].emailAddress == myAddress){
          write("You are sending this email to yourself, is this intentional?");
      }

    }
    

  }

  function checkReplyAll(emails){
    if(emails.length>1){ 
      write("There are many recipients to this email, is it intentional?")
    }
  }

  function matchAddressee(emails){
    if(emails.length == 1 ){
         getAddresseeFromBody(emails); 
     
      
    }

  }

  function checkAddresses (emails) {
      checkReplySelf(emails);
      checkReplyAll(emails);
      matchAddressee(emails);
}

function getAddresseeFromBody(emails){
  var salutation = "";
  Office.context.mailbox.item.body.getAsync("text",
        function callback(asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else { 
              var message = asyncResult.value;
              var lines = message.split('\n'); 
              //get the first line with words, its most likely the salutation. 
              var i = 0;
              var firstLine = "";
              while(firstLine.length<=0 && i<lines.length){
                  firstLine = lines[i];
                  i++;
              }
              if(i!= lines.length){
                 salutation = firstLine.trim().split(' ');
                   if(salutation.length == 0){
                      write("Your email has no salutation, please add that in and press this button again");
                      //check that you are not addressing yourself - 

                    }else{
                        
                        var possibleName = salutation[(salutation.length)-1];
                        possibleName = possibleName.replace(/[^a-zA-Z0-9-]/g, '')
                        if(emails[0].displayName.indexOf(possibleName) < 0){
                        //it could be that its not a name , check from dictionary
                            if(EnglishWords.indexOf(possibleName.toLowerCase()) != -1){//its an english word ,thanks, you , congratulations , hi , etc
                                write("We think you haven't addressed the recipient , we might be wrong");
                            }
                            write("Looks like this email is addressed "+ possibleName+ " but is being sent to "+ emails[0].displayName);

                        }

                        
                }     

            }
          } 
        });
      return salutation;
}

  function write(message){
    document.getElementById('notification-message-body').innerText += message+ '\n'; 
}

})();
