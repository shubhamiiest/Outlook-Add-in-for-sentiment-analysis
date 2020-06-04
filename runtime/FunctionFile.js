/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {};


// Check if the local storage has signature
async function checkSignature(event) {
	
	var item = Office.context.mailbox.item;
    var body = item.body;
    var response = "Sentiment"
    Office.context.mailbox.item.body.getAsync("text", function callback(result) {
      var data = JSON.stringify(result.value);
      var url = "https://shubhamiiest.pythonanywhere.com/ok?text="+data;
      var xhr = new XMLHttpRequest();
      xhr.open('GET', url);
      xhr.onload = function(e) {
        response = JSON.parse(this.response);
        document.getElementById("item-subject").innerHTML = "Sentiment Analysis Result is :  '<b><i>" + response.key+"'</i></b>" ;
        document.getElementById("item-body").innerHTML = data ;

      }
      xhr.send();
    });
    AddClickableNotification(event, response.key);
    event.completed();
    Office.context.mailbox.item.body.prependAsync(response, function (asyncResult) {
      if (asyncResult.status == "failed") {
        showMessage("Action failed with error: " + asyncResult.error.message);
      }
    });
	//console.log(JSON.parse(signature));
	
	event.completed();
}

function AddClickableNotification(event, response) {
    Office.context.mailbox.item.notificationMessages.addAsync
    (
        "my_progress_infobar_id_00",
        {
            type : Microsoft.Office.WebExtension.MailboxEnums.ItemNotificationMessageType.InsightMessage,
            message : "Sentiment Analysis Result is : "+response,
            icon : "smiley",
            actions :
            [
                {
                    "actionType" : Microsoft.Office.WebExtension.MailboxEnums.ActionType.ShowTaskPane,
                    "actionText" : "set signature",
                    "commandId" : "MRCS_TpBtn0",
                    "contextData" : "{''}"
                }
            ]
        },
        function (asyncResult)
        {
            console.log(JSON.stringify(asyncResult));
            event.completed();
        }
    );
}


// Set Signature in Body Code Start
async function AddSignature(event) {
	var data = JSON.parse(localStorage.getItem ('persona_data'));
	var src;
	if(data.logo_pref == 2)
	{
		src = "https://signaturetestaddins.azurewebsites.net/happy_image.png";
	}
	else{
		src = "https://signaturetestaddins.azurewebsites.net/new_img.jpg";
	}
const Signature = "<br><br><br><br><br><html>"+
	"<style>img {float: left; padding-left: 10px; padding-right: 5px;}</style>"+
	"<body>"+
		"<div>"+
			"<img width=\"100px\" height=\100px\" src="+ src +">"+
			"<strong style=\"font-family: Arial Black;font-size: 22pt;color:#000000;\">"+data.name+"</strong>"+
			"<br/>"+
			"<strong style=\"font-size: 15pt; font-family: Arial, sans-serif; height:105px; \">"+data.email+"</strong>"+
			"<br/>"+
			"<strong style=\"font-size: 15pt; font-family: Arial, sans-serif; color:#000000;\" class=\"ng-binding\">"+data.anything+"</strong>"+
			"<br/>"+
		"</div>"+
	"</body>"+
"</html>";
Office.context.mailbox.item.body.setSignatureAsync(
    Signature, {coercionType:Office.CoercionType.Html},
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Failed to set signature");
        }
        else {
            console.log("Signature Set");
		}
		event.completed();
	});
}
// Set Signature in Body Code End







