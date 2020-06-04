/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

  }
});

export async function run(event) {

    var item = Office.context.mailbox.item;
    var body = item.body;
    var response = "Sentiment"
    Office.context.mailbox.item.body.getAsync("text", function callback(result) {
      var data = JSON.stringify(result.value);
      var url = "https://shubhamiiest.pythonanywhere.com/ok?text="+item.subject+" "+data;
      var xhr = new XMLHttpRequest();
      xhr.open('GET', url);
      xhr.onload = function(e) {
        response = JSON.parse(this.response);
        if(response.key === 'Positive'){
          document.getElementById("temp").src = "../../assets/positive.png";
        }else{
          document.getElementById("temp").src = "../../assets/negative.png";
        }
        document.getElementById("item-subject").innerHTML = "Sentiment Analysis Result is :  '<b><u>" + response.key+"'</u></b>" ;
      }
      xhr.send();
    });

    
    var data = null;
    var xhr1 = new XMLHttpRequest();
    xhr1.addEventListener("readystatechange", function () {
      if (this.readyState === this.DONE) {
        var t = JSON.parse(this.responseText);
        document.getElementById("corona").innerHTML = "Corona cases in India >= '<b>" + t.key+"'</b></br>Please stay at home to stay safe" ;
        //console.log(this.responseText);
      }
    });

    xhr1.open("GET", "https://shubhamiiest.pythonanywhere.com/corona");
    xhr1.send(data);
    Office.context.mailbox.item.body.prependAsync(response, function (asyncResult) {
      if (asyncResult.status == "failed") {
        showMessage("Action failed with error: " + asyncResult.error.message);
      }
    });
   
  }
  async function checkSignature(event) {
    var signature = localStorage.getItem('flag') ;//&& localStorage.getItem ('persona_data');
    if(signature && signature == 'true')
    {
        AddSignature(event);
        console.log(JSON.parse (localStorage.getItem ('persona_data')).name);
    }
    else {
        AddClickableNotification(event);
        console.log(JSON.parse(signature));
    }
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