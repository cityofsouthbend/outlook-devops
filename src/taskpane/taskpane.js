/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

let textAttachment = "";
let attachments = [];
let selectedAttachments = [];
let attachmentArray = [];
let token = "";


// declaring variables (or functions that return urls) for url addresses to use in various api calls
const idURL = ($) => `https://dev.azure.com/southbendin/_apis/wit/workitems?ids=${$.val}&api-version=6.0`;
const wiqlURL = 'https://dev.azure.com/southbendin/_apis/wit/wiql?api-version=6.0';
const ticketURL = ($) =>
  `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/workitems/$${$.val}?bypassrules=true&api-version=6.0`;
const updateAttachmentURL = ($) => `https://dev.azure.com/southbendin/_apis/wit/workitems/${$.val}?api-version=6.0`;
const createAttachmentURL = `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/attachments?fileName=OriginalEmail.html&api-version=6.0`;
const createAttachmentURL2 = ($) => `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/attachments?fileName=${$.val}&api-version=5.1`;

// the token is assigned to appsadmin@southbendin.gov - we may want to rethink this in future releases with using the Microsoft Graph API
// this token has an expiration date of 1/23/23
const paToken = "Basic " + btoa("Basic" + ":" + "p6l4ydakwngcbswzyilmfhidth5vx57veh3djy6vxiiefap5dfaq");


async function createToken() {
  const tokenHeaders = {method: 'POST', body: JSON.stringify(
    {
      "displayName": "new_token",
      "scope": "app_token",
      "validTo": "2022-12-01T23:46:23.319Z",
      "allOrgs": false
    }
  )};

  let token = await fetch('https://vssps.dev.azure.com/{organization}/_apis/tokens/pats?api-version=6.1-preview.1', tokenHeaders);
  let tokenResponse = await token.json();
  return tokenResponse.patToken.token;
}
let paToken = "Basic " + btoa("Basic" + ":" + createToken());
//"iinmtwdby2a5k3v6dekdc5y53raw7vsavivuss4fm47l4bu6fwzq"
const queryWIQL = `{
  "query": "Select [System.Id], [System.Title], [System.State] From WorkItems Where [System.WorkItemType] = 'Maintenance' AND [System.State] <> 'Archived'"}`;

const headersWIQL = {
  method: "POST",
  headers: { "Content-Type": "application/json;charset=utf-8", Authorization: paToken },
  body: queryWIQL
};

const headersMaintenance = {
  method: "GET",
  headers: { "Content-Type": "application/json", Authorization: paToken }};

function concatIDs(data) {
  let str = "";
  data.workItems.forEach((element) => { str += element.id + ","  });
  return str.slice(0, -1);
}

function createDropdown(data) {
  const maintenanceItems = new Map();
  let html = `<select id="parentItem" name="parentItem">`;
  for (let d = 0; d < data.length; d++) {
    maintenanceItems.set(data[d].id, data[d].fields["System.Title"]);
  }
  const sortedItems = new Map([...maintenanceItems].sort((a, b) => String(a[1]).localeCompare(b[1])));
  sortedItems.forEach(function (v, k) {
    let htmlSegment = `<option value="${k}">${v} - Maintenance Item: ${k}</option>`;
    html += htmlSegment;
  }); 
  html += `</select>`;
  let container = document.querySelector("#dropdown");
  container.innerHTML = html;
}

function grabFormItems() {
  let selectResult = document.getElementById("parentItem").value;
  //Get bug or task ticket
  let ticketType = document.getElementById("bug-ticket").checked ? "Bug" : "Task";
  //get title for ticket
  let ticketTitle = document.getElementById("item-subject");
  let ticketTitleValue = "";
  if ( ticketTitle.value == "") {
    ticketTitleValue = ticketTitle.placeholder;
  } else {
    ticketTitleValue = ticketTitle.value;
  }
  let ticketAssign = document.getElementById("users").value;
  return [ticketType, selectResult, ticketTitleValue, ticketAssign];
}

function createTicketHeader(parent, ticketTitle, creator, dev) {
  fetch('https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/work/teamsettings/iterations?$timeframe=current&api-version=6.0')
    .then( res => console.log(res));
  const newBody = ($1, $2) => [
    {"op": "add", "path": "/fields/System.Title", "from": null, "value": `${$1.val}` },
    {
      "op": "add",
      "path": "/fields/System.CreatedBy",
      "value": {
        "displayName": `${creator.displayName}`,
        "url": `${creator.url}`,
        "_links": {
          "avatar": {
            "href": `${creator._links.avatar.href}`
          }
      },
        "id": `${creator.originId}`,
        "uniqueName": `${creator.uniqueName}`,
        "imageUrl": `${creator.imageUrl}`,
        "descriptor": `${creator.descriptor}`
        }
    },{
      "op": "add",
      "path": "/fields/System.ChangedBy",
      "value": {
          "displayName": `${creator.displayName}`,
          "url": `${creator.url}`,
          "_links": {
            "avatar": {
              "href": `${creator._links.avatar.href}`
            }
        },
          "id": `${creator.originId}`,
          "uniqueName": `${creator.uniqueName}`,
          "imageUrl": `${creator.imageUrl}`,
          "descriptor": `${creator.descriptor}`
      }
      },
    {
      "op": "add",
      "path": "/fields/System.AssignedTo",
      "value": `${dev}`
    },
    {"op": "add", "path": "/relations/-", "value": { "rel": "System.LinkTypes.Hierarchy-Reverse", "url": `https://dev.azure.com/southbendin/_apis/wit/workItems/${$2.val}`, "attributes": { "isLocked": false, "name": "Parent" }}}];
    
  const headersNewTicket = {method: 'POST', headers: { 'Content-Type': 'application/json-patch+json','Authorization': paToken}, body: JSON.stringify(newBody({val: ticketTitle}, {val: parent}))};
  return headersNewTicket;
}

function createAttachmentHeader(emailText) {
  const textBody = ($) => `${$.val}`;
  const makeAttachment = {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", Authorization: paToken },
    body: textBody({val: emailText})
  };
  return makeAttachment;
}

function addAttachmentHeader(attachURL, textAttachment) {
  const newBody = ($) => [
    { "op": "add", 
      "path": "/relations/-", 
      "value": { 
        "rel": "AttachedFile", "url": `${$.val}`,
        "attributes": {
          "comment": "Created with Outlook Azure add-in"
        }
      }
    },
    {
      "op": "replace",
      "path": "/fields/System.Description",
      "value": `${textAttachment}`
    }
  
  ];
  const headerAddAttachment = {
    method: "PATCH",
    headers: { "Content-Type": "application/json-patch+json", Authorization: paToken },
    body: JSON.stringify(newBody({val: attachURL}))
  };
  return headerAddAttachment;
}

async function getMaintenanceItems() {
  let responseWIQL = await fetch(wiqlURL, headersWIQL);
  let dataWIQL = await responseWIQL.json();
  let ids = concatIDs(dataWIQL);
  let responseMaintenance = await fetch(idURL({val: ids}), headersMaintenance);
  let dataMaintenance = await responseMaintenance.json();
  createDropdown(dataMaintenance.value);
}

async function handleAttachments(attachment, ticketInfo) {
  // this works in grabing the attachment and adding it to the ticket - there doesn't seem to be size limits either
    await fetch(`data:${attachment.ContentType};base64, ${attachment.ContentBytes}`)
      .then( res => res.blob())
      .then( async (blob) => {
        await fetch(createAttachmentURL2({val: attachment.Name}), {
          method: "POST",
          headers: { "Content-Type": "application/octet-stream", Authorization: paToken },
          body: blob
        }).then( res =>  res.json())
          .then ( async (data) => {
            console.log(data);
            let temp = textAttachment.replace(`src="cid:${attachment.ContentId}"`, `src="${data.url}"`);
            textAttachment = temp;
            console.log(textAttachment)
            await fetch(updateAttachmentURL({val: ticketInfo.id}), addAttachmentHeader(data.url, textAttachment))
            .then( res => res.json())
          });
      })
} // end handleAttachments function

function createNewTicket() {
  let [ticketType, parent, ticketTitle, assignedDev] = grabFormItems();
  let getCreator = fetch("https://vssps.dev.azure.com/southbendin/_apis/graph/users?api-version=6.0-preview.1", {
    method: "GET",
    headers: {Authorization: paToken}
    }).then( res => res.json() )
      .then( (data) => {
        for (let i = 0; i < data.value.length; i++) {
          if (data.value[i].displayName == document.getElementById("ticketCreator").value) {
            return data.value[i];
          }
        }
    });

  let createTicketMethod = () => {
    getCreator.then( (creator) => {
      fetch(ticketURL({val: ticketType}), createTicketHeader(parent, ticketTitle, creator, assignedDev))
        .then( res => res.json())
        .then( async (dataTicket) => {

          // adding the email body as a html file (includes inline images)
          let resAttach = await fetch(createAttachmentURL, createAttachmentHeader(textAttachment));
          let dataAttach = await resAttach.json()
          let addAttach = await fetch(updateAttachmentURL({val: dataTicket.id}), addAttachmentHeader(dataAttach.url));
  
          // adding attachments selected by user
          const tempSelected = document.querySelectorAll('input[type="checkbox"]:checked');
          for (let g = 0; g < tempSelected.length; g++ ) {
            for ( let h = 0; h < attachmentArray.length; h++ ) {
              if ( tempSelected[g].value == attachmentArray[h].Name) {
                selectedAttachments.push(attachmentArray[h]);
              }
            }
          }
          
          for ( let i = 0; i < selectedAttachments.length; i++ ) {
            await handleAttachments(selectedAttachments[i], dataTicket);
          }


          // When sideloaded app has completed this message is printed to the user's sideload window
          document.getElementById("app-body").innerHTML = `<div>DevOps ticket ${dataTicket.id} has been created. View the ticket in DevOps <a href="https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_workitems/edit/${dataTicket.id}" target="_blank">here</a></div>`;
        })
      });
  }
  createTicketMethod();
} // end createNewTicket()

// Office generated code when add-in is sideloaded
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    let btn = document.getElementById("run");
    btn.addEventListener("click", () => {
      run();
      btn.setAttribute("disabled", "");
    })
    
    getMaintenanceItems();
    var item = Office.context.mailbox.item;
<<<<<<< HEAD
    console.log(item);
    textAttachment += `<p>Begin display of original email</p>
                      <hr>
                      <p>From: ${item.from.displayName} <br>
                      Email: ${item.from.emailAddress} <br>`
    var ticketCreator = Office.context.mailbox.userProfile.displayName;

    // grab the info about attachments from the email 
    attachments = Office.context.mailbox.item.attachments;

    // add the attachments to the add-in 
    let attachP = document.getElementById("files");
    let filesToInclude = "";
    let conditions = ['.png', '.jpg', '.gif', '.jpeg', '.bmp', '.eps', '.jpeg']
    for (let i = 0; i < attachments.length; i++) {
      let test = conditions.some( c => attachments[i].name.includes(c));
      if (!test) {
        filesToInclude += `<input type="checkbox" name="attachedFile" checked value="${attachments[i].name}">${attachments[i].name}<br>`;
      } else {
        filesToInclude += `<input type="checkbox" name="attachedFile" checked value="${attachments[i].name}" style="visibility:hidden;">${attachments[i].name}<br>`;
      }
    }
    filesToInclude += "<hr><p>All inline image files are included by default</p>"
    attachP.innerHTML = filesToInclude;

=======

        
    var attachments = Office.context.mailbox.item.attachments;
    console.log(attachments);
    
>>>>>>> d3c72047e391e7856725fbffc93fd833ce8c836d
    // Write message property value to the task pane
    document.getElementById("item-subject").placeholder = item.subject;
    document.getElementById("ticketCreator").value = ticketCreator; //This is hidden on the html form
    
    // This callback is to grab the inner
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      var ewsId = Office.context.mailbox.item.itemId;
      token = result.value;
  
      // var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0); this does not work on API version 1.1
      let restId = ewsId.replaceAll("/", "-").replaceAll("+", "_"); // Convert ewsId to restId
      var getMessageUrl = (Office.context.mailbox.restUrl || 'https://outlook.office365.com/api') + '/v2.0/me/messages/' + restId;
      var getAttachmentsUrl = getMessageUrl + '/attachments';
      var xhr = new XMLHttpRequest();
      xhr.open('GET', getMessageUrl);
      xhr.setRequestHeader('Prefer', 'outlook.body-content-type="html"') 
      xhr.setRequestHeader("Authorization", "Bearer " + token);
      xhr.onload = (e) => {
        var json = JSON.parse(xhr.responseText);
        textAttachment += json.Body.Content;
        var xhr2 = new XMLHttpRequest();
        xhr2.open('GET', getAttachmentsUrl);
        xhr2.setRequestHeader("Authorization", "Bearer " + token);
        xhr2.onload = (e) => {
          let json2 = JSON.parse(xhr2.responseText);
          attachmentArray = json2.value;
        
        }
        xhr2.send();
      }
      xhr.send();
    });
  }
});

export async function run() {
  createNewTicket();
} 
