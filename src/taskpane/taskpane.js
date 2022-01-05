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

// global variables
let textAttachment = "";

const idURL = ($) => `https://dev.azure.com/southbendin/_apis/wit/workitems?ids=${$.val}&api-version=6.0`;
const wiqlURL = 'https://dev.azure.com/southbendin/_apis/wit/wiql?api-version=6.0';
const ticketURL = ($) =>
  `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/workitems/$${$.val}?api-version=6.0`;
const updateAttachmentURL = ($) => `https://dev.azure.com/southbendin/_apis/wit/workitems/${$.val}?api-version=6.0`;
const createAttachmentURL = `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/attachments?fileName=EmailAsFileAttachment.txt&api-version=6.0`;

const paToken = "Basic " + btoa("Basic" + ":" + "iinmtwdby2a5k3v6dekdc5y53raw7vsavivuss4fm47l4bu6fwzq");

const queryWIQL = `{
  "query": "Select [System.Id], [System.Title], [System.State] From WorkItems Where [System.WorkItemType] = 'Maintenance'"}`;

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
  let ticketTitle = document.getElementById("item-subject").value;

  return [ticketType, selectResult, ticketTitle];
}

function createTicketHeader(parent, ticketTitle) {
  const newBody = ($1, $2) => [
    {"op": "add", "path": "/fields/System.Title", "from": null, "value": `${$1.val}` },
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

function addAttachmentHeader(attachURL) {
  const newBody = ($) => [{"op": "add", "path": "/relations/-", "value": { "rel": "AttachedFile", "url": `${$.val}`, "attributes": {"comment": "Created with Outlook Azure add-in"}}}];

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

async function createNewTicket() {
  let [ticketType, parent, ticketTitle] = grabFormItems();
  
  let responseTicket = await fetch(ticketURL({val: ticketType}), createTicketHeader(parent, ticketTitle));
  let dataTicket = await responseTicket.json();

  let responseNewAttachment = await fetch(createAttachmentURL, createAttachmentHeader(textAttachment));
  let dataAttachment = await responseNewAttachment.json();
  
  let responseAddAttachment = await fetch(updateAttachmentURL({val: dataTicket.id}), addAttachmentHeader(dataAttachment.url));
  let dataAddAttachment = await responseAddAttachment.json();

  document.getElementById("app-body").innerHTML = `<div>DevOps ticket ${dataTicket.id} has been created. View the ticket in DevOps <a href="https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_workitems/edit/${dataTicket.id}">here</a></div>`;
}

// function getIteration() {
//   let request2 = new XMLHttpRequest();
//   let newURL = `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/work/teamsettings/iterations?$timeframe=current&api-version=6.0`;
//   request2.open("GET", newURL);
//   request2.setRequestHeader("Content-Type", "application/json");
//   request2.setRequestHeader(
//     "Authorization",
//     "Basic " + btoa("Basic" + ":" + "iinmtwdby2a5k3v6dekdc5y53raw7vsavivuss4fm47l4bu6fwzq")
//   );
//   request2.send();
//   request2.onload = () => {
//     res2 = JSON.parse(request2.response);
//     let currentPath = res2.value[0].path;
//     createTicket(currentPath);
//   };
// }

// Office generated code when add-in is sideloaded
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    getMaintenanceItems();
    var item = Office.context.mailbox.item;

        
    var attachments = Office.context.mailbox.item.attachments;
    console.log(attachments);
    
    // Write message property value to the task pane
    document.getElementById("item-subject").placeholder = item.subject;
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      var ewsId = Office.context.mailbox.item.itemId;
      var token = result.value;
  
      // var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0); this does not work on API version 1.1
      var restId = ewsId.replaceAll("/", "-").replaceAll("+", "_"); // Convert ewsId to restId
      var getMessageUrl = (Office.context.mailbox.restUrl || 'https://outlook.office365.com/api') + '/v2.0/me/messages/' + restId;
      var xhr = new XMLHttpRequest();
      xhr.open('GET', getMessageUrl);
      xhr.setRequestHeader('Prefer', 'outlook.body-content-type="text"') 
      xhr.setRequestHeader("Authorization", "Bearer " + token);
      xhr.onload = (e) => {
        var json = JSON.parse(xhr.responseText);
        textAttachment = json.Body.Content;
      }
      xhr.send();
    });
    // Office.context.mailbox.item.from.getAsync(function(asyncResult) {
    //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //     textAttachment += asyncResult.value + "\n\n";
    //     console.log("Message from: " + msgFrom.displayName + " (" + msgFrom.emailAddress + ")");
    //   } else {
    //     console.error(asyncResult.error);
    //   }
    // });
    // Office.context.mailbox.item.body.getAsync(
    //   "text",
    //   { asyncContext: "This is passed to the callback" },
    //   function callback(result) {
    //     textAttachment += result.value;
    //     // document.getElementById("item-body").innerHTML = "<b>New Ticket Description:</b> <br/>" + result.value;
    //     // Do something with the result.
    //   }
    // );
     //API call to get Maintenance Items
  }
});

export async function run() {
  createNewTicket();
} // end run()
