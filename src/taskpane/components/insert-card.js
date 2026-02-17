/* Insert rich HTML card into compose email */

import { getWorkItemWebUrl } from "../services/devops-api.js";
import { insertHtmlIntoCompose } from "../services/outlook-mail.js";
import { setButtonLoading, showStatus } from "./ui-helpers.js";

let currentItem = null;

export function init() {
  document.addEventListener("item-selected-for-insert", (e) => {
    currentItem = e.detail.item;
    renderPreview(currentItem);
  });

  const btn = document.getElementById("btn-insert");
  if (btn) {
    btn.addEventListener("click", () => handleInsert());
  }
}

function renderPreview(item) {
  const preview = document.getElementById("insert-preview");
  const btn = document.getElementById("btn-insert");
  if (!preview) return;

  preview.innerHTML = buildCardHtml(item);
  if (btn) btn.style.display = "inline-block";
}

function buildCardHtml(item) {
  const fields = item.fields;
  const type = fields["System.WorkItemType"] || "Work Item";
  const title = fields["System.Title"] || "";
  const state = fields["System.State"] || "";
  const assignedTo = fields["System.AssignedTo"]
    ? typeof fields["System.AssignedTo"] === "string"
      ? fields["System.AssignedTo"]
      : fields["System.AssignedTo"].displayName
    : "Unassigned";
  const url = getWorkItemWebUrl(item.id);

  const stateColor = getStateColor(state);

  return `<table cellpadding="0" cellspacing="0" border="0" width="500" style="width:500px;min-width:300px;border:1px solid #e1e1e1;border-radius:4px;font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;overflow:hidden;margin-bottom:16px;">
  <tr>
    <td style="background:#0078d4;padding:10px 14px;color:#ffffff;font-size:14px;font-weight:600;">
      ${type} #${item.id}
    </td>
  </tr>
  <tr>
    <td style="padding:12px 14px;">
      <div style="font-size:15px;font-weight:600;color:#323130;margin-bottom:8px;">
        <a href="${url}" target="_blank" style="color:#0078d4;text-decoration:none;">${title}</a>
      </div>
      <table cellpadding="0" cellspacing="0" border="0" style="font-size:13px;color:#605e5c;">
        <tr>
          <td style="padding:2px 8px 2px 0;font-weight:600;">State:</td>
          <td style="padding:2px 0;">
            <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${stateColor};margin-right:4px;vertical-align:middle;"></span>
            ${state}
          </td>
        </tr>
        <tr>
          <td style="padding:2px 8px 2px 0;font-weight:600;">Assigned To:</td>
          <td style="padding:2px 0;">${assignedTo}</td>
        </tr>
      </table>
      <div style="margin-top:10px;">
        <a href="${url}" target="_blank" style="display:inline-block;padding:6px 14px;background:#0078d4;color:#ffffff;text-decoration:none;border-radius:3px;font-size:12px;font-weight:600;">View in Azure DevOps</a>
      </div>
    </td>
  </tr>
</table>`;
}

function getStateColor(state) {
  const s = (state || "").toLowerCase();
  if (s === "new" || s === "proposed") return "#b2b2b2";
  if (s === "active" || s === "committed") return "#0078d4";
  if (s === "resolved" || s === "closed" || s === "done") return "#107c10";
  return "#a19f9d";
}

async function handleInsert() {
  if (!currentItem) return;
  const btn = document.getElementById("btn-insert");
  setButtonLoading(btn, true);

  try {
    const html = buildCardHtml(currentItem);
    await insertHtmlIntoCompose(html);
    showStatus("Card inserted into email!", "success");
  } catch (err) {
    console.error("Insert error:", err);
    showStatus("Failed to insert: " + err.message, "error");
  } finally {
    setButtonLoading(btn, false, "Insert Card");
  }
}
