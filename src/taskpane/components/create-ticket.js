/* Create work item module */

import {
  createWorkItem,
  uploadAttachment,
  addAttachmentRelation,
  updateWorkItemField,
  setCurrentIteration,
  getGraphUsers,
  getWorkItemWebUrl,
  searchParentItems,
} from "../services/devops-api.js";
import {
  getReadMessageInfo,
  getCurrentUserName,
  getEmailBody,
  getAttachmentContent,
} from "../services/outlook-mail.js";
import { setButtonLoading, showStatus, debounce, hideStatus } from "./ui-helpers.js";

let selectedParentId = null;

function getBodyField() {
  return "System.Description";
}

export function init() {
  const info = getReadMessageInfo();

  // Populate subject placeholder from email subject
  const subjectInput = document.getElementById("item-subject");
  if (subjectInput && info.subject) {
    subjectInput.placeholder = info.subject;
  }

  // Render attachment checkboxes
  renderAttachments(info.attachments);

  // Parent search
  const parentInput = document.getElementById("parent-search");
  if (parentInput) {
    parentInput.addEventListener("input", debounce(handleParentSearch, 400));
  }

  // Listen for parent-selected from search tab
  document.addEventListener("parent-selected", (e) => {
    selectParent(e.detail.id, e.detail.title);
  });

  // Create button
  const btn = document.getElementById("btn-create");
  if (btn) {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      handleCreate();
    });
  }
}

function renderAttachments(attachments) {
  if (!attachments || attachments.length === 0) return;
  const group = document.getElementById("attachments-group");
  const list = document.getElementById("attachment-list");
  if (!group || !list) return;

  const imageExts = [".png", ".jpg", ".gif", ".jpeg", ".bmp", ".eps"];
  let html = "";
  attachments.forEach((att) => {
    const isImage = imageExts.some((ext) => att.name.toLowerCase().endsWith(ext));
    if (!isImage) {
      html += `<div class="attachment-item">
        <input type="checkbox" checked value="${att.name}" />
        <span>${att.name}</span>
      </div>`;
    }
  });
  if (html) {
    html += '<p style="font-size:11px;color:#605e5c;margin-top:4px;">Inline images are included by default.</p>';
    list.innerHTML = html;
    group.style.display = "block";
  }
}

async function handleParentSearch() {
  const input = document.getElementById("parent-search");
  const resultsDiv = document.getElementById("parent-results");
  const text = input.value.trim();

  if (!text) {
    resultsDiv.classList.remove("visible");
    resultsDiv.innerHTML = "";
    return;
  }

  try {
    const items = await searchParentItems(text);
    if (items.length === 0) {
      resultsDiv.innerHTML = '<div class="parent-result-item">No results found</div>';
    } else {
      resultsDiv.innerHTML = items
        .map(
          (item) =>
            `<div class="parent-result-item" data-id="${item.id}" data-title="${item.fields["System.Title"]}">${item.fields["System.Title"]} (${item.fields["System.WorkItemType"]} - ${item.id})</div>`
        )
        .join("");
      resultsDiv.querySelectorAll(".parent-result-item[data-id]").forEach((el) => {
        el.addEventListener("click", () => {
          selectParent(el.dataset.id, el.dataset.title);
          resultsDiv.classList.remove("visible");
          resultsDiv.innerHTML = "";
          input.value = "";
        });
      });
    }
    resultsDiv.classList.add("visible");
  } catch (err) {
    console.error("Parent search error:", err);
    resultsDiv.innerHTML = '<div class="parent-result-item">Search error</div>';
    resultsDiv.classList.add("visible");
  }
}

function selectParent(id, title) {
  selectedParentId = id;
  const chip = document.getElementById("selected-parent");
  chip.innerHTML = `<span>${title} (${id})</span><span class="remove-parent" title="Remove">x</span>`;
  chip.style.display = "inline-flex";
  chip.querySelector(".remove-parent").addEventListener("click", () => {
    selectedParentId = null;
    chip.style.display = "none";
    chip.innerHTML = "";
  });
}

async function handleCreate() {
  const btn = document.getElementById("btn-create");
  hideStatus();

  // Gather form values
  const typeRadio = document.querySelector('input[name="workItemType"]:checked');
  if (!typeRadio) {
    showStatus("Please select a work item type.", "error");
    return;
  }
  const ticketType = typeRadio.value;

  if (!selectedParentId) {
    showStatus("Please select a parent item.", "error");
    return;
  }

  const subjectInput = document.getElementById("item-subject");
  const ticketTitle = subjectInput.value.trim() || subjectInput.placeholder;

  const assignedTo = document.getElementById("team-select").value;
  const creatorName = getCurrentUserName();

  setButtonLoading(btn, true);

  try {
    // Fetch email body
    const info = getReadMessageInfo();
    let emailBody = "";

    try {
      const body = await getEmailBody();
      const header = `<p>Begin display of original email</p><hr>
        <p>From: ${info.from.displayName}<br>Email: ${info.from.emailAddress}</p>`;
      emailBody = header + body;
      console.log("[DevOps] Email body fetched, length:", emailBody.length);
    } catch (err) {
      console.error("[DevOps] Failed to load email body:", err);
      showStatus("Warning: could not load email body.", "info");
    }

    const bodyField = getBodyField(ticketType);
    console.log("[DevOps] Using body field:", bodyField, "for type:", ticketType);
    console.log("[DevOps] emailBody empty?", emailBody.length === 0);

    // Look up creator in Graph users
    const users = await getGraphUsers();
    const creator = users.find((u) => u.displayName === creatorName);

    // Build patch ops - include body field at creation time
    const patchOps = [
      { op: "add", path: "/fields/System.Title", from: null, value: ticketTitle },
      { op: "add", path: `/fields/${bodyField}`, value: emailBody },
    ];

    if (creator) {
      const creatorObj = {
        displayName: creator.displayName,
        url: creator.url,
        _links: { avatar: { href: creator._links.avatar.href } },
        id: creator.originId,
        uniqueName: creator.uniqueName,
        imageUrl: creator.imageUrl,
        descriptor: creator.descriptor,
      };
      patchOps.push({ op: "add", path: "/fields/System.CreatedBy", value: creatorObj });
      patchOps.push({ op: "add", path: "/fields/System.ChangedBy", value: creatorObj });
    }

    if (assignedTo) {
      patchOps.push({ op: "add", path: "/fields/System.AssignedTo", value: assignedTo });
    }

    // Parent link
    const ORG = process.env.DEVOPS_ORG || "southbendin";
    patchOps.push({
      op: "add",
      path: "/relations/-",
      value: {
        rel: "System.LinkTypes.Hierarchy-Reverse",
        url: `https://dev.azure.com/${ORG}/_apis/wit/workItems/${selectedParentId}`,
        attributes: { isLocked: false, name: "Parent" },
      },
    });

    // Create the work item
    console.log("[DevOps] Creating work item with patchOps:", JSON.stringify(patchOps.map(op => ({ op: op.op, path: op.path, valueLength: typeof op.value === 'string' ? op.value.length : 'object' }))));
    const ticket = await createWorkItem(ticketType, patchOps);
    console.log("[DevOps] Work item created:", ticket.id, "fields:", Object.keys(ticket.fields || {}).filter(f => f.includes("Repro") || f.includes("Description")));

    // Check if attachment content API is available (requires Mailbox 1.8)
    const allAttachments = info.attachments || [];
    const canGetContent =
      Office.context.mailbox.item &&
      typeof Office.context.mailbox.item.getAttachmentContentAsync === "function";

    if (canGetContent) {
      // Upload inline images first and resolve cid: references in the body
      const inlineAtts = allAttachments.filter((a) => a.isInline);
      for (let i = 0; i < inlineAtts.length; i++) {
        const att = inlineAtts[i];
        try {
          // Small delay between calls to avoid message channel conflicts
          if (i > 0) await new Promise((r) => setTimeout(r, 200));

          const content = await getAttachmentContent(att.id);

          // Handle both base64 and url formats returned by Office.js
          let blob;
          if (content.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
            blob = await fetch(content.content).then((r) => r.blob());
          } else {
            blob = await fetch(`data:${att.contentType};base64,${content.content}`).then((r) =>
              r.blob()
            );
          }

          const fileData = await uploadAttachment(att.name, blob);
          const escapedName = att.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          const nameNoExt = att.name.replace(/\.[^.]+$/, "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          emailBody = emailBody
            .replace(new RegExp(`src="cid:${escapedName}[^"]*"`, "gi"), `src="${fileData.url}"`)
            .replace(new RegExp(`src="cid:${nameNoExt}[^"]*"`, "gi"), `src="${fileData.url}"`);
        } catch (err) {
          console.warn(`Skipping inline attachment ${att.name}: ${err.message}`);
        }
      }
    } else {
      console.warn("getAttachmentContentAsync not available - inline images will not be resolved");
    }

    // Update body field with resolved inline images (if any were processed)
    if (canGetContent && allAttachments.some((a) => a.isInline)) {
      await updateWorkItemField(ticket.id, bodyField, emailBody);
    }

    // Upload email body as HTML file attachment
    const attachData = await uploadAttachment("OriginalEmail.html", emailBody);
    await addAttachmentRelation(ticket.id, attachData.url);

    // Set current iteration
    await setCurrentIteration(ticket.id);

    // Upload selected file attachments (non-inline)
    if (canGetContent) {
      const checked = document.querySelectorAll('#attachment-list input[type="checkbox"]:checked');
      let fileIdx = 0;
      for (const checkbox of checked) {
        const att = allAttachments.find((a) => a.name === checkbox.value);
        if (att) {
          try {
            if (fileIdx > 0) await new Promise((r) => setTimeout(r, 200));
            fileIdx++;

            const content = await getAttachmentContent(att.id);

            let blob;
            if (content.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
              blob = await fetch(content.content).then((r) => r.blob());
            } else {
              blob = await fetch(`data:${att.contentType};base64,${content.content}`).then((r) =>
                r.blob()
              );
            }

            const fileData = await uploadAttachment(att.name, blob);
            await addAttachmentRelation(ticket.id, fileData.url);
          } catch (err) {
            console.warn(`Skipping attachment ${att.name}: ${err.message}`);
          }
        }
      }
    }

    const url = getWorkItemWebUrl(ticket.id);
    showStatus(`Ticket ${ticket.id} created!`, "success");
    setButtonLoading(btn, false, "Create Ticket");

    // Show link
    const panel = document.getElementById("tab-create");
    const link = document.createElement("p");
    link.innerHTML = `<a href="${url}" target="_blank" style="color:#0078d4;">View ticket ${ticket.id} in Azure DevOps</a>`;
    panel.appendChild(link);
  } catch (err) {
    console.error("Create ticket error:", err);
    showStatus("Failed to create ticket: " + err.message, "error");
    setButtonLoading(btn, false, "Create Ticket");
  }
}
