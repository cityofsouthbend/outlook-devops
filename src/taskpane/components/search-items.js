/* Search/browse work items module */

import { searchWorkItems, getWorkItemWebUrl, isAllowedParentType } from "../services/devops-api.js";
import { debounce, getTypeBadgeClass, getStateDotClass, showStatus } from "./ui-helpers.js";
import { showTab } from "./tab-manager.js";

export function init() {
  const searchInput = document.getElementById("search-input");
  const typeFilter = document.getElementById("search-type-filter");

  const doSearch = debounce(() => runSearch(), 400);
  if (searchInput) searchInput.addEventListener("input", doSearch);
  if (typeFilter) typeFilter.addEventListener("change", () => runSearch());
}

async function runSearch() {
  const input = document.getElementById("search-input");
  const typeFilter = document.getElementById("search-type-filter");
  const resultsDiv = document.getElementById("search-results");
  const text = input.value.trim();

  if (!text) {
    resultsDiv.innerHTML = '<p class="placeholder-text">Enter a search term above.</p>';
    return;
  }

  resultsDiv.innerHTML = '<p class="placeholder-text">Searching...</p>';

  try {
    const items = await searchWorkItems(text, typeFilter.value);

    if (items.length === 0) {
      resultsDiv.innerHTML = '<p class="placeholder-text">No results found.</p>';
      return;
    }

    resultsDiv.innerHTML = items.map((item) => renderCard(item)).join("");

    // Wire up action buttons
    resultsDiv.querySelectorAll("[data-action]").forEach((btn) => {
      btn.addEventListener("click", () => handleAction(btn, items));
    });
  } catch (err) {
    console.error("Search error:", err);
    resultsDiv.innerHTML = '<p class="placeholder-text">Search failed. Check PAT token.</p>';
    showStatus("Search failed: " + err.message, "error");
  }
}

function renderCard(item) {
  const fields = item.fields;
  const type = fields["System.WorkItemType"] || "Unknown";
  const title = fields["System.Title"] || "";
  const state = fields["System.State"] || "";
  const assignedTo = fields["System.AssignedTo"]
    ? typeof fields["System.AssignedTo"] === "string"
      ? fields["System.AssignedTo"]
      : fields["System.AssignedTo"].displayName
    : "Unassigned";
  const badgeClass = getTypeBadgeClass(type);
  const stateClass = getStateDotClass(state);
  const url = getWorkItemWebUrl(item.id);

  return `
    <div class="wi-card">
      <div class="wi-card-header">
        <span class="wi-type-badge ${badgeClass}">${type}</span>
        <span class="wi-id">#${item.id}</span>
        <span class="wi-state">
          <span class="state-dot ${stateClass}"></span>
          ${state}
        </span>
      </div>
      <div class="wi-title">${title}</div>
      <div class="wi-assigned">${assignedTo}</div>
      <div class="wi-actions">
        ${isAllowedParentType(type) ? `<button class="btn btn-small btn-outline" data-action="parent" data-id="${item.id}" data-title="${title}">Use as Parent</button>` : ""}
        <button class="btn btn-small btn-outline" data-action="insert" data-id="${item.id}">Insert</button>
        <a class="btn btn-small btn-outline" href="${url}" target="_blank">Open</a>
      </div>
    </div>`;
}

function handleAction(btn, items) {
  const action = btn.dataset.action;
  const id = parseInt(btn.dataset.id, 10);
  const item = items.find((i) => i.id === id);

  if (action === "parent") {
    document.dispatchEvent(
      new CustomEvent("parent-selected", {
        detail: { id: btn.dataset.id, title: btn.dataset.title },
      })
    );
    showTab("create");
  }

  if (action === "insert" && item) {
    document.dispatchEvent(
      new CustomEvent("item-selected-for-insert", { detail: { item } })
    );
    showTab("insert");
  }
}
