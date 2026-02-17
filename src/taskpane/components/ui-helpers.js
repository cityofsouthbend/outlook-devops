/* Shared DOM utility helpers */

import teamMembers from "../config/team-members.json";
import workItemTypes from "../config/work-item-types.json";

export function populateTeamDropdown(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;
  select.innerHTML = "";
  teamMembers.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m.value;
    opt.textContent = m.label;
    select.appendChild(opt);
  });
}

export function populateWorkItemTypes(containerId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = "";
  workItemTypes.forEach((t, i) => {
    const label = document.createElement("label");
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "workItemType";
    radio.value = t.value;
    if (i === 0) radio.checked = true;
    const span = document.createElement("span");
    span.textContent = t.label;
    label.appendChild(radio);
    label.appendChild(span);
    container.appendChild(label);
  });
}

export function populateTypeFilter(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;
  workItemTypes.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t.value;
    opt.textContent = t.label;
    select.appendChild(opt);
  });
}

export function setButtonLoading(btn, loading, originalText) {
  if (loading) {
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Working...';
  } else {
    btn.disabled = false;
    btn.textContent = originalText || "Submit";
  }
}

export function showStatus(message, type) {
  const banner = document.getElementById("status-banner");
  if (!banner) return;
  banner.textContent = message;
  banner.className = "status-banner " + (type || "info");
  banner.style.display = "block";
  if (type === "success") {
    setTimeout(() => {
      banner.style.display = "none";
    }, 5000);
  }
}

export function hideStatus() {
  const banner = document.getElementById("status-banner");
  if (banner) banner.style.display = "none";
}

export function debounce(fn, delay) {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

export function getTypeBadgeClass(type) {
  if (!type) return "";
  return type.toLowerCase().replace(/\s+/g, "-");
}

export function getStateDotClass(state) {
  if (!state) return "";
  return state.toLowerCase().replace(/\s+/g, "-");
}
