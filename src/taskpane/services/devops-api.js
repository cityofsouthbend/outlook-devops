/* Azure DevOps REST API service module */

const ORG = process.env.DEVOPS_ORG || "southbendin";
const PROJECT = process.env.DEVOPS_PROJECT || "Digital%20-%20Product%20Portfolio";
const PAT = process.env.DEVOPS_PAT || "";

const BASE = `https://dev.azure.com/${ORG}`;
const PROJECT_BASE = `${BASE}/${PROJECT}`;
const AUTH = "Basic " + btoa(":" + PAT);

function headers(contentType) {
  return { "Content-Type": contentType, Authorization: AUTH };
}

function assertJson(res) {
  const ct = (res.headers.get("content-type") || "").toLowerCase();
  if (!ct.includes("application/json")) {
    throw new Error(
      "Expected JSON but received HTML. Your PAT token is likely missing, invalid, or expired."
    );
  }
}

async function apiGet(url) {
  const res = await fetch(url, { method: "GET", headers: headers("application/json") });
  if (!res.ok) throw new Error(`GET ${url} failed: ${res.status} ${res.statusText}`);
  assertJson(res);
  return res.json();
}

async function apiPost(url, body, contentType = "application/json") {
  const res = await fetch(url, {
    method: "POST",
    headers: headers(contentType),
    body: typeof body === "string" ? body : JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`POST ${url} failed: ${res.status} ${res.statusText}`);
  assertJson(res);
  return res.json();
}

async function apiPatch(url, body) {
  const res = await fetch(url, {
    method: "PATCH",
    headers: headers("application/json-patch+json"),
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`PATCH ${url} failed: ${res.status} ${res.statusText}`);
  assertJson(res);
  return res.json();
}

function escapeWiql(text) {
  return text.replace(/'/g, "''");
}

export async function queryWorkItems(wiqlQuery) {
  return apiPost(`${PROJECT_BASE}/_apis/wit/wiql?api-version=6.0`, { query: wiqlQuery });
}

export async function getWorkItemsByIds(ids) {
  if (!ids || ids.length === 0) return [];
  const idStr = ids.join(",");
  const data = await apiGet(`${BASE}/_apis/wit/workitems?ids=${idStr}&$expand=all&api-version=6.0`);
  return data.value || [];
}

export async function getWorkItemById(id) {
  return apiGet(`${BASE}/_apis/wit/workitems/${id}?$expand=all&api-version=6.0`);
}

export async function createWorkItem(type, patchOps) {
  return apiPost(
    `${PROJECT_BASE}/_apis/wit/workitems/$${encodeURIComponent(type)}?bypassrules=true&api-version=6.0`,
    patchOps,
    "application/json-patch+json"
  );
}

export async function uploadAttachment(fileName, blob) {
  const url = `${PROJECT_BASE}/_apis/wit/attachments?fileName=${encodeURIComponent(fileName)}&api-version=6.0`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", Authorization: AUTH },
    body: blob,
  });
  if (!res.ok) throw new Error(`Upload attachment failed: ${res.status}`);
  return res.json();
}

export async function addAttachmentToWorkItem(workItemId, attachUrl, bodyHtml, bodyField) {
  const field = bodyField || "System.Description";
  const ops = [
    {
      op: "add",
      path: "/relations/-",
      value: {
        rel: "AttachedFile",
        url: attachUrl,
        attributes: { comment: "Created with Outlook Azure add-in" },
      },
    },
    { op: "add", path: `/fields/${field}`, value: bodyHtml },
  ];
  return apiPatch(`${BASE}/_apis/wit/workitems/${workItemId}?api-version=6.0`, ops);
}

export async function updateWorkItemField(workItemId, field, value) {
  return apiPatch(`${BASE}/_apis/wit/workitems/${workItemId}?api-version=6.0`, [
    { op: "add", path: `/fields/${field}`, value: value },
  ]);
}

export async function setCurrentIteration(workItemId) {
  const data = await apiGet(
    `${PROJECT_BASE}/_apis/work/teamsettings/iterations?$timeframe=current&api-version=6.0`
  );
  if (!data.value || data.value.length === 0) return null;
  const path = data.value[0].path;
  return apiPatch(`${BASE}/_apis/wit/workitems/${workItemId}?api-version=6.0`, [
    { op: "add", path: "/fields/System.IterationPath", from: null, value: path },
  ]);
}

export async function getGraphUsers() {
  const data = await apiGet(
    `https://vssps.dev.azure.com/${ORG}/_apis/graph/users?api-version=6.0-preview.1`
  );
  return data.value || [];
}

export function getWorkItemWebUrl(id) {
  const project = decodeURIComponent(PROJECT);
  return `https://dev.azure.com/${ORG}/${encodeURIComponent(project)}/_workitems/edit/${id}`;
}

export function getProjectName() {
  return decodeURIComponent(PROJECT);
}

export async function searchWorkItems(searchText, typeFilter) {
  const trimmed = (searchText || "").trim();
  if (!trimmed) return [];

  let whereClause;
  const isId = /^\d+$/.test(trimmed);

  if (isId) {
    whereClause = `[System.Id] = ${trimmed}`;
  } else {
    whereClause = `[System.Title] CONTAINS '${escapeWiql(trimmed)}'`;
  }

  if (typeFilter) {
    whereClause += ` AND [System.WorkItemType] = '${escapeWiql(typeFilter)}'`;
  }

  const wiql = `SELECT [System.Id], [System.Title], [System.State], [System.WorkItemType], [System.AssignedTo] FROM WorkItems WHERE ${whereClause} ORDER BY [System.ChangedDate] DESC`;
  const result = await queryWorkItems(wiql);
  const ids = (result.workItems || []).slice(0, 50).map((w) => w.id);
  if (ids.length === 0) return [];
  return getWorkItemsByIds(ids);
}

export async function getMaintenanceItems() {
  const wiql = `SELECT [System.Id], [System.Title], [System.State] FROM WorkItems WHERE [System.WorkItemType] = 'Maintenance' AND [System.State] <> 'Archived'`;
  const result = await queryWorkItems(wiql);
  const ids = (result.workItems || []).map((w) => w.id);
  if (ids.length === 0) return [];
  return getWorkItemsByIds(ids);
}

const PARENT_TYPES = ["Request", "Issue", "Project", "Requirement", "Detail", "Planned Work", "Maintenance", "Product"];

export function isAllowedParentType(type) {
  return PARENT_TYPES.includes(type);
}

export async function searchParentItems(searchText) {
  const trimmed = (searchText || "").trim();
  if (!trimmed) return [];

  const isId = /^\d+$/.test(trimmed);
  let whereClause;

  if (isId) {
    whereClause = `[System.Id] = ${trimmed}`;
  } else {
    whereClause = `[System.Title] CONTAINS '${escapeWiql(trimmed)}'`;
  }

  const typeIn = PARENT_TYPES.map((t) => `'${t}'`).join(", ");
  whereClause += ` AND [System.WorkItemType] IN (${typeIn})`;

  const wiql = `SELECT [System.Id], [System.Title], [System.State], [System.WorkItemType], [System.AssignedTo] FROM WorkItems WHERE ${whereClause} ORDER BY [System.ChangedDate] DESC`;
  const result = await queryWorkItems(wiql);
  const ids = (result.workItems || []).slice(0, 20).map((w) => w.id);
  if (ids.length === 0) return [];
  return getWorkItemsByIds(ids);
}
