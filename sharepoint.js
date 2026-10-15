// Load Orders list access -- reads/writes the orders people enter through
// this app's "+ New Order", as opposed to the ORDERS_LIST_2026.xlsx
// spreadsheet (which stays read-only, synced in by scripts/sync_orders.py).
//
// Column-name handling: lists created via SharePoint's "New list > From
// Excel" import don't reliably keep a column's internal (Graph) name
// matching its display name -- imports can land as "field_1", "field_2",
// etc even when the Excel header had no spaces. So every read/write here
// goes through resolveColumnMap(), which looks the list's real columns up
// by DISPLAY name once and caches the display->internal mapping, rather
// than assuming internal names ever match what's in
// setup/Load_Orders_list_template.xlsx.

import { graphFetch, resolveSite, findList } from "./auth.js";

const cfg = (typeof window !== "undefined" && window.SALLEE_LOAD_BOARD_CONFIG) || {};
const LIST_NAME = cfg.listName || "Load Orders";

// our field key -> the exact column DISPLAY name in
// setup/Load_Orders_list_template.xlsx / LOAD_ORDERS_SETUP.md.
const FIELD_DISPLAY_NAMES = {
  requestedBy: "RequestedBy",
  requestedDates: "RequestedDates",
  dateTaken: "DateTaken",
  takenBy: "TakenBy",
  origin: "Origin",
  originTrack: "OriginTrack",
  originFarm: "OriginFarm",
  originTrainer: "OriginTrainer",
  destination: "Destination",
  track: "DestinationTrack",
  destinationTrainer: "DestinationTrainer",
  destinationFarm: "DestinationFarm",
  horseNames: "HorseNames",
  stallSpace: "StallSpace",
  notes: "Notes",
  tripDate: "TripDate",
  drivers: "Drivers",
  status: "Status",
};

let listIdPromise = null;
function getListId() {
  if (!listIdPromise) listIdPromise = findList(LIST_NAME);
  return listIdPromise;
}

let columnMapPromise = null;
async function resolveColumnMap() {
  if (columnMapPromise) return columnMapPromise;
  columnMapPromise = (async () => {
    const listId = await getListId();
    const sid = await resolveSite();
    const resp = await graphFetch(`/sites/${sid}/lists/${listId}/columns`);
    const byDisplayName = {};
    for (const col of resp.value || []) {
      byDisplayName[col.displayName] = col.name; // col.name = internal/Graph field name
    }
    const map = {};
    const missing = [];
    for (const [ourKey, displayName] of Object.entries(FIELD_DISPLAY_NAMES)) {
      if (byDisplayName[displayName]) {
        map[ourKey] = byDisplayName[displayName];
      } else {
        missing.push(displayName);
      }
    }
    if (missing.length) {
      console.warn(
        `Load Orders list is missing expected column(s): ${missing.join(", ")}. ` +
          `Double-check setup/Load_Orders_list_template.xlsx was imported with every column intact.`
      );
    }
    return map;
  })();
  return columnMapPromise;
}

function splitList(raw) {
  if (raw === null || raw === undefined || raw === "") return [];
  return String(raw)
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s !== "");
}

// Mirrors sync_orders.py's parse_horses(): pairs HorseNames / StallSpace up
// by position. An empty names list with stall values still produces that
// many unnamed horses, matching the spreadsheet's own fallback behavior.
function parseHorses(namesRaw, stallsRaw) {
  const names = splitList(namesRaw);
  const stallStrs = splitList(stallsRaw);
  if (!names.length && !stallStrs.length) return [];
  const stalls = stallStrs.map((s) => {
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  });
  const count = Math.max(names.length, stalls.length);
  const horses = [];
  for (let i = 0; i < count; i++) {
    horses.push({
      name: names[i] || null,
      stallSpace: i < stalls.length ? stalls[i] : null,
    });
  }
  return horses;
}

// Reverse of parseHorses(): our horses[] -> paired comma-separated strings,
// same comma-separated format as the ORDERS spreadsheet's own
// HORSE NAME(S) / STALL SPACE columns.
function joinHorses(horses) {
  const list = horses || [];
  return {
    names: list.map((h) => h.name || "").join(", "),
    stalls: list.map((h) => (h.stallSpace != null && h.stallSpace !== "" ? h.stallSpace : "")).join(", "),
  };
}

function orderToFields(order, colMap) {
  const { names, stalls } = joinHorses(order.horses);
  const raw = {
    requestedBy: order.requestedBy || "",
    requestedDates: order.requestedDates || "",
    dateTaken: order.dateTaken || "",
    takenBy: order.takenBy || "",
    origin: order.origin || "",
    originTrack: order.originTrack || "",
    originFarm: order.originFarm || "",
    originTrainer: order.originTrainer || "",
    destination: order.destination || "",
    track: order.track || "",
    destinationTrainer: order.destinationTrainer || "",
    destinationFarm: order.destinationFarm || "",
    horseNames: names,
    stallSpace: stalls,
    notes: order.notes || "",
    tripDate: order.tripDate || "",
    drivers: order.drivers || "",
    status: order.status || "Open",
  };
  const fields = {
    // Title always exists on every SharePoint list under its standard
    // internal name "Title" -- no need to route it through the display-name
    // lookup. Purely so a human browsing the raw list sees a readable lane
    // instead of "(no title)"; the app itself never reads Title back.
    Title: `${order.origin || "?"} -> ${order.destination || "?"}`.slice(0, 255),
  };
  for (const [ourKey, val] of Object.entries(raw)) {
    const internalName = colMap[ourKey];
    if (internalName) fields[internalName] = val;
  }
  return fields;
}

function fieldsToOrder(item, colMap) {
  const f = item.fields || {};
  const get = (ourKey) => {
    const internalName = colMap[ourKey];
    return internalName ? f[internalName] : undefined;
  };
  const horsesParsed = parseHorses(get("horseNames"), get("stallSpace"));
  const horses = horsesParsed.length ? horsesParsed : [{ name: null, stallSpace: null }];
  return {
    id: `sp-${item.id}`,
    _spItemId: item.id,
    source: "app",
    status: get("status") || "Open",
    requestedBy: get("requestedBy") || "",
    requestedDates: get("requestedDates") || "",
    dateTaken: get("dateTaken") || "",
    takenBy: get("takenBy") || "",
    origin: get("origin") || "",
    originTrack: get("originTrack") || "",
    originFarm: get("originFarm") || "",
    originTrainer: get("originTrainer") || "",
    destination: get("destination") || "",
    track: get("track") || "",
    destinationTrainer: get("destinationTrainer") || "",
    destinationFarm: get("destinationFarm") || "",
    notes: get("notes") || "",
    tripDate: get("tripDate") || "",
    drivers: get("drivers") || "",
    horses,
    stallSpace: horses.reduce((sum, h) => sum + (Number(h.stallSpace) || 0), 0),
  };
}

// Returns every OPEN order currently on the Load Orders list, already
// mapped to the same order shape the board uses for spreadsheet orders
// (so StallStrip / LoadModal / etc need no special-casing to render them).
export async function fetchAppOrders() {
  const listId = await getListId();
  const colMap = await resolveColumnMap();
  const sid = await resolveSite();
  const results = [];
  let url = `/sites/${sid}/lists/${listId}/items?expand=fields&$top=200`;
  while (url) {
    const page = await graphFetch(url);
    results.push(...(page.value || []));
    url = page["@odata.nextLink"] || null;
  }
  return results.map((item) => fieldsToOrder(item, colMap)).filter((o) => (o.status || "Open") === "Open");
}

export async function createAppOrder(orderDraft) {
  const listId = await getListId();
  const colMap = await resolveColumnMap();
  const sid = await resolveSite();
  const fields = orderToFields({ ...orderDraft, status: "Open" }, colMap);
  const created = await graphFetch(`/sites/${sid}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields }),
  });
  return fieldsToOrder(created, colMap);
}

// Partial update -- pass only the our-field-key(s) that changed, e.g.
// { tripDate: "2026-11-02", drivers: "BEN" } or { status: "Completed" }.
// A `horses` key is accepted too and gets split into horseNames/stallSpace
// the same way createAppOrder does.
export async function updateAppOrder(spItemId, patch) {
  const listId = await getListId();
  const colMap = await resolveColumnMap();
  const sid = await resolveSite();

  const rawPatch = { ...patch };
  if (rawPatch.horses) {
    const { names, stalls } = joinHorses(rawPatch.horses);
    rawPatch.horseNames = names;
    rawPatch.stallSpace = stalls;
    delete rawPatch.horses;
  }

  const fieldPatch = {};
  for (const [ourKey, val] of Object.entries(rawPatch)) {
    const internalName = colMap[ourKey];
    if (internalName && val !== undefined) fieldPatch[internalName] = val;
  }

  await graphFetch(`/sites/${sid}/lists/${listId}/items/${spItemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fieldPatch),
  });
}

export function markCompleted(spItemId) {
  return updateAppOrder(spItemId, { status: "Completed" });
}
export function markCanceled(spItemId) {
  return updateAppOrder(spItemId, { status: "Canceled" });
}
