import React, { useState, useEffect, useMemo, useRef } from "react";

/* ============================================================
   SALLEE LOAD BOARD — live departures-style board for
   inbound shipping requests, grouped into loads like a
   horse van's stall map.

   LIVE DATA: this component polls ORDERS_JSON_URL (see below)
   for data/orders.json, which the sync_orders.py script + the
   sync-orders.yml GitHub Action keep up to date from the
   SharePoint-hosted ORDERS worksheet. Every fetch is diffed
   against the previous one, so the ticker and the row
   highlight only fire for rows that are genuinely new on the
   sheet.

   If that fetch fails (e.g. previewing this file standalone,
   or the JSON isn't deployed alongside it yet), it falls back
   to SNAPSHOT_ORDERS below — a one-time pull from
   automated_load_build_up_4.xlsm — so the board still renders
   something sensible.

   MANUALLY BUILT LOADS: dispatchers can also hand-pick requests
   into a load with "+ Build a Load", independent of the
   auto-grouped board above. Those live only in this browser's
   localStorage (CREATED_LOADS_STORAGE_KEY) — they never touch
   the ORDERS sheet. Any order used in a manually built load is
   pulled out of the auto-grouped pool so it doesn't show up
   twice.
   ============================================================ */

const ORDERS_JSON_URL = "/data/orders.json"; // adjust if you deploy the JSON somewhere else
const POLL_INTERVAL_MS = 30000; // how often to check SharePoint for changes
const CREATED_LOADS_STORAGE_KEY = "sallee-load-board:created-loads:v1";

function loadCreatedLoads() {
  try {
    const raw = localStorage.getItem(CREATED_LOADS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveCreatedLoads(loads) {
  try {
    localStorage.setItem(CREATED_LOADS_STORAGE_KEY, JSON.stringify(loads));
  } catch {
    // ignore write failures (e.g. private browsing / storage full)
  }
}

function generateId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

const SNAPSHOT_ORDERS = [{"id": 2, "dateTaken": "2024-07-31", "requestedDates": "WEEK OF 8/5", "takenBy": "KK", "requestedBy": "NANCY", "trainerFarm": "PATRICK DIXON", "origin": "WOODBINE", "destination": "KENTUCKY", "track": null, "farmTrainer": "KENNEALLY", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 3, "dateTaken": "2024-07-08", "requestedDates": "8/28- 8/31", "takenBy": "LS", "requestedBy": "CARLOS", "trainerFarm": "U ORTEGA", "origin": "OCALA", "destination": "KENTUCKY", "track": "CHURCHILL", "farmTrainer": "I. QUESADA", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 4, "dateTaken": "2024-07-15", "requestedDates": "2024-08-20", "takenBy": "MR", "requestedBy": "CARRIE", "trainerFarm": null, "origin": "SARATOGA", "destination": "SOUTH CAROLINA", "track": null, "farmTrainer": null, "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 5, "dateTaken": "2024-07-24", "requestedDates": "NEXT AVAILABLE", "takenBy": "MD", "requestedBy": "DEAN", "trainerFarm": null, "origin": "EAST COAST", "destination": "SOUTH CAROLINA", "track": null, "farmTrainer": null, "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 6, "dateTaken": "2024-07-16", "requestedDates": "EARLY SEPT", "takenBy": "FS", "requestedBy": null, "trainerFarm": "AUDREY HICKS", "origin": "SARATOGA", "destination": "SOUTH CAROLINA", "track": null, "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 7, "dateTaken": "2024-08-01", "requestedDates": "NEXT AVAILABLE", "takenBy": "AB", "requestedBy": "SANDY", "trainerFarm": "NORM CASSE", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "PARX", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 8, "dateTaken": "2024-04-01", "requestedDates": "MID AUGUST", "takenBy": "AB", "requestedBy": "BAILEE", "trainerFarm": null, "origin": "KENTUCKY", "destination": "UPSTATE NEW YORK", "track": "UPSTATE NEW YORK", "farmTrainer": "STONEBRIDGE", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 9, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "BELMONT", "farmTrainer": "BROWN", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 10, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "LAUREL", "farmTrainer": "RUSSEL", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 11, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "SARATOGA", "track": "SARATOGA", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 12, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "UPSTATE NEW YORK", "track": "FARM", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 13, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "LAUREL", "farmTrainer": "MAKER", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 14, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "FAIR HILL", "farmTrainer": "JACKSON", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 15, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "DELAWARE", "farmTrainer": "M.DINI", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 16, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "MONMOUTH", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 17, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "EAST COAST", "track": "PARX", "farmTrainer": "LAKE", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 18, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "KENTUCKY", "track": "FARM", "farmTrainer": "STONE", "stallSpace": 1.5, "tripDate": null, "drivers": null}, {"id": 19, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "KENTUCKY", "track": "KEENELAND", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 20, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "SOUTH CAROLINA", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 21, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "OCALA", "destination": "KENTUCKY", "track": "TURFWAY", "farmTrainer": "MAKER", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 22, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "FAIR HILL", "farmTrainer": "JACKSON", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 23, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "SFL", "track": "PAYSON", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 24, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "OCALA", "track": "FARM", "farmTrainer": "KINSMAN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 25, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "SOUTH CAROLINA", "track": "FARM", "farmTrainer": "DURR", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 26, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "UPSTATE NEW YORK", "track": "FARM", "farmTrainer": "STONEBRIDGE", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 27, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "MICHIGAN", "track": "FARM", "farmTrainer": "DJ JOHNSON", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 28, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "PARX", "farmTrainer": "RODRIGUEZ", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 29, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "COLONIAL DOWNS", "farmTrainer": "CASSE", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 30, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "SARATOGA", "track": "TRACK", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 31, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "MONMOUTH", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 32, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "BELMONT", "track": "TRACK", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 33, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "DELAWARE", "farmTrainer": "DINI", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 34, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "FAIRGROUNDS", "track": "TRACK", "farmTrainer": "STIDHAM", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 35, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "SFL", "track": "GULF", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 36, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "SFL", "track": "PMTC", "farmTrainer": "BROWN", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 37, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "KENTUCKY", "destination": "EAST COAST", "track": "LAUREL", "farmTrainer": "B.RUSSEL", "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 44, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "KENTUCKY", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 45, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "KENTUCKY", "track": "CHURCHILL", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 46, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "KENTUCKY", "track": "KEENELAND", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 47, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "OCALA", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 48, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "PARX", "track": "TRACK", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 49, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "BELMONT", "track": "TRACK", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 50, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "FAIR HILL", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 51, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "UPSTATE NEW YORK", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 52, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "SOUTH CAROLINA", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 53, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "SARATOGA", "track": "TRACK", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 54, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "FAIRGROUNDS", "track": "TRACK", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 55, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "PARX", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 56, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "LAUREL", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 57, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "FAIR HILL", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 60, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "MONMOUTH", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 62, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 63, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}, {"id": 64, "dateTaken": "2024-08-08", "requestedDates": "NEXT AVAILABLE", "takenBy": "ME", "requestedBy": "ME", "trainerFarm": "WHAT EVER", "origin": "SFL", "destination": "EAST COAST", "track": "FARM", "farmTrainer": null, "stallSpace": 3, "tripDate": null, "drivers": null}]

const DEFAULT_CAPACITY = 21; // stalls per load — matches how loads are built in ORDERS
const STALL_UNIT = 1.5; // every request is 1.5 or 3 stalls, so the strip is drawn in 1.5-stall bays

/* ---------- helpers ---------- */

function loadKey(o) {
  return [o.origin, o.originTrack || "—", o.destination, o.track || "—"].join("|");
}

function buildLoads(orders) {
  const map = new Map();
  for (const o of orders) {
    const key = loadKey(o);
    if (!map.has(key)) {
      map.set(key, {
        key,
        origin: o.origin,
        originTrack: o.originTrack,
        destination: o.destination,
        track: o.track,
        orders: [],
      });
    }
    map.get(key).orders.push(o);
  }
  return Array.from(map.values()).sort((a, b) => {
    if (a.origin !== b.origin) return a.origin.localeCompare(b.origin);
    if (a.destination !== b.destination) return a.destination.localeCompare(b.destination);
    return (a.track || "").localeCompare(b.track || "");
  });
}

/* A load's orders can carry different requested dates. Surface that
   plainly instead of silently picking one, so a dispatcher building a
   load can catch a mismatch before it becomes a problem on the road. */
function requestedDateSummary(load) {
  const dates = Array.from(
    new Set(load.orders.map((o) => o.requestedDates || "date flexible"))
  );
  if (dates.length === 1) return { text: dates[0], mismatched: false };
  return { text: "MULTIPLE DATES", mismatched: true };
}

function stallsBooked(load) {
  return load.orders.reduce((sum, o) => sum + (Number(o.stallSpace) || 0), 0);
}

function trainerFarmOf(o) {
  return o.trainerFarm || null;
}

function farmTrainerOf(o) {
  return o.farmTrainer || null;
}

/* Combines an order's two "who" fields into one readable name for
   places that show a single order (the load detail manifest, the
   ticker) — falls back gracefully if only one side is filled in. */
function combinedName(o) {
  const a = trainerFarmOf(o);
  const b = farmTrainerOf(o);
  return a && b ? `${a} → ${b}` : a || b || "Unassigned trainer/farm";
}

function uniqueList(items, getter) {
  const out = [];
  for (const item of items) {
    const v = getter(item);
    if (v && !out.includes(v)) out.push(v);
  }
  return out;
}

/* "BROWN, KENNEALLY +2 more" — used to summarize several orders'
   worth of names into one line without the row growing unbounded. */
function joinWithMore(list, max = 2) {
  if (list.length === 0) return null;
  if (list.length <= max) return list.join(", ");
  return `${list.slice(0, max).join(", ")} +${list.length - max} more`;
}

function statusFor(booked, capacity) {
  if (booked >= capacity) return booked > capacity ? "OVERBOOKED" : "FULL";
  if (booked >= capacity * 0.7) return "FILLING";
  return "OPEN";
}

const STATUS_STYLE = {
  OPEN: { color: "#7be0a8", label: "OPEN" },
  FILLING: { color: "#e8c96a", label: "FILLING" },
  FULL: { color: "#ff9d5c", label: "FULLY BOOKED" },
  OVERBOOKED: { color: "#ff6b6b", label: "OVERBOOKED" },
};

/* ---------- split flap character ---------- */

function Flap({ ch }) {
  return (
    <span
      style={{
        display: "inline-block",
        minWidth: "0.62em",
        textAlign: "center",
      }}
    >
      {ch}
    </span>
  );
}

function FlapText({ text, size = 15, color = "#f0c95a", weight = 700, spacing = 1 }) {
  return (
    <span
      style={{
        fontFamily:
          "ui-monospace, SFMono-Regular, 'Courier New', Courier, monospace",
        fontSize: size,
        fontWeight: weight,
        letterSpacing: spacing,
        color,
        whiteSpace: "pre",
      }}
    >
      {text.split("").map((c, i) => (
        <Flap key={i} ch={c} />
      ))}
    </span>
  );
}

/* ---------- stall strip (seat map) ---------- */

function formatStalls(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(1);
}

function StallStrip({ load, capacity, onCapacityChange }) {
  const orders = load.orders;
  const totalUnits = Math.round(capacity / STALL_UNIT);
  let cursor = 0;
  const placed = [];
  const overflow = [];

  for (const o of orders) {
    const span = Math.max(1, Math.round((Number(o.stallSpace) || 0) / STALL_UNIT));
    if (cursor >= totalUnits) {
      overflow.push(o);
      continue;
    }
    const start = cursor;
    const end = Math.min(totalUnits, cursor + span);
    placed.push({ order: o, start, end, span: end - start });
    cursor = end;
  }

  const openUnits = Math.max(0, totalUnits - cursor);
  const bookedStalls = stallsBooked(load);
  const palette = [
    "#3fa66b", "#3f8fa6", "#a68f3f", "#8a5fd6", "#d65f9a", "#5fd6c6", "#d68a5f",
  ];

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <div style={{ fontSize: 12, letterSpacing: 1, color: "#9fb8a8", fontFamily: "ui-monospace, monospace" }}>
          TRAILER CAPACITY
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button
            onClick={() => onCapacityChange(Math.max(STALL_UNIT, capacity - STALL_UNIT))}
            style={btnSmall}
          >
            −
          </button>
          <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, color: "#f0c95a", minWidth: 70, textAlign: "center" }}>
            {formatStalls(capacity)} stalls
          </div>
          <button
            onClick={() => onCapacityChange(capacity + STALL_UNIT)}
            style={btnSmall}
          >
            +
          </button>
        </div>
      </div>

      <div
        style={{
          display: "flex",
          gap: 4,
          padding: 8,
          background: "#07130d",
          borderRadius: 10,
          border: "1px solid #1d3527",
          flexWrap: "wrap",
        }}
      >
        {placed.map((p, idx) => {
          const color = palette[idx % palette.length];
          const name = p.order.trainerFarm || p.order.farmTrainer || "Unassigned";
          return (
            <div
              key={p.order.id ?? idx}
              title={`${name} — ${formatStalls(p.order.stallSpace)} stall(s)`}
              style={{
                flex: `${p.span} 1 0`,
                minWidth: p.span * 46,
                height: 50,
                borderRadius: 6,
                background: color,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                padding: "0 6px",
                overflow: "hidden",
              }}
            >
              <span
                style={{
                  fontSize: 12,
                  fontWeight: 800,
                  color: "#06110b",
                  whiteSpace: "nowrap",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  maxWidth: "100%",
                }}
              >
                {name}
              </span>
              <span style={{ fontSize: 10, color: "#0a2417", fontWeight: 700, marginTop: 1 }}>
                {formatStalls(p.order.stallSpace)} stall{p.order.stallSpace === 1 ? "" : "s"}
              </span>
            </div>
          );
        })}
        {Array.from({ length: openUnits }).map((_, i) => (
          <div
            key={`open-${i}`}
            style={{
              flex: "1 1 0",
              minWidth: 46,
              height: 50,
              borderRadius: 6,
              border: "1.5px dashed #2c4a3a",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 10,
              color: "#4a6a58",
              fontWeight: 700,
              letterSpacing: 0.5,
            }}
          >
            OPEN
          </div>
        ))}
      </div>

      <div style={{ marginTop: 6, fontSize: 12, color: "#7f9c8c", fontFamily: "ui-monospace, monospace" }}>
        {formatStalls(bookedStalls)} of {formatStalls(capacity)} stalls booked · {formatStalls(openUnits * STALL_UNIT)} open
      </div>

      {overflow.length > 0 && (
        <div
          style={{
            marginTop: 14,
            padding: "12px 14px",
            background: "rgba(255,107,107,0.1)",
            border: "1px solid #ff6b6b",
            borderRadius: 10,
          }}
        >
          <div style={{ fontSize: 12, fontWeight: 800, color: "#ff6b6b", letterSpacing: 1 }}>
            {overflow.length} REQUEST{overflow.length === 1 ? "" : "S"} OVER CAPACITY
          </div>
          <div style={{ fontSize: 12, color: "#e0a3a3", marginTop: 4 }}>
            These don't fit on this trailer as sized — raise the capacity above, or plan a second load.
          </div>
        </div>
      )}
    </div>
  );
}

const btnSmall = {
  width: 28,
  height: 28,
  borderRadius: 6,
  border: "1px solid #2c4a3a",
  background: "#0f2419",
  color: "#e8c96a",
  fontWeight: 800,
  fontSize: 16,
  cursor: "pointer",
  lineHeight: "1",
};

/* ---------- load detail modal (for auto-grouped loads) ---------- */

function LoadModal({ load, capacity, onCapacityChange, onClose }) {
  const booked = stallsBooked(load);
  const status = statusFor(booked, capacity);
  const s = STATUS_STYLE[status];
  const dateGroups = new Set(load.orders.map((o) => o.requestedDates || "date flexible"));
  const datesMismatched = dateGroups.size > 1;

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(4,10,7,0.72)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 50,
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%",
          maxWidth: 640,
          maxHeight: "88vh",
          overflowY: "auto",
          background: "#0b1c14",
          border: "1px solid #22402f",
          borderRadius: 16,
          boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
        }}
      >
        <div style={{ padding: "22px 26px", borderBottom: "1px solid #1d3527" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div>
              <div style={{ fontSize: 11, letterSpacing: 3, color: "#7f9c8c", fontFamily: "ui-monospace, monospace" }}>
                LOAD DETAIL
              </div>
              <div style={{ display: "flex", alignItems: "baseline", gap: 12, marginTop: 4 }}>
                <FlapText text={load.origin} size={26} />
                <span style={{ color: "#4a6a58", fontSize: 22 }}>→</span>
                <FlapText text={load.destination} size={26} />
              </div>
              <div style={{ display: "flex", gap: 16, marginTop: 4, fontSize: 13, color: "#9fb8a8" }}>
                {load.originTrack && <span>via {load.originTrack}</span>}
                {load.track && <span>via {load.track}</span>}
              </div>
            </div>
            <button onClick={onClose} style={{ ...btnSmall, width: 32, height: 32, fontSize: 18 }}>
              ×
            </button>
          </div>
          <div style={{ marginTop: 12, display: "inline-block", padding: "4px 12px", borderRadius: 999, border: `1px solid ${s.color}`, color: s.color, fontSize: 12, fontWeight: 800, letterSpacing: 1 }}>
            {s.label}
          </div>
        </div>

        <div style={{ padding: "22px 26px" }}>
          <StallStrip load={load} capacity={capacity} onCapacityChange={onCapacityChange} />

          <div style={{ marginTop: 22, fontSize: 11, letterSpacing: 2, color: "#7f9c8c", fontFamily: "ui-monospace, monospace", marginBottom: 10 }}>
            PASSENGER MANIFEST — {load.orders.length} {load.orders.length === 1 ? "REQUEST" : "REQUESTS"}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {load.orders.map((o) => {
              const orderDate = o.requestedDates || "date flexible";
              return (
                <div
                  key={o.id}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "10px 14px",
                    background: "#0f2419",
                    borderRadius: 10,
                    border: datesMismatched ? "1px solid #6b4a1f" : "1px solid #1d3527",
                  }}
                >
                  <div>
                    <div style={{ fontWeight: 700, color: "#eaf3ec", fontSize: 14 }}>
                      {combinedName(o)}
                    </div>
                    <div style={{ fontSize: 12, color: "#8fae9c", marginTop: 2 }}>
                      Requested by {o.requestedBy || "—"} ·{" "}
                      <span style={{ fontWeight: 800, color: datesMismatched ? "#f0a95a" : "#c9dcd0", fontSize: 13 }}>
                        {orderDate}
                      </span>
                    </div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "ui-monospace, monospace", color: "#f0c95a", fontWeight: 800 }}>
                      {o.stallSpace} stall{o.stallSpace === 1 ? "" : "s"}
                    </div>
                    <div style={{ fontSize: 11, color: "#6f8c7c" }}>taken by {o.takenBy || "—"}</div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------- build-a-load modal (manual load creation/editing) ---------- */

const labelStyle = {
  fontSize: 10,
  letterSpacing: 2,
  color: "#6f8c7c",
  fontWeight: 700,
  fontFamily: "ui-monospace, monospace",
};

const selectStyle = {
  background: "#0f2419",
  border: "1px solid #2c4a3a",
  color: "#eaf3ec",
  borderRadius: 8,
  padding: "8px 10px",
  fontSize: 13,
  fontFamily: "ui-monospace, monospace",
  outline: "none",
};

function BuildLoadModal({ pool, origins, destinations, initial, nextLoadNumber, onSave, onDelete, onClose }) {
  const editing = !!initial;
  const [origin, setOrigin] = useState(initial?.origin || "");
  const [destination, setDestination] = useState(initial?.destination || "");
  const [selectedIds, setSelectedIds] = useState(() => new Set(initial?.orderIds || []));
  const [label, setLabel] = useState(initial?.label || `Load #${nextLoadNumber}`);
  const [driver, setDriver] = useState(initial?.driver || "");
  const [tripDate, setTripDate] = useState(initial?.tripDate || "");
  const [capacity, setCapacity] = useState(initial?.capacity || DEFAULT_CAPACITY);

  const selectedOrders = useMemo(
    () => pool.filter((o) => selectedIds.has(o.id)),
    [pool, selectedIds]
  );
  const previewLoad = { orders: selectedOrders };
  const booked = stallsBooked(previewLoad);
  const hasLane = !!origin && !!destination;
  const laneMatches = useMemo(
    () => (hasLane ? pool.filter((o) => o.origin === origin && o.destination === destination) : []),
    [pool, origin, destination, hasLane]
  );

  function toggleOrder(id) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function autoFillToCapacity() {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      let total = selectedOrders.reduce((sum, o) => sum + (Number(o.stallSpace) || 0), 0);
      for (const o of laneMatches) {
        if (next.has(o.id)) continue;
        const stall = Number(o.stallSpace) || 0;
        if (total + stall > capacity) continue;
        next.add(o.id);
        total += stall;
      }
      return next;
    });
  }

  function handleSave() {
    if (selectedIds.size === 0) return;
    onSave({
      id: initial?.id || generateId(),
      label: label.trim() || `Load #${nextLoadNumber}`,
      origin: origin || selectedOrders[0]?.origin || "",
      destination: destination || selectedOrders[0]?.destination || "",
      driver: driver.trim(),
      tripDate: tripDate.trim(),
      capacity,
      orderIds: Array.from(selectedIds),
    });
  }

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(4,10,7,0.72)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 60,
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%",
          maxWidth: 680,
          maxHeight: "90vh",
          overflowY: "auto",
          background: "#0b1c14",
          border: "1px solid #22402f",
          borderRadius: 16,
          boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
        }}
      >
        <div style={{ padding: "22px 26px", borderBottom: "1px solid #1d3527" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div>
              <div style={{ fontSize: 11, letterSpacing: 3, color: "#7f9c8c", fontFamily: "ui-monospace, monospace" }}>
                {editing ? "EDIT LOAD" : "BUILD A LOAD"}
              </div>
              <div style={{ fontSize: 22, fontWeight: 900, marginTop: 4 }}>
                {label || "Untitled load"}
              </div>
              <div style={{ fontSize: 12, color: "#7f9c8c", marginTop: 4 }}>
                This never touches the ORDERS sheet — it's just a local grouping in this browser.
              </div>
            </div>
            <button onClick={onClose} style={{ ...btnSmall, width: 32, height: 32, fontSize: 18 }}>
              ×
            </button>
          </div>
        </div>

        <div style={{ padding: "22px 26px" }}>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 18 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={labelStyle}>ORIGIN</label>
              <select value={origin} onChange={(e) => setOrigin(e.target.value)} style={selectStyle}>
                <option value="">Choose origin…</option>
                {origins.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))}
              </select>
            </div>
            <div style={{ fontSize: 18, color: "#3f5f4d", marginTop: 18 }}>→</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={labelStyle}>DESTINATION</label>
              <select value={destination} onChange={(e) => setDestination(e.target.value)} style={selectStyle}>
                <option value="">Choose destination…</option>
                {destinations.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
            </div>
          </div>

          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 22 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: "1 1 140px" }}>
              <label style={labelStyle}>LOAD NAME</label>
              <input value={label} onChange={(e) => setLabel(e.target.value)} style={selectStyle} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: "1 1 140px" }}>
              <label style={labelStyle}>DRIVER (OPTIONAL)</label>
              <input value={driver} onChange={(e) => setDriver(e.target.value)} placeholder="—" style={selectStyle} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: "1 1 140px" }}>
              <label style={labelStyle}>TRIP DATE (OPTIONAL)</label>
              <input value={tripDate} onChange={(e) => setTripDate(e.target.value)} placeholder="—" style={selectStyle} />
            </div>
          </div>

          <StallStrip load={previewLoad} capacity={capacity} onCapacityChange={setCapacity} />

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 22, marginBottom: 10 }}>
            <div style={{ fontSize: 11, letterSpacing: 2, color: "#7f9c8c", fontFamily: "ui-monospace, monospace" }}>
              {hasLane ? `OPEN REQUESTS ON THIS LANE — ${laneMatches.length}` : "OPEN REQUESTS"}
            </div>
            <button
              onClick={autoFillToCapacity}
              disabled={!hasLane}
              title={hasLane ? "" : "Pick an origin and a destination first"}
              style={{
                ...btnSmall,
                width: "auto",
                padding: "0 12px",
                height: 30,
                fontSize: 11,
                opacity: hasLane ? 1 : 0.4,
                cursor: hasLane ? "pointer" : "not-allowed",
              }}
            >
              Auto-fill to capacity
            </button>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 280, overflowY: "auto" }}>
            {!hasLane && (
              <div style={{ padding: "18px 4px", color: "#6f8c7c", fontSize: 13 }}>
                Pick both an origin and a destination above — a load only ever pulls requests that match both exactly, so nothing gets mixed onto the wrong truck.
              </div>
            )}
            {hasLane && laneMatches.length === 0 && (
              <div style={{ padding: "18px 4px", color: "#6f8c7c", fontSize: 13 }}>
                No open requests match this exact lane right now.
              </div>
            )}
            {laneMatches.map((o) => {
              const checked = selectedIds.has(o.id);
              return (
                <label
                  key={o.id}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "10px 14px",
                    background: checked ? "rgba(240,201,90,0.08)" : "#0f2419",
                    border: `1px solid ${checked ? "#f0c95a" : "#1d3527"}`,
                    borderRadius: 10,
                    cursor: "pointer",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <input type="checkbox" checked={checked} onChange={() => toggleOrder(o.id)} />
                    <div>
                      <div style={{ fontWeight: 700, color: "#eaf3ec", fontSize: 14 }}>
                        {combinedName(o)}
                      </div>
                      <div style={{ fontSize: 12, color: "#8fae9c", marginTop: 2 }}>
                        {o.origin}{o.originTrack ? ` (${o.originTrack})` : ""} → {o.destination}{o.track ? ` (${o.track})` : ""} · Requested by {o.requestedBy || "—"} · {o.requestedDates || "date flexible"}
                      </div>
                    </div>
                  </div>
                  <div style={{ fontFamily: "ui-monospace, monospace", color: "#f0c95a", fontWeight: 800, whiteSpace: "nowrap", marginLeft: 12 }}>
                    {formatStalls(o.stallSpace)} stall{o.stallSpace === 1 ? "" : "s"}
                  </div>
                </label>
              );
            })}
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 24 }}>
            <div>
              {editing && (
                <button
                  onClick={() => {
                    window.confirm("Delete this load? Its orders go back to the open board.") && onDelete();
                  }}
                  style={{
                    background: "transparent",
                    border: "1px solid #ff6b6b",
                    color: "#ff6b6b",
                    borderRadius: 8,
                    padding: "10px 16px",
                    fontWeight: 700,
                    fontSize: 13,
                    cursor: "pointer",
                  }}
                >
                  Delete load
                </button>
              )}
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button
                onClick={onClose}
                style={{
                  background: "transparent",
                  border: "1px solid #2c4a3a",
                  color: "#9fb8a8",
                  borderRadius: 8,
                  padding: "10px 16px",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: "pointer",
                }}
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                disabled={selectedIds.size === 0}
                style={{
                  background: selectedIds.size === 0 ? "#2c4a3a" : "#f0c95a",
                  color: selectedIds.size === 0 ? "#6f8c7c" : "#0a2417",
                  border: "none",
                  borderRadius: 8,
                  padding: "10px 20px",
                  fontWeight: 800,
                  fontSize: 13,
                  cursor: selectedIds.size === 0 ? "not-allowed" : "pointer",
                }}
              >
                {editing ? "Save changes" : "Create load"} · {formatStalls(booked)} stall{booked === 1 ? "" : "s"}
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------- row for a manually built load ---------- */

function CreatedLoadRow({ load, resolvedOrders, capacity, onClick, onDelete }) {
  const booked = stallsBooked({ orders: resolvedOrders });
  const status = statusFor(booked, capacity);
  const s = STATUS_STYLE[status];

  return (
    <div
      onClick={onClick}
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        padding: "14px 18px",
        background: "#0b1c14",
        border: "1px solid #1d3527",
        borderRadius: 12,
        cursor: "pointer",
        gap: 12,
      }}
    >
      <div style={{ minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontWeight: 900, fontSize: 15 }}>{load.label}</span>
          <span style={{ fontSize: 13, color: "#9fb8a8", fontFamily: "ui-monospace, monospace" }}>
            {load.origin || "?"} → {load.destination || "?"}
          </span>
        </div>
        <div style={{ fontSize: 12, color: "#7f9c8c", marginTop: 4 }}>
          {resolvedOrders.length} {resolvedOrders.length === 1 ? "request" : "requests"}
          {load.driver ? ` · driver: ${load.driver}` : ""}
          {load.tripDate ? ` · ${load.tripDate}` : ""}
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexShrink: 0 }}>
        <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 13, color: "#e8c96a", fontWeight: 700, whiteSpace: "nowrap" }}>
          {formatStalls(booked)} / {formatStalls(capacity)}
        </div>
        <span
          style={{
            fontSize: 11,
            fontWeight: 800,
            letterSpacing: 1,
            padding: "5px 10px",
            borderRadius: 999,
            border: `1px solid ${s.color}`,
            color: s.color,
            whiteSpace: "nowrap",
          }}
        >
          {s.label}
        </span>
        <button
          onClick={(e) => {
            e.stopPropagation();
            window.confirm(`Delete "${load.label}"? Its orders go back to the open board.`) && onDelete();
          }}
          style={{ ...btnSmall, width: 28, height: 28, fontSize: 14, color: "#ff9d9d", borderColor: "#4a2c2c" }}
        >
          ×
        </button>
      </div>
    </div>
  );
}

/* ---------- live data hook ----------
   Polls ORDERS_JSON_URL, falls back to the snapshot if it's
   unreachable, and tracks which order IDs are new since the
   last successful fetch so the UI can announce real changes
   instead of a canned demo loop. */

function useLiveOrders() {
  const [orders, setOrders] = useState(SNAPSHOT_ORDERS);
  const [source, setSource] = useState("snapshot"); // "live" | "snapshot"
  const [syncedAt, setSyncedAt] = useState(null);
  const [newlyAdded, setNewlyAdded] = useState([]); // most-recent-first queue of new orders
  const seenIds = useRef(new Set(SNAPSHOT_ORDERS.map((o) => o.id)));
  const firstLiveFetch = useRef(true);

  useEffect(() => {
    let cancelled = false;

    async function poll() {
      try {
        const res = await fetch(`${ORDERS_JSON_URL}?t=${Date.now()}`, { cache: "no-store" });
        if (!res.ok) throw new Error(`status ${res.status}`);
        const payload = await res.json();
        const liveOrders = Array.isArray(payload) ? payload : payload.orders;
        if (!Array.isArray(liveOrders)) throw new Error("unexpected payload shape");
        if (cancelled) return;

        if (firstLiveFetch.current) {
          seenIds.current = new Set(liveOrders.map((o) => o.id));
          firstLiveFetch.current = false;
        } else {
          const fresh = liveOrders.filter((o) => !seenIds.current.has(o.id));
          if (fresh.length) {
            setNewlyAdded((prev) => [...fresh, ...prev].slice(0, 20));
            seenIds.current = new Set(liveOrders.map((o) => o.id));
          }
        }

        setOrders(liveOrders);
        setSource("live");
        setSyncedAt(payload.syncedAt || new Date().toISOString());
      } catch (err) {
        if (source !== "live") setSource("snapshot");
      }
    }

    poll();
    const id = setInterval(poll, POLL_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(id);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return { orders, source, syncedAt, newlyAdded };
}

function Ticker({ items, source }) {
  const text = items.length
    ? items
        .map((o) => `NEW REQUEST · ${o.origin} → ${o.destination}${o.track ? " (" + o.track + ")" : ""} · ${o.stallSpace} stall${o.stallSpace === 1 ? "" : "s"} · ${combinedName(o)}`)
        .join("      ✦      ")
    : source === "live"
    ? "MONITORING SHAREPOINT FOR NEW REQUESTS      ✦      MONITORING SHAREPOINT FOR NEW REQUESTS"
    : "SHOWING SNAPSHOT DATA — CONNECT data/orders.json FOR LIVE UPDATES      ✦      SHOWING SNAPSHOT DATA — CONNECT data/orders.json FOR LIVE UPDATES";

  return (
    <div
      style={{
        background: "#0f2419",
        borderTop: "1px solid #1d3527",
        borderBottom: "1px solid #1d3527",
        overflow: "hidden",
        whiteSpace: "nowrap",
        padding: "8px 0",
      }}
    >
      <div
        style={{
          display: "inline-block",
          paddingLeft: "100%",
          animation: "sallee-ticker 38s linear infinite",
          fontFamily: "ui-monospace, monospace",
          fontSize: 13,
          color: "#e8c96a",
          letterSpacing: 0.5,
        }}
      >
        {text}
      </div>
      <style>{`
        @keyframes sallee-ticker {
          0% { transform: translateX(0); }
          100% { transform: translateX(-100%); }
        }
      `}</style>
    </div>
  );
}

/* ---------- board row (for auto-grouped loads) ---------- */

function BoardRow({ load, capacity, booked, justAnnounced, onClick }) {
  const status = statusFor(booked, capacity);
  const s = STATUS_STYLE[status];
  const originNames = joinWithMore(uniqueList(load.orders, trainerFarmOf));
  const destNames = joinWithMore(uniqueList(load.orders, farmTrainerOf));
  const dateSummary = requestedDateSummary(load);

  const originSubtext = [load.originTrack ? `via ${load.originTrack}` : null, originNames]
    .filter(Boolean)
    .join(" · ") || "—";
  const destSubtext = [load.track ? `via ${load.track}` : null, destNames]
    .filter(Boolean)
    .join(" · ") || "—";

  return (
    <div
      onClick={onClick}
      style={{
        display: "grid",
        gridTemplateColumns: "1.3fr 1.3fr 0.9fr 1fr 1fr",
        alignItems: "center",
        padding: "14px 18px",
        borderBottom: "1px solid #16281f",
        cursor: "pointer",
        background: justAnnounced ? "rgba(240,201,90,0.08)" : "transparent",
        transition: "background 0.6s ease",
        gap: 10,
      }}
    >
      <div>
        <FlapText text={load.origin} size={16} />
        <div
          style={{ fontSize: 11, color: "#7f9c8c", marginTop: 3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}
          title={originSubtext}
        >
          {originSubtext}
        </div>
      </div>
      <div>
        <FlapText text={load.destination} size={16} color="#eaf3ec" />
        <div
          style={{ fontSize: 11, color: "#7f9c8c", marginTop: 3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}
          title={destSubtext}
        >
          {destSubtext}
        </div>
      </div>
      <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 13, color: "#e8c96a", fontWeight: 700 }}>
        {formatStalls(booked)} / {formatStalls(capacity)} stalls
      </div>
      <div
        style={{
          fontFamily: "ui-monospace, monospace",
          fontSize: 12,
          fontWeight: dateSummary.mismatched ? 800 : 500,
          color: dateSummary.mismatched ? "#f0a95a" : "#c9dcd0",
        }}
        title={dateSummary.mismatched ? "Orders in this load have different requested dates — open the load to check" : undefined}
      >
        {dateSummary.mismatched && "⚠ "}
        {dateSummary.text}
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <span
          style={{
            fontSize: 11,
            fontWeight: 800,
            letterSpacing: 1,
            padding: "5px 10px",
            borderRadius: 999,
            border: `1px solid ${s.color}`,
            color: s.color,
          }}
        >
          {s.label}
        </span>
      </div>
    </div>
  );
}

/* ---------- main app ---------- */

export default function SalleeLoadBoard() {
  const { orders, source, syncedAt, newlyAdded } = useLiveOrders();

  // Manually built loads persist in this browser only (see
  // CREATED_LOADS_STORAGE_KEY above) — they never touch the sheet.
  const [createdLoads, setCreatedLoads] = useState(() => loadCreatedLoads());
  useEffect(() => {
    saveCreatedLoads(createdLoads);
  }, [createdLoads]);

  const ordersById = useMemo(() => new Map(orders.map((o) => [o.id, o])), [orders]);
  const usedOrderIds = useMemo(
    () => new Set(createdLoads.flatMap((cl) => cl.orderIds)),
    [createdLoads]
  );
  // The auto-grouped board and the "build a load" picker only ever see
  // requests that haven't already been claimed by a manually built load.
  const openPool = useMemo(() => orders.filter((o) => !usedOrderIds.has(o.id)), [orders, usedOrderIds]);

  const [editingLoad, setEditingLoad] = useState(null); // null closed, {} = create, {load} = edit

  function handleSaveCreatedLoad(saved) {
    setCreatedLoads((prev) =>
      prev.some((cl) => cl.id === saved.id)
        ? prev.map((cl) => (cl.id === saved.id ? saved : cl))
        : [...prev, saved]
    );
    setEditingLoad(null);
  }

  function handleDeleteCreatedLoad(id) {
    setCreatedLoads((prev) => prev.filter((cl) => cl.id !== id));
    setEditingLoad(null);
  }

  const allLoads = useMemo(() => buildLoads(openPool), [openPool]);

  const [capacities, setCapacities] = useState({});
  const getCapacity = (key) => capacities[key] ?? DEFAULT_CAPACITY;
  const setCapacity = (key, val) =>
    setCapacities((prev) => ({ ...prev, [key]: val }));

  const [fromFilter, setFromFilter] = useState("ALL");
  const [toFilter, setToFilter] = useState("ALL");
  const [query, setQuery] = useState("");

  const origins = useMemo(
    () => Array.from(new Set(orders.map((o) => o.origin))).sort(),
    [orders]
  );
  const destinations = useMemo(
    () => Array.from(new Set(orders.map((o) => o.destination))).sort(),
    [orders]
  );

  const isSearching = fromFilter !== "ALL" || toFilter !== "ALL" || query.trim() !== "";

  const filteredLoads = useMemo(() => {
    return allLoads.filter((l) => {
      if (fromFilter !== "ALL" && l.origin !== fromFilter) return false;
      if (toFilter !== "ALL" && l.destination !== toFilter) return false;
      if (query.trim()) {
        const q = query.trim().toUpperCase();
        const hay = l.orders
          .map((o) => `${o.trainerFarm || ""} ${o.farmTrainer || ""} ${o.requestedBy || ""}`)
          .join(" ")
          .toUpperCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [allLoads, fromFilter, toFilter, query]);

  /* pagination / looping for the full board — grouped by origin so a
     page never splits one origin's loads across two screens */
  const PAGE_SIZE = 8;
  const pages = useMemo(() => {
    const byOrigin = new Map();
    for (const l of filteredLoads) {
      if (!byOrigin.has(l.origin)) byOrigin.set(l.origin, []);
      byOrigin.get(l.origin).push(l);
    }
    const out = [];
    for (const [origin, rows] of byOrigin) {
      for (let i = 0; i < rows.length; i += PAGE_SIZE) {
        out.push({ origin, rows: rows.slice(i, i + PAGE_SIZE) });
      }
    }
    return out;
  }, [filteredLoads]);
  const pageCount = Math.max(1, pages.length);

  const [page, setPage] = useState(0);
  const [paused, setPaused] = useState(false);

  useEffect(() => {
    setPage(0);
  }, [fromFilter, toFilter, query]);
  useEffect(() => {
    if (page >= pageCount) setPage(0);
  }, [pageCount, page]);
  useEffect(() => {
    if (isSearching || paused) return;
    const id = setInterval(() => {
      setPage((p) => (p + 1) % pageCount);
    }, 7000);
    return () => clearInterval(id);
  }, [pageCount, isSearching, paused]);

  const currentPage = pages[page] || { origin: null, rows: [] };
  const visibleLoads = isSearching ? filteredLoads : currentPage.rows;

  // Recently-new rows (real ones, from the live diff) stay highlighted on
  // the board for a couple of minutes after they show up.
  const [announcedKeys, setAnnouncedKeys] = useState(new Set());
  useEffect(() => {
    if (!newlyAdded.length) return;
    const keys = newlyAdded.map(loadKey);
    setAnnouncedKeys((prev) => new Set([...prev, ...keys]));
    const timers = keys.map((k) =>
      setTimeout(() => {
        setAnnouncedKeys((prev) => {
          const next = new Set(prev);
          next.delete(k);
          return next;
        });
      }, 120000)
    );
    return () => timers.forEach(clearTimeout);
  }, [newlyAdded]);

  const [selectedKey, setSelectedKey] = useState(null);
  const selectedLoad = allLoads.find((l) => l.key === selectedKey) || null;

  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);

  const totalActive = openPool.length;
  const totalStalls = openPool.reduce((s, o) => s + (Number(o.stallSpace) || 0), 0);

  return (
    <div
      style={{
        minHeight: "100%",
        background: "radial-gradient(ellipse at top, #0f2a1c 0%, #081611 60%, #05100b 100%)",
        color: "#eaf3ec",
        fontFamily: "Arial, Helvetica, sans-serif",
      }}
    >
      {/* header */}
      <div style={{ padding: "26px 26px 0 26px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ fontSize: 12, letterSpacing: 6, color: "#e8c96a", fontWeight: 800 }}>
                SALLEE HORSE VANS
              </div>
              <span
                title={source === "live" ? `Last synced ${syncedAt ? new Date(syncedAt).toLocaleTimeString() : ""}` : "Live feed not connected yet — showing snapshot data"}
                style={{
                  fontSize: 10,
                  fontWeight: 800,
                  letterSpacing: 1,
                  padding: "2px 8px",
                  borderRadius: 999,
                  border: `1px solid ${source === "live" ? "#7be0a8" : "#6f8c7c"}`,
                  color: source === "live" ? "#7be0a8" : "#8fae9c",
                }}
              >
                {source === "live" ? "● LIVE" : "○ SNAPSHOT"}
              </span>
            </div>
            <div style={{ fontSize: 34, fontWeight: 900, letterSpacing: 1, marginTop: 2 }}>
              LOAD BOARD
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "flex-start", gap: 16 }}>
            <div style={{ textAlign: "right", fontFamily: "ui-monospace, monospace" }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: "#f0c95a" }}>
                {now.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
              </div>
              <div style={{ fontSize: 12, color: "#8fae9c" }}>
                {now.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" })}
              </div>
            </div>
            <button
              onClick={() => setEditingLoad({})}
              style={{
                background: "#f0c95a",
                color: "#0a2417",
                border: "none",
                borderRadius: 10,
                padding: "12px 18px",
                fontWeight: 800,
                fontSize: 13,
                letterSpacing: 0.5,
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              + Build a Load
            </button>
          </div>
        </div>

        <div style={{ display: "flex", gap: 22, marginTop: 18, fontFamily: "ui-monospace, monospace", fontSize: 13, color: "#8fae9c" }}>
          <div><span style={{ color: "#f0c95a", fontWeight: 800 }}>{totalActive}</span> open requests</div>
          <div><span style={{ color: "#f0c95a", fontWeight: 800 }}>{filteredLoads.length}</span> loads building</div>
          <div><span style={{ color: "#f0c95a", fontWeight: 800 }}>{totalStalls}</span> stalls requested</div>
          {createdLoads.length > 0 && (
            <div><span style={{ color: "#f0c95a", fontWeight: 800 }}>{createdLoads.length}</span> {createdLoads.length === 1 ? "load" : "loads"} created</div>
          )}
          {source === "live" && syncedAt && (
            <div>synced {new Date(syncedAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}</div>
          )}
        </div>
      </div>

      {/* ticker */}
      <div style={{ marginTop: 18 }}>
        <Ticker items={newlyAdded} source={source} />
      </div>

      {/* search / booking kiosk */}
      <div style={{ padding: "20px 26px 0 26px" }}>
        <div
          style={{
            background: "#0b1c14",
            border: "1px solid #1d3527",
            borderRadius: 14,
            padding: "16px 18px",
            display: "flex",
            gap: 14,
            flexWrap: "wrap",
            alignItems: "center",
          }}
        >
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={labelStyle}>FROM</label>
            <select value={fromFilter} onChange={(e) => setFromFilter(e.target.value)} style={selectStyle}>
              <option value="ALL">Any origin</option>
              {origins.map((o) => (
                <option key={o} value={o}>{o}</option>
              ))}
            </select>
          </div>

          <div style={{ fontSize: 20, color: "#3f5f4d", marginTop: 18 }}>→</div>

          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={labelStyle}>TO</label>
            <select value={toFilter} onChange={(e) => setToFilter(e.target.value)} style={selectStyle}>
              <option value="ALL">Any destination</option>
              {destinations.map((d) => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1, minWidth: 180 }}>
            <label style={labelStyle}>SEARCH TRAINER / FARM</label>
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="e.g. BROWN, KENNEALLY..."
              style={{ ...selectStyle, width: "100%" }}
            />
          </div>

          {(fromFilter !== "ALL" || toFilter !== "ALL" || query) && (
            <button
              onClick={() => { setFromFilter("ALL"); setToFilter("ALL"); setQuery(""); }}
              style={{ ...btnSmall, width: "auto", padding: "0 14px", height: 38, alignSelf: "flex-end" }}
            >
              Clear
            </button>
          )}
        </div>
      </div>

      {/* board */}
      <div style={{ padding: "20px 26px 40px 26px" }}>
        {!isSearching && currentPage.origin && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, fontFamily: "ui-monospace, monospace", fontSize: 12, color: "#7f9c8c", letterSpacing: 1 }}>
            <span>SHOWING ORIGIN</span>
            <span style={{ color: "#f0c95a", fontWeight: 800, fontSize: 14, letterSpacing: 1.5 }}>{currentPage.origin}</span>
          </div>
        )}
        <div
          onMouseEnter={() => setPaused(true)}
          onMouseLeave={() => setPaused(false)}
          style={{
            background: "#0b1c14",
            border: "1px solid #1d3527",
            borderRadius: 14,
            overflow: "hidden",
          }}
        >
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1.3fr 1.3fr 0.9fr 1fr 1fr",
              padding: "10px 18px",
              background: "#0f2419",
              fontFamily: "ui-monospace, monospace",
              fontSize: 11,
              letterSpacing: 2,
              color: "#7f9c8c",
              fontWeight: 700,
              gap: 10,
            }}
          >
            <div>ORIGIN / FROM</div>
            <div>DESTINATION / TO</div>
            <div>STALLS</div>
            <div>REQ. DATE</div>
            <div style={{ textAlign: "right" }}>STATUS</div>
          </div>

          {visibleLoads.length === 0 && (
            <div style={{ padding: "40px 18px", textAlign: "center", color: "#6f8c7c" }}>
              No loads match that search. Try a different origin, destination, or name.
            </div>
          )}

          {visibleLoads.map((l) => (
            <BoardRow
              key={l.key}
              load={l}
              capacity={getCapacity(l.key)}
              booked={stallsBooked(l)}
              justAnnounced={announcedKeys.has(l.key)}
              onClick={() => setSelectedKey(l.key)}
            />
          ))}
        </div>

        {!isSearching && pageCount > 1 && (
          <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 10, marginTop: 16 }}>
            <button onClick={() => setPage((p) => (p - 1 + pageCount) % pageCount)} style={btnSmall}>‹</button>
            {Array.from({ length: pageCount }).map((_, i) => (
              <div
                key={i}
                onClick={() => setPage(i)}
                style={{
                  width: 8, height: 8, borderRadius: 999,
                  background: i === page ? "#f0c95a" : "#2c4a3a",
                  cursor: "pointer",
                }}
              />
            ))}
            <button onClick={() => setPage((p) => (p + 1) % pageCount)} style={btnSmall}>›</button>
            <span style={{ fontFamily: "ui-monospace, monospace", fontSize: 11, color: "#6f8c7c", marginLeft: 8 }}>
              {paused ? "PAUSED" : "AUTO-ADVANCING · hover to pause"}
            </span>
          </div>
        )}
      </div>

      {/* manually built loads */}
      {createdLoads.length > 0 && (
        <div style={{ padding: "0 26px 40px 26px" }}>
          <div style={{ fontSize: 11, letterSpacing: 2, color: "#7f9c8c", fontFamily: "ui-monospace, monospace", marginBottom: 10 }}>
            LOADS BUILT — {createdLoads.length}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {createdLoads.map((cl) => {
              const resolvedOrders = cl.orderIds.map((id) => ordersById.get(id)).filter(Boolean);
              return (
                <CreatedLoadRow
                  key={cl.id}
                  load={cl}
                  resolvedOrders={resolvedOrders}
                  capacity={cl.capacity || DEFAULT_CAPACITY}
                  onClick={() => setEditingLoad({ load: cl })}
                  onDelete={() => handleDeleteCreatedLoad(cl.id)}
                />
              );
            })}
          </div>
        </div>
      )}

      {selectedLoad && (
        <LoadModal
          load={selectedLoad}
          capacity={getCapacity(selectedLoad.key)}
          onCapacityChange={(v) => setCapacity(selectedLoad.key, v)}
          onClose={() => setSelectedKey(null)}
        />
      )}

      {editingLoad && (
        <BuildLoadModal
          pool={
            editingLoad.load
              ? orders.filter((o) => {
                  const owner = createdLoads.find((cl) => cl.orderIds.includes(o.id));
                  return !owner || owner.id === editingLoad.load.id;
                })
              : openPool
          }
          origins={origins}
          destinations={destinations}
          initial={editingLoad.load || null}
          nextLoadNumber={createdLoads.length + 1}
          onSave={handleSaveCreatedLoad}
          onDelete={() => editingLoad.load && handleDeleteCreatedLoad(editingLoad.load.id)}
          onClose={() => setEditingLoad(null)}
        />
      )}
    </div>
  );
}
