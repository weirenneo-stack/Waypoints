import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import { ComposableMap, Geographies, Geography, Marker } from "@vnedyalk0v/react19-simple-maps";
import {
  Plane, MapPin, CalendarDays, Trash2, AlertTriangle, Stamp, ListOrdered,
  BarChart3, PlusCircle, ChevronDown, X, Loader2, Home, Briefcase, Pencil,
  Download, Upload, Map as MapIcon, RefreshCw, CloudOff, Search, StickyNote,
  Settings2,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/* GOOGLE SHEET SETUP — paste your deployed Web App URL here               */
/* ---------------------------------------------------------------------- */
//
// This app uses a Google Sheet as its database, via a small Apps Script
// "Web App" that sits in front of it and speaks JSON over HTTP. It's free,
// needs no server of your own, and there's no login screen — the moment
// anyone opens this published artifact, it loads straight from the sheet.
//
// Because there's no login, this is ONE shared dataset for anyone who opens
// this published artifact — perfect for using it yourself across your own
// devices, but treat the Web App URL like a shared secret: anyone who has
// it (or opens this artifact) can read and write the same trips, the same
// way anyone with an "anyone with the link can edit" spreadsheet link can.
//
// SETUP (about 5 minutes):
// 1. Create a new Google Sheet (sheets.new).
// 2. Extensions → Apps Script. Delete any starter code and paste in the
//    script below. Save it.
// 3. Click Deploy → New deployment → gear icon → "Web app".
//      - Execute as: Me
//      - Who has access: Anyone
//    Click Deploy, authorize it when prompted, then copy the Web app URL.
// 4. Paste that URL into SHEET_WEB_APP_URL below.
//
// ---- Apps Script code to paste in step 2 ----
//
//   const SHEET_NAME = "Trips";
//
//   function getSheet_() {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     let sheet = ss.getSheetByName(SHEET_NAME);
//     if (!sheet) {
//       sheet = ss.insertSheet(SHEET_NAME);
//       sheet.appendRow(["id", "country", "start_date", "end_date", "trip_type", "created_at", "notes", "stops"]);
//     }
//     return sheet;
//   }
//
//   function doGet(e) {
//     const sheet = getSheet_();
//     const rows = sheet.getDataRange().getValues();
//     const headers = rows.shift();
//     const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
//     const data = rows
//       .filter((r) => r[0])
//       .map((r) => {
//         const obj = {};
//         headers.forEach((h, i) => {
//           let v = r[i];
//           if (h === "start_date" || h === "end_date") {
//             if (Object.prototype.toString.call(v) === "[object Date]") v = Utilities.formatDate(v, tz, "yyyy-MM-dd");
//             else if (typeof v === "string" && v.indexOf("T") > -1) v = v.slice(0, 10);
//           }
//           obj[h] = v;
//         });
//         return obj;
//       });
//     return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
//   }
//
//   function doPost(e) {
//     const body = JSON.parse(e.postData.contents);
//     const sheet = getSheet_();
//     if (body.action === "add") {
//       sheet.appendRow([body.id, body.country, body.start_date, body.end_date, body.trip_type, new Date().toISOString(), body.notes || "", body.stops || ""]);
//     } else if (body.action === "bulkAdd") {
//       body.rows.forEach((r) =>
//         sheet.appendRow([r.id, r.country, r.start_date, r.end_date, r.trip_type, new Date().toISOString(), r.notes || "", r.stops || ""])
//       );
//     } else if (body.action === "update") {
//       const rows = sheet.getDataRange().getValues();
//       for (let i = 1; i < rows.length; i++) {
//         if (String(rows[i][0]) === String(body.id)) {
//           sheet.getRange(i + 1, 1, 1, 5).setValues([[body.id, body.country, body.start_date, body.end_date, body.trip_type]]);
//           sheet.getRange(i + 1, 7, 1, 2).setValues([[body.notes || "", body.stops || ""]]);
//           break;
//         }
//       }
//     } else if (body.action === "delete") {
//       const rows = sheet.getDataRange().getValues();
//       for (let i = 1; i < rows.length; i++) {
//         if (String(rows[i][0]) === String(body.id)) { sheet.deleteRow(i + 1); break; }
//       }
//     }
//     return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
//   }
//
// ---- end Apps Script code ----
//
// UPGRADING AN EXISTING SHEET: if you already had this set up before notes
// and stops existed, just add "notes" in cell G1 and "stops" in H1 of your
// existing header row (row 1) — no need to recreate the sheet. Then replace
// doGet/doPost with the versions above and redeploy as a New version
// (Deploy → Manage deployments → pencil icon → Version: New version).

const SHEET_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwtSB-MtyxR1bbMmln4yDCfMJE_9cILhPus_lUQucF0KuKkEcvGrpWWGQJVRthQadqSFg/exec";
const sheetConfigured = SHEET_WEB_APP_URL !== "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL";
const POLL_INTERVAL_MS = 30000; // background refresh, since Apps Script has no realtime push

async function fetchTripsFromSheet() {
  const res = await fetch(SHEET_WEB_APP_URL, { method: "GET" });
  if (!res.ok) throw new Error("Sheet fetch failed");
  const data = await res.json();
  return data.map(rowToTrip);
}

async function postToSheet(payload) {
  // text/plain avoids a CORS preflight that Apps Script web apps don't handle well.
  const res = await fetch(SHEET_WEB_APP_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error("Sheet write failed");
  return res.json();
}

function rowToTrip(row) {
  return {
    id: String(row.id),
    country: row.country,
    start: row.start_date,
    end: row.end_date,
    type: row.trip_type || "personal",
    notes: row.notes || "",
    stops: row.stops ? String(row.stops).split(",").map((s) => s.trim()).filter(Boolean) : [],
  };
}

/* ---------------------------------------------------------------------- */
/* Reference data                                                          */
/* ---------------------------------------------------------------------- */

const COUNTRY_CONTINENT = {
  "United States": "Americas", Canada: "Americas", Mexico: "Americas",
  Brazil: "Americas", Argentina: "Americas", Chile: "Americas", Peru: "Americas",
  Colombia: "Americas", Cuba: "Americas", "Costa Rica": "Americas", Panama: "Americas",
  Ecuador: "Americas", Uruguay: "Americas", Bolivia: "Americas", Guatemala: "Americas",
  Jamaica: "Americas", "Dominican Republic": "Americas",
  "United Kingdom": "Europe", France: "Europe", Germany: "Europe", Italy: "Europe",
  Spain: "Europe", Portugal: "Europe", Netherlands: "Europe", Switzerland: "Europe",
  Austria: "Europe", Belgium: "Europe", Greece: "Europe", Ireland: "Europe",
  Iceland: "Europe", Norway: "Europe", Sweden: "Europe", Denmark: "Europe",
  Finland: "Europe", Poland: "Europe", "Czech Republic": "Europe", Hungary: "Europe",
  Croatia: "Europe", Romania: "Europe", Ukraine: "Europe", Malta: "Europe",
  Slovenia: "Europe", Slovakia: "Europe", Estonia: "Europe", Latvia: "Europe",
  Lithuania: "Europe", Turkey: "Europe",
  China: "Asia", Japan: "Asia", "South Korea": "Asia", India: "Asia",
  Thailand: "Asia", Vietnam: "Asia", Indonesia: "Asia", Malaysia: "Asia",
  Singapore: "Asia", Philippines: "Asia", Cambodia: "Asia", Laos: "Asia",
  "Sri Lanka": "Asia", Nepal: "Asia", "United Arab Emirates": "Asia",
  "Saudi Arabia": "Asia", Qatar: "Asia", Israel: "Asia", Jordan: "Asia",
  Taiwan: "Asia", "Hong Kong": "Asia", Mongolia: "Asia", Kazakhstan: "Asia",
  Egypt: "Africa", "South Africa": "Africa", Morocco: "Africa", Kenya: "Africa",
  Tanzania: "Africa", Nigeria: "Africa", Ghana: "Africa", Ethiopia: "Africa",
  Tunisia: "Africa", Namibia: "Africa", Botswana: "Africa", Rwanda: "Africa",
  Senegal: "Africa", Uganda: "Africa", Zambia: "Africa", Zimbabwe: "Africa",
  Australia: "Oceania", "New Zealand": "Oceania", Fiji: "Oceania",
  "Papua New Guinea": "Oceania",
};
const COUNTRY_LIST = Object.keys(COUNTRY_CONTINENT).sort();
const CONTINENT_ORDER = ["Americas", "Europe", "Asia", "Africa", "Oceania", "Other"];
const CONTINENT_ACCENT = {
  Americas: "#4FB3A9", Europe: "#E3A23C", Asia: "#D97E7E",
  Africa: "#C9A6E8", Oceania: "#7FB0E8", Other: "#8B93A7",
};
const RESIDENCY_THRESHOLD = 183;
const WORLD_GEO_URL = "https://unpkg.com/world-atlas@2.0.2/countries-110m.json";

const COUNTRY_LATLON = {
  "United States": [39.8, -98.6], Canada: [56.1, -106.3], Mexico: [23.6, -102.5],
  Brazil: [-14.2, -51.9], Argentina: [-38.4, -63.6], Chile: [-35.7, -71.5], Peru: [-9.2, -75.0],
  Colombia: [4.6, -74.1], Cuba: [21.5, -77.8], "Costa Rica": [9.7, -83.8], Panama: [8.5, -80.8],
  Ecuador: [-1.8, -78.2], Uruguay: [-32.5, -55.8], Bolivia: [-16.3, -63.6], Guatemala: [15.8, -90.2],
  Jamaica: [18.1, -77.3], "Dominican Republic": [18.7, -70.2],
  "United Kingdom": [54.0, -2.0], France: [46.6, 2.2], Germany: [51.2, 10.5], Italy: [41.9, 12.6],
  Spain: [40.0, -4.0], Portugal: [39.4, -8.2], Netherlands: [52.1, 5.3], Switzerland: [46.8, 8.2],
  Austria: [47.5, 14.6], Belgium: [50.5, 4.5], Greece: [39.1, 21.8], Ireland: [53.4, -8.2],
  Iceland: [64.9, -19.0], Norway: [60.5, 8.5], Sweden: [60.1, 18.6], Denmark: [56.3, 9.5],
  Finland: [61.9, 25.7], Poland: [51.9, 19.1], "Czech Republic": [49.8, 15.5], Hungary: [47.2, 19.5],
  Croatia: [45.1, 15.2], Romania: [45.9, 24.9], Ukraine: [48.4, 31.2], Malta: [35.9, 14.4],
  Slovenia: [46.1, 14.8], Slovakia: [48.7, 19.7], Estonia: [58.6, 25.0], Latvia: [56.9, 24.6],
  Lithuania: [55.2, 23.9], Turkey: [38.9, 35.2],
  China: [35.9, 104.2], Japan: [36.2, 138.3], "South Korea": [35.9, 127.8], India: [20.6, 79.0],
  Thailand: [15.9, 100.9], Vietnam: [14.1, 108.3], Indonesia: [-0.8, 113.9], Malaysia: [4.2, 101.9],
  Singapore: [1.35, 103.8], Philippines: [12.9, 121.8], Cambodia: [12.6, 104.9], Laos: [19.9, 102.5],
  "Sri Lanka": [7.9, 80.7], Nepal: [28.4, 84.1], "United Arab Emirates": [23.4, 53.8],
  "Saudi Arabia": [23.9, 45.1], Qatar: [25.4, 51.2], Israel: [31.0, 34.8], Jordan: [30.6, 36.2],
  Taiwan: [23.7, 121.0], "Hong Kong": [22.3, 114.2], Mongolia: [46.9, 103.8], Kazakhstan: [48.0, 66.9],
  Egypt: [26.8, 30.8], "South Africa": [-30.6, 22.9], Morocco: [31.8, -7.1], Kenya: [-0.02, 37.9],
  Tanzania: [-6.4, 34.9], Nigeria: [9.1, 8.7], Ghana: [7.9, -1.0], Ethiopia: [9.1, 40.5],
  Tunisia: [33.9, 9.5], Namibia: [-22.9, 18.5], Botswana: [-22.3, 24.7], Rwanda: [-1.9, 29.9],
  Senegal: [14.5, -14.5], Uganda: [1.4, 32.3], Zambia: [-13.1, 27.8], Zimbabwe: [-19.0, 29.2],
  Australia: [-25.3, 133.8], "New Zealand": [-41.0, 174.8], Fiji: [-17.7, 178.1],
  "Papua New Guinea": [-6.3, 143.9],
};

const COUNTRY_ISO2 = {
  "United States": "US", Canada: "CA", Mexico: "MX", Brazil: "BR", Argentina: "AR",
  Chile: "CL", Peru: "PE", Colombia: "CO", Cuba: "CU", "Costa Rica": "CR", Panama: "PA",
  Ecuador: "EC", Uruguay: "UY", Bolivia: "BO", Guatemala: "GT", Jamaica: "JM",
  "Dominican Republic": "DO",
  "United Kingdom": "GB", France: "FR", Germany: "DE", Italy: "IT", Spain: "ES",
  Portugal: "PT", Netherlands: "NL", Switzerland: "CH", Austria: "AT", Belgium: "BE",
  Greece: "GR", Ireland: "IE", Iceland: "IS", Norway: "NO", Sweden: "SE", Denmark: "DK",
  Finland: "FI", Poland: "PL", "Czech Republic": "CZ", Hungary: "HU", Croatia: "HR",
  Romania: "RO", Ukraine: "UA", Malta: "MT", Slovenia: "SI", Slovakia: "SK",
  Estonia: "EE", Latvia: "LV", Lithuania: "LT", Turkey: "TR",
  China: "CN", Japan: "JP", "South Korea": "KR", India: "IN", Thailand: "TH",
  Vietnam: "VN", Indonesia: "ID", Malaysia: "MY", Singapore: "SG", Philippines: "PH",
  Cambodia: "KH", Laos: "LA", "Sri Lanka": "LK", Nepal: "NP", "United Arab Emirates": "AE",
  "Saudi Arabia": "SA", Qatar: "QA", Israel: "IL", Jordan: "JO", Taiwan: "TW",
  "Hong Kong": "HK", Mongolia: "MN", Kazakhstan: "KZ",
  Egypt: "EG", "South Africa": "ZA", Morocco: "MA", Kenya: "KE", Tanzania: "TZ",
  Nigeria: "NG", Ghana: "GH", Ethiopia: "ET", Tunisia: "TN", Namibia: "NA",
  Botswana: "BW", Rwanda: "RW", Senegal: "SN", Uganda: "UG", Zambia: "ZM", Zimbabwe: "ZW",
  Australia: "AU", "New Zealand": "NZ", Fiji: "FJ", "Papua New Guinea": "PG",
};

function countryFlag(country) {
  const iso2 = COUNTRY_ISO2[country];
  if (!iso2) return "🏳️";
  return iso2.replace(/./g, (c) => String.fromCodePoint(127397 + c.charCodeAt(0)));
}

// Local device settings — home country + per-country thresholds. These are
// personal preferences rather than trip data, so they live in this browser
// only (not synced to the Sheet).
const SETTINGS_KEY = "waypoints:settings";

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* ignore */ }
  return { homeCountry: "", thresholds: {} };
}

function saveSettings(settings) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch (e) { /* ignore */ }
}

/* ---------------------------------------------------------------------- */
/* Design tokens                                                            */
/* ---------------------------------------------------------------------- */

const BG = "#12151C";
const SURFACE = "#181D27";
const SURFACE_RAISED = "#1F2530";
const BORDER = "#2A3140";
const TEXT = "#EDEFF3";
const MUTED = "#8B93A7";
const ACCENT = "#E3A23C";
const ACCENT_DIM = "#3A2F1E";
const TEAL = "#4FB3A9";
const DANGER = "#E2596B";
const DANGER_DIM = "#3A2126";
const BUSINESS = "#8AADE8";
const SERIF = "'Fraunces', ui-serif, Georgia, serif";
const SANS = "'Inter', ui-sans-serif, system-ui, sans-serif";
const inputStyle = { background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: TEXT };

const TRIP_TYPES = {
  personal: { label: "Personal", icon: Home, color: TEAL },
  business: { label: "Business", icon: Briefcase, color: BUSINESS },
};

/* ---------------------------------------------------------------------- */
/* Helpers                                                                  */
/* ---------------------------------------------------------------------- */

function toDate(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}
function inclusiveDays(startStr, endStr) {
  return Math.round((toDate(endStr) - toDate(startStr)) / 86400000) + 1;
}
function formatRange(startStr, endStr) {
  const opts = { month: "short", day: "numeric", year: "numeric" };
  return `${toDate(startStr).toLocaleDateString("en-US", opts)} \u2192 ${toDate(endStr).toLocaleDateString("en-US", opts)}`;
}
function continentOf(country) {
  return COUNTRY_CONTINENT[country] || "Other";
}
function overlapDaysInYear(trip, year) {
  const yStart = new Date(year, 0, 1);
  const yEnd = new Date(year, 11, 31);
  const lo = toDate(trip.start) > yStart ? toDate(trip.start) : yStart;
  const hi = toDate(trip.end) < yEnd ? toDate(trip.end) : yEnd;
  if (hi < lo) return 0;
  return Math.round((hi - lo) / 86400000) + 1;
}
function overlapDaysInRollingWindow(trip, windowStart, windowEnd) {
  const lo = toDate(trip.start) > windowStart ? toDate(trip.start) : windowStart;
  const hi = toDate(trip.end) < windowEnd ? toDate(trip.end) : windowEnd;
  if (hi < lo) return 0;
  return Math.round((hi - lo) / 86400000) + 1;
}
function validateTrip({ country, start, end }) {
  if (!country || !country.trim()) return "Enter a country name.";
  if (!start || !end) return "Pick a start and end date.";
  if (toDate(end) < toDate(start)) return "End date can't be before the start date.";
  return null;
}
function uid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}
function timeAgo(date) {
  if (!date) return "never";
  const secs = Math.round((Date.now() - date.getTime()) / 1000);
  if (secs < 5) return "just now";
  if (secs < 60) return `${secs}s ago`;
  const mins = Math.round(secs / 60);
  if (mins < 60) return `${mins}m ago`;
  return `${Math.round(mins / 60)}h ago`;
}

/* ---------------------------------------------------------------------- */
/* Root                                                                     */
/* ---------------------------------------------------------------------- */

export default function App() {
  const [trips, setTrips] = useState([]);
  const [loading, setLoading] = useState(sheetConfigured);
  const [syncing, setSyncing] = useState(false);
  const [lastSynced, setLastSynced] = useState(null);
  const [tab, setTab] = useState("log");
  const [toast, setToast] = useState(null);
  const [editingTrip, setEditingTrip] = useState(null);

  const refresh = useCallback(async (silent) => {
    if (!sheetConfigured) return;
    if (!silent) setLoading(true);
    setSyncing(true);
    try {
      const fresh = await fetchTripsFromSheet();
      setTrips(fresh);
      setLastSynced(new Date());
    } catch (e) {
      setToast("Couldn't reach your Google Sheet.");
    } finally {
      setLoading(false);
      setSyncing(false);
    }
  }, []);

  // Loads automatically the moment the artifact opens — no login, no gate.
  useEffect(() => {
    refresh(false);
    if (!sheetConfigured) return;
    const t = setInterval(() => refresh(true), POLL_INTERVAL_MS);
    return () => clearInterval(t);
  }, [refresh]);

  const showToast = useCallback((msg) => setToast(msg), []);
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2800);
    return () => clearTimeout(t);
  }, [toast]);

  // --- CRUD (optimistic, backed by the sheet) ------------------------------
  const addTrip = useCallback(
    async (trip) => {
      if (!sheetConfigured) { showToast("Connect your Google Sheet first — see the setup notes at the top of this file."); return; }
      const newTrip = { ...trip, id: uid(), notes: trip.notes || "", stops: trip.stops || [] };
      setTrips((prev) => [newTrip, ...prev]);
      try {
        await postToSheet({
          action: "add", id: newTrip.id, country: trip.country, start_date: trip.start, end_date: trip.end,
          trip_type: trip.type, notes: trip.notes || "", stops: (trip.stops || []).join(", "),
        });
        showToast(`${trip.country} logged \u2014 ${inclusiveDays(trip.start, trip.end)} days`);
      } catch (e) {
        setTrips((prev) => prev.filter((t) => t.id !== newTrip.id));
        showToast("Couldn't save that trip to your sheet — try again.");
      }
    },
    [showToast]
  );

  const updateTrip = useCallback(
    async (id, fields) => {
      if (!sheetConfigured) return;
      const prevTrips = trips;
      setTrips((p) => p.map((t) => (t.id === id ? { ...t, ...fields } : t)));
      try {
        await postToSheet({
          action: "update", id, country: fields.country, start_date: fields.start, end_date: fields.end,
          trip_type: fields.type, notes: fields.notes || "", stops: (fields.stops || []).join(", "),
        });
        showToast(`${fields.country} updated`);
        setEditingTrip(null);
      } catch (e) {
        setTrips(prevTrips);
        showToast("Couldn't save your changes — try again.");
      }
    },
    [trips, showToast]
  );

  const deleteTrip = useCallback(
    async (id) => {
      if (!sheetConfigured) return;
      const prevTrips = trips;
      setTrips((p) => p.filter((t) => t.id !== id));
      try {
        await postToSheet({ action: "delete", id });
      } catch (e) {
        setTrips(prevTrips);
        showToast("Couldn't delete that trip.");
      }
    },
    [trips, showToast]
  );

  const bulkDeleteTrips = useCallback(
    async (ids) => {
      if (!sheetConfigured || ids.length === 0) return;
      const prevTrips = trips;
      const idSet = new Set(ids);
      setTrips((p) => p.filter((t) => !idSet.has(t.id)));
      try {
        await Promise.all(ids.map((id) => postToSheet({ action: "delete", id })));
        showToast(`Deleted ${ids.length} trip${ids.length === 1 ? "" : "s"}.`);
      } catch (e) {
        setTrips(prevTrips);
        showToast("Couldn't delete some trips — try again.");
      }
    },
    [trips, showToast]
  );

  const importTrips = useCallback(
    async (file) => {
      if (!file) return;
      if (!sheetConfigured) { showToast("Connect your Google Sheet first — see the setup notes at the top of this file."); return; }
      try {
        const text = await file.text();
        const data = JSON.parse(text);
        if (!Array.isArray(data)) throw new Error("not an array");
        let skipped = 0;
        const rows = [];
        data.forEach((r) => {
          const country = (r?.country || "").toString().trim();
          const start = r?.start || r?.start_date;
          const end = r?.end || r?.end_date;
          const type = r?.type === "business" ? "business" : "personal";
          const notes = (r?.notes || "").toString();
          const stops = Array.isArray(r?.stops) ? r.stops : (r?.stops ? String(r.stops).split(",").map((s) => s.trim()).filter(Boolean) : []);
          if (!country || !start || !end || validateTrip({ country, start, end })) { skipped++; return; }
          rows.push({ id: uid(), country, start_date: start, end_date: end, trip_type: type, notes, stops: stops.join(", ") });
        });
        if (rows.length === 0) { showToast("No valid trips found in that file."); return; }
        await postToSheet({ action: "bulkAdd", rows });
        setTrips((prev) => [
          ...rows.map((r) => ({
            id: r.id, country: r.country, start: r.start_date, end: r.end_date, type: r.trip_type,
            notes: r.notes, stops: r.stops ? r.stops.split(",").map((s) => s.trim()).filter(Boolean) : [],
          })),
          ...prev,
        ]);
        showToast(`Imported ${rows.length} trip${rows.length === 1 ? "" : "s"}${skipped ? `, skipped ${skipped}` : ""}.`);
      } catch (e) {
        showToast("Couldn't read or import that file.");
      }
    },
    [showToast]
  );

  const tabs = [
    { id: "log", label: "Log Trip", icon: PlusCircle },
    { id: "history", label: "History", icon: ListOrdered },
    { id: "insights", label: "Insights", icon: BarChart3 },
    { id: "map", label: "Stamp Wall", icon: Stamp },
    { id: "worldmap", label: "World Map", icon: MapIcon },
  ];

  return (
    <div className="min-h-screen w-full" style={{ background: BG, color: TEXT }}>
      <FontStyles />
      {!sheetConfigured && <ConfigBanner />}

      <div className="mx-auto max-w-5xl px-4 sm:px-6 pb-24">
        <Header syncing={syncing} lastSynced={lastSynced} onRefresh={() => refresh(true)} />

        <nav className="sticky top-0 z-20 -mx-4 sm:-mx-6 px-4 sm:px-6 pt-2 pb-3 backdrop-blur-md" style={{ background: `${BG}E8` }}>
          <div className="flex gap-1 sm:gap-2 overflow-x-auto rounded-xl p-1" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
            {tabs.map(({ id, label, icon: Icon }) => {
              const active = tab === id;
              return (
                <button key={id} onClick={() => setTab(id)}
                  className="flex items-center gap-1.5 whitespace-nowrap rounded-lg px-3 sm:px-4 py-2 text-sm font-medium transition-all duration-200"
                  style={{ background: active ? SURFACE_RAISED : "transparent", color: active ? ACCENT : MUTED }}>
                  <Icon size={16} strokeWidth={2.2} />
                  <span style={{ fontFamily: SANS }}>{label}</span>
                </button>
              );
            })}
          </div>
        </nav>

        <main className="mt-6">
          {loading ? (
            <FullScreenLoader label="Loading trips from your sheet…" compact />
          ) : (
            <>
              {tab === "log" && <LogTab trips={trips} onAdd={addTrip} onJump={setTab} />}
              {tab === "history" && <HistoryTab trips={trips} onDelete={deleteTrip} onBulkDelete={bulkDeleteTrips} onEdit={setEditingTrip} onImport={importTrips} />}
              {tab === "insights" && <InsightsTab trips={trips} />}
              {tab === "map" && <MapTab trips={trips} onEdit={setEditingTrip} />}
              {tab === "worldmap" && <WorldMapTab trips={trips} onEdit={setEditingTrip} />}
            </>
          )}
        </main>
      </div>

      {editingTrip && (
        <TripEditModal trip={editingTrip} onClose={() => setEditingTrip(null)} onSave={(fields) => updateTrip(editingTrip.id, fields)} />
      )}

      {toast && (
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 rounded-full px-4 py-2.5 text-sm shadow-lg animate-[fadein_.2s_ease-out]"
          style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: TEXT, fontFamily: SANS }}>
          {toast}
        </div>
      )}
    </div>
  );
}

function FontStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');
      @keyframes fadein { from { opacity:0; transform: translate(-50%, 6px);} to { opacity:1; transform: translate(-50%, 0);} }
      * { box-sizing: border-box; }
      body { font-family: ${SANS}; }
      input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(0.7); cursor: pointer; }
      ::selection { background: ${ACCENT}; color: #1a1a1a; }
    `}</style>
  );
}

function ConfigBanner() {
  return (
    <div className="flex items-center gap-2 px-4 py-2.5 text-xs sm:text-sm" style={{ background: DANGER_DIM, color: DANGER, borderBottom: `1px solid ${DANGER}55` }}>
      <CloudOff size={15} className="shrink-0" />
      <span>
        Not connected to a Google Sheet yet — follow the setup notes at the top of this file, then paste your
        deployed Apps Script URL into <code>SHEET_WEB_APP_URL</code>.
      </span>
    </div>
  );
}

function FullScreenLoader({ label, compact }) {
  return (
    <div className={`flex items-center justify-center ${compact ? "py-20" : "min-h-screen"}`} style={{ color: MUTED }}>
      <Loader2 className="animate-spin mr-2" size={18} /> {label}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Header + sync status                                                     */
/* ---------------------------------------------------------------------- */

function Header({ syncing, lastSynced, onRefresh }) {
  return (
    <header className="pt-8 pb-4 flex items-center justify-between gap-3">
      <div className="flex items-center gap-3 min-w-0">
        <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full" style={{ background: ACCENT_DIM, color: ACCENT }}>
          <Plane size={20} strokeWidth={2.2} />
        </div>
        <div className="min-w-0">
          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "1.6rem", letterSpacing: "-0.01em" }}>Waypoints</h1>
          <p className="text-sm mt-0.5 truncate" style={{ color: MUTED }}>Track every border you've crossed.</p>
        </div>
      </div>
      <button onClick={onRefresh} title="Refresh from Google Sheet"
        className="flex items-center gap-1.5 rounded-full px-3 py-2 text-xs font-medium shrink-0"
        style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: MUTED }}>
        <RefreshCw size={13} className={syncing ? "animate-spin" : ""} style={{ color: syncing ? ACCENT : MUTED }} />
        <span className="hidden sm:inline">{syncing ? "Syncing…" : `Synced ${timeAgo(lastSynced)}`}</span>
      </button>
    </header>
  );
}

/* ---------------------------------------------------------------------- */
/* Shared trip-form fields (used by Log Trip + Edit modal)                  */
/* ---------------------------------------------------------------------- */

function TripFormFields({ country, setCountry, start, setStart, end, setEnd, type, setType, notes, setNotes, stops, setStops }) {
  return (
    <div className="space-y-4">
      <Field label="Country">
        <div className="relative">
          <MapPin size={16} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
          <input
            list="wp-country-list"
            value={country}
            onChange={(e) => setCountry(e.target.value)}
            placeholder="e.g. Japan"
            className="w-full rounded-lg pl-9 pr-3 py-2.5 text-sm outline-none"
            style={inputStyle}
          />
          <datalist id="wp-country-list">
            {COUNTRY_LIST.map((c) => <option key={c} value={c} />)}
          </datalist>
        </div>
      </Field>

      <div className="grid grid-cols-2 gap-4">
        <Field label="Start date">
          <input type="date" value={start} max={end || undefined} onChange={(e) => setStart(e.target.value)}
            className="w-full rounded-lg px-3 py-2.5 text-sm outline-none" style={inputStyle} />
        </Field>
        <Field label="End date">
          <input type="date" value={end} min={start || undefined} onChange={(e) => setEnd(e.target.value)}
            className="w-full rounded-lg px-3 py-2.5 text-sm outline-none" style={inputStyle} />
        </Field>
      </div>

      <Field label="Trip type">
        <div className="grid grid-cols-2 gap-2">
          {Object.entries(TRIP_TYPES).map(([key, meta]) => {
            const Icon = meta.icon;
            const active = type === key;
            return (
              <button key={key} type="button" onClick={() => setType(key)}
                className="flex items-center justify-center gap-1.5 rounded-lg py-2.5 text-sm font-medium transition-colors"
                style={{ background: active ? `${meta.color}22` : SURFACE_RAISED, border: `1px solid ${active ? meta.color : BORDER}`, color: active ? meta.color : MUTED }}>
                <Icon size={15} /> {meta.label}
              </button>
            );
          })}
        </div>
      </Field>

      {setNotes && (
        <Field label="Notes (optional)">
          <textarea
            value={notes}
            onChange={(e) => setNotes(e.target.value)}
            placeholder="e.g. Conference, family visit…"
            rows={2}
            className="w-full rounded-lg px-3 py-2.5 text-sm outline-none resize-none"
            style={inputStyle}
          />
        </Field>
      )}

      {setStops && (
        <Field label="Other stops (optional, reference only)">
          <input
            value={stops}
            onChange={(e) => setStops(e.target.value)}
            placeholder="e.g. Layover in Singapore"
            className="w-full rounded-lg px-3 py-2.5 text-sm outline-none"
            style={inputStyle}
          />
          <span className="block text-xs mt-1" style={{ color: MUTED }}>
            Comma-separated. Shown for context only — day counts still track the main country above.
          </span>
        </Field>
      )}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs mb-1.5" style={{ color: MUTED, fontFamily: SANS }}>{label}</span>
      {children}
    </label>
  );
}

/* ---------------------------------------------------------------------- */
/* Log Trip tab                                                             */
/* ---------------------------------------------------------------------- */

function LogTab({ trips, onAdd, onJump }) {
  const [country, setCountry] = useState("");
  const [start, setStart] = useState("");
  const [end, setEnd] = useState("");
  const [type, setType] = useState("personal");
  const [notes, setNotes] = useState("");
  const [stopsText, setStopsText] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const previewDays = start && end && !validateTrip({ country: country || "x", start, end }) ? inclusiveDays(start, end) : null;

  const handleSubmit = async (e) => {
    e.preventDefault();
    const err = validateTrip({ country, start, end });
    if (err) { setError(err); return; }
    setError("");
    setSaving(true);
    const stops = stopsText.split(",").map((s) => s.trim()).filter(Boolean);
    await onAdd({ country: country.trim(), start, end, type, notes: notes.trim(), stops });
    setSaving(false);
    setCountry(""); setStart(""); setEnd(""); setType("personal"); setNotes(""); setStopsText("");
  };

  const recent = [...trips].sort((a, b) => toDate(b.start) - toDate(a.start)).slice(0, 4);

  return (
    <div className="grid gap-6 lg:grid-cols-5">
      <form onSubmit={handleSubmit} className="lg:col-span-3 rounded-2xl p-6" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
        <h2 style={{ fontFamily: SERIF, fontSize: "1.15rem", fontWeight: 600 }}>New trip</h2>
        <p className="text-sm mt-1 mb-5" style={{ color: MUTED }}>Add a country and the dates you were there. Days are counted inclusively.</p>

        <TripFormFields
          country={country} setCountry={setCountry} start={start} setStart={setStart} end={end} setEnd={setEnd}
          type={type} setType={setType} notes={notes} setNotes={setNotes} stops={stopsText} setStops={setStopsText}
        />

        {previewDays !== null && (
          <div className="text-sm rounded-lg px-3 py-2 mt-4" style={{ background: ACCENT_DIM, color: ACCENT }}>
            {previewDays} {previewDays === 1 ? "day" : "days"} in {country || "this country"}
          </div>
        )}
        {error && (
          <div className="flex items-center gap-2 text-sm rounded-lg px-3 py-2 mt-4" style={{ background: DANGER_DIM, color: DANGER }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}

        <button type="submit" disabled={saving}
          className="w-full mt-4 flex items-center justify-center gap-2 rounded-lg py-2.5 text-sm font-medium transition-opacity hover:opacity-90 disabled:opacity-60"
          style={{ background: ACCENT, color: "#1A1408" }}>
          {saving && <Loader2 size={15} className="animate-spin" />}
          Add trip
        </button>
      </form>

      <div className="lg:col-span-2 rounded-2xl p-6" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
        <h3 style={{ fontFamily: SERIF, fontSize: "1.05rem", fontWeight: 600 }}>Recently logged</h3>
        {recent.length === 0 ? (
          <p className="text-sm mt-3" style={{ color: MUTED }}>Nothing yet — your first trip will show up here.</p>
        ) : (
          <ul className="mt-4 space-y-3">
            {recent.map((t) => (
              <li key={t.id} className="flex items-center justify-between text-sm gap-2">
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5">
                    <span aria-hidden="true">{countryFlag(t.country)}</span>
                    <span style={{ fontFamily: SANS, fontWeight: 500 }} className="truncate">{t.country}</span>
                    <TypeBadge type={t.type} />
                  </div>
                  <div style={{ color: MUTED, fontSize: "0.78rem" }}>{formatRange(t.start, t.end)}</div>
                </div>
                <span className="shrink-0 rounded-full px-2.5 py-1 text-xs font-medium" style={{ background: SURFACE_RAISED, color: TEAL }}>
                  {inclusiveDays(t.start, t.end)}d
                </span>
              </li>
            ))}
          </ul>
        )}
        {trips.length > 0 && (
          <button onClick={() => onJump("history")} className="mt-5 text-sm font-medium" style={{ color: ACCENT }}>
            View all {trips.length} trips →
          </button>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Edit modal                                                               */
/* ---------------------------------------------------------------------- */

function TripEditModal({ trip, onClose, onSave }) {
  const [country, setCountry] = useState(trip.country);
  const [start, setStart] = useState(trip.start);
  const [end, setEnd] = useState(trip.end);
  const [type, setType] = useState(trip.type || "personal");
  const [notes, setNotes] = useState(trip.notes || "");
  const [stopsText, setStopsText] = useState((trip.stops || []).join(", "));
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    const err = validateTrip({ country, start, end });
    if (err) { setError(err); return; }
    setError("");
    setSaving(true);
    const stops = stopsText.split(",").map((s) => s.trim()).filter(Boolean);
    await onSave({ country: country.trim(), start, end, type, notes: notes.trim(), stops });
    setSaving(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: "#00000099" }} onClick={onClose}>
      <form onClick={(e) => e.stopPropagation()} onSubmit={submit} className="w-full max-w-md rounded-2xl p-6" style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}` }}>
        <div className="flex items-center justify-between mb-4">
          <h3 style={{ fontFamily: SERIF, fontSize: "1.15rem", fontWeight: 600 }}>Edit trip</h3>
          <button type="button" onClick={onClose} style={{ color: MUTED }}><X size={18} /></button>
        </div>

        <TripFormFields
          country={country} setCountry={setCountry} start={start} setStart={setStart} end={end} setEnd={setEnd}
          type={type} setType={setType} notes={notes} setNotes={setNotes} stops={stopsText} setStops={setStopsText}
        />

        {error && (
          <div className="flex items-center gap-2 text-sm rounded-lg px-3 py-2 mt-4" style={{ background: DANGER_DIM, color: DANGER }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}

        <div className="flex gap-2 mt-5">
          <button type="button" onClick={onClose} className="flex-1 rounded-lg py-2.5 text-sm font-medium" style={{ background: SURFACE, border: `1px solid ${BORDER}`, color: MUTED }}>
            Cancel
          </button>
          <button type="submit" disabled={saving}
            className="flex-1 flex items-center justify-center gap-2 rounded-lg py-2.5 text-sm font-medium transition-opacity hover:opacity-90 disabled:opacity-60"
            style={{ background: ACCENT, color: "#1A1408" }}>
            {saving && <Loader2 size={15} className="animate-spin" />}
            Save changes
          </button>
        </div>
      </form>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* History tab                                                              */
/* ---------------------------------------------------------------------- */

function HistoryTab({ trips, onDelete, onBulkDelete, onEdit, onImport }) {
  const [confirmId, setConfirmId] = useState(null);
  const [typeFilter, setTypeFilter] = useState("all");
  const [query, setQuery] = useState("");
  const [sortBy, setSortBy] = useState("newest"); // newest | oldest | country | longest
  const [selected, setSelected] = useState(() => new Set());
  const [bulkConfirm, setBulkConfirm] = useState(false);
  const fileInputRef = useRef(null);

  const sorted = useMemo(() => {
    const q = query.trim().toLowerCase();
    let list = trips
      .filter((t) => typeFilter === "all" || (t.type || "personal") === typeFilter)
      .filter((t) => !q || t.country.toLowerCase().includes(q));
    switch (sortBy) {
      case "oldest": list = list.sort((a, b) => toDate(a.start) - toDate(b.start)); break;
      case "country": list = list.sort((a, b) => a.country.localeCompare(b.country)); break;
      case "longest": list = list.sort((a, b) => inclusiveDays(b.start, b.end) - inclusiveDays(a.start, a.end)); break;
      default: list = list.sort((a, b) => toDate(b.start) - toDate(a.start));
    }
    return list;
  }, [trips, typeFilter, query, sortBy]);

  // Drop any selected ids that no longer match the current filter/search
  useEffect(() => {
    setSelected((prev) => {
      const visible = new Set(sorted.map((t) => t.id));
      const next = new Set([...prev].filter((id) => visible.has(id)));
      return next.size === prev.size ? prev : next;
    });
  }, [sorted]);

  const toggleSelect = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const toggleSelectAll = () => {
    setSelected((prev) => (prev.size === sorted.length ? new Set() : new Set(sorted.map((t) => t.id))));
  };

  const handleExport = () => {
    const blob = new Blob([JSON.stringify(trips, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `waypoints-trips-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleImportChange = (e) => {
    const file = e.target.files?.[0];
    if (file) onImport(file);
    e.target.value = "";
  };

  if (trips.length === 0) {
    return (
      <div className="space-y-4">
        <EmptyState text="No trips logged yet. Add one from the Log Trip tab, or import a previous export." />
        <div className="flex justify-center">
          <button onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium"
            style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: MUTED }}>
            <Upload size={14} /> Import trips
          </button>
          <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={handleImportChange} />
        </div>
      </div>
    );
  }

  const SORT_OPTIONS = [
    { id: "newest", label: "Newest first" },
    { id: "oldest", label: "Oldest first" },
    { id: "country", label: "Country A–Z" },
    { id: "longest", label: "Longest first" },
  ];

  return (
    <div className="rounded-2xl overflow-hidden" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
      <div className="px-6 py-4 space-y-3" style={{ borderBottom: `1px solid ${BORDER}` }}>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 style={{ fontFamily: SERIF, fontSize: "1.15rem", fontWeight: 600 }}>All trips</h2>
            <p className="text-sm mt-1" style={{ color: MUTED }}>{sorted.length} of {trips.length} trips</p>
          </div>
          <div className="flex items-center gap-2">
            <TypeFilter value={typeFilter} onChange={setTypeFilter} />
            <button onClick={handleExport} title="Export trips as JSON"
              className="flex items-center gap-1.5 rounded-lg px-2.5 py-2 text-xs font-medium"
              style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: MUTED }}>
              <Download size={14} /> <span className="hidden sm:inline">Export</span>
            </button>
            <button onClick={() => fileInputRef.current?.click()} title="Import trips from JSON"
              className="flex items-center gap-1.5 rounded-lg px-2.5 py-2 text-xs font-medium"
              style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: MUTED }}>
              <Upload size={14} /> <span className="hidden sm:inline">Import</span>
            </button>
            <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={handleImportChange} />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="relative flex-1 min-w-[160px]">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search by country…"
              className="w-full rounded-lg pl-8 pr-3 py-2 text-sm outline-none"
              style={inputStyle}
            />
          </div>
          <div className="relative">
            <select value={sortBy} onChange={(e) => setSortBy(e.target.value)}
              className="appearance-none rounded-lg pl-3 pr-8 py-2 text-sm outline-none cursor-pointer"
              style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: TEXT }}>
              {SORT_OPTIONS.map((o) => <option key={o.id} value={o.id}>{o.label}</option>)}
            </select>
            <ChevronDown size={13} className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
          </div>
        </div>

        {selected.size > 0 && (
          <div className="flex items-center justify-between gap-3 rounded-lg px-3 py-2" style={{ background: DANGER_DIM }}>
            <span className="text-sm" style={{ color: TEXT }}>{selected.size} selected</span>
            {bulkConfirm ? (
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => { onBulkDelete([...selected]); setSelected(new Set()); setBulkConfirm(false); }}
                  className="rounded-md px-2.5 py-1 text-xs font-medium" style={{ background: DANGER, color: "#1A0A0C" }}>
                  Confirm delete
                </button>
                <button onClick={() => setBulkConfirm(false)} className="rounded-md px-2.5 py-1 text-xs" style={{ color: MUTED }}>Cancel</button>
              </div>
            ) : (
              <button onClick={() => setBulkConfirm(true)} className="flex items-center gap-1.5 rounded-md px-2.5 py-1 text-xs font-medium" style={{ color: DANGER }}>
                <Trash2 size={13} /> Delete selected
              </button>
            )}
          </div>
        )}
      </div>

      {sorted.length === 0 ? (
        <div className="px-6 py-10 text-sm text-center" style={{ color: MUTED }}>No trips match this search/filter.</div>
      ) : (
        <>
          <div className="flex items-center gap-2 px-6 py-2 text-xs" style={{ color: MUTED, borderBottom: `1px solid ${BORDER}` }}>
            <input type="checkbox" checked={selected.size === sorted.length} onChange={toggleSelectAll} className="accent-current" />
            Select all
          </div>
          <ul>
            {sorted.map((t, i) => (
              <li key={t.id} className="flex items-center justify-between gap-3 px-6 py-4" style={{ borderBottom: i < sorted.length - 1 ? `1px solid ${BORDER}` : "none" }}>
                <div className="flex items-center gap-3 min-w-0">
                  <input type="checkbox" checked={selected.has(t.id)} onChange={() => toggleSelect(t.id)} className="accent-current shrink-0" />
                  <span className="hidden sm:inline text-lg shrink-0" aria-hidden="true">{countryFlag(t.country)}</span>
                  <div className="min-w-0">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <span style={{ fontFamily: SANS, fontWeight: 500 }} className="truncate">{t.country}</span>
                      <TypeBadge type={t.type} />
                      {t.notes && <StickyNote size={12} style={{ color: MUTED }} aria-label="Has notes" />}
                    </div>
                    <div className="flex items-center gap-1.5 text-xs mt-0.5" style={{ color: MUTED }}>
                      <CalendarDays size={12} /> {formatRange(t.start, t.end)}
                    </div>
                    {t.stops && t.stops.length > 0 && (
                      <div className="text-xs mt-0.5" style={{ color: MUTED }}>Also: {t.stops.join(", ")}</div>
                    )}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <span className="rounded-full px-2.5 py-1 text-xs font-medium" style={{ background: SURFACE_RAISED, color: TEAL }}>
                    {inclusiveDays(t.start, t.end)}d
                  </span>
                  <button onClick={() => onEdit(t)} aria-label={`Edit trip to ${t.country}`} className="rounded-md p-1.5" style={{ color: MUTED }}>
                    <Pencil size={16} />
                  </button>
                  {confirmId === t.id ? (
                    <div className="flex items-center gap-1.5">
                      <button onClick={() => { onDelete(t.id); setConfirmId(null); }} className="rounded-md px-2 py-1 text-xs font-medium" style={{ background: DANGER, color: "#1A0A0C" }}>
                        Confirm
                      </button>
                      <button onClick={() => setConfirmId(null)} className="rounded-md px-2 py-1 text-xs" style={{ color: MUTED }}>Cancel</button>
                    </div>
                  ) : (
                    <button onClick={() => setConfirmId(t.id)} aria-label={`Delete trip to ${t.country}`} className="rounded-md p-1.5" style={{ color: MUTED }}>
                      <Trash2 size={16} />
                    </button>
                  )}
                </div>
              </li>
            ))}
          </ul>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Insights tab                                                             */
/* ---------------------------------------------------------------------- */

function InsightsTab({ trips }) {
  const years = useMemo(() => {
    const s = new Set();
    trips.forEach((t) => { s.add(toDate(t.start).getFullYear()); s.add(toDate(t.end).getFullYear()); });
    s.add(new Date().getFullYear());
    return [...s].sort((a, b) => b - a);
  }, [trips]);

  const [year, setYear] = useState(years[0]);
  const [typeFilter, setTypeFilter] = useState("all");
  const [windowMode, setWindowMode] = useState("calendar"); // calendar | rolling
  const [settings, setSettings] = useState(loadSettings);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [editingThreshold, setEditingThreshold] = useState(null); // country currently being edited
  useEffect(() => { if (!years.includes(year)) setYear(years[0]); }, [years]); // eslint-disable-line

  const updateSettings = (next) => {
    setSettings(next);
    saveSettings(next);
  };

  const rollingRange = useMemo(() => {
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 364);
    return { start, end };
  }, []);

  const byCountry = useMemo(() => {
    const map = {};
    trips.filter((t) => typeFilter === "all" || (t.type || "personal") === typeFilter).forEach((t) => {
      const d = windowMode === "rolling"
        ? overlapDaysInRollingWindow(t, rollingRange.start, rollingRange.end)
        : overlapDaysInYear(t, year);
      if (d <= 0) return;
      const kind = t.type || "personal";
      if (!map[t.country]) map[t.country] = { country: t.country, days: 0, personal: 0, business: 0 };
      map[t.country].days += d;
      map[t.country][kind] += d;
    });
    return Object.values(map).sort((a, b) => b.days - a.days);
  }, [trips, year, typeFilter, windowMode, rollingRange]);

  const thresholdFor = (country) => settings.thresholds[country] || RESIDENCY_THRESHOLD;
  const isHome = (country) => settings.homeCountry && country === settings.homeCountry;

  const totalDays = byCountry.reduce((s, c) => s + c.days, 0);
  const overThreshold = byCountry.filter((c) => !isHome(c.country) && c.days > thresholdFor(c.country));

  const periodLabel = windowMode === "rolling" ? "the last 365 days" : `${year}`;

  const saveThreshold = (country, value) => {
    const num = parseInt(value, 10);
    const next = { ...settings, thresholds: { ...settings.thresholds } };
    if (!num || num === RESIDENCY_THRESHOLD) delete next.thresholds[country];
    else next.thresholds[country] = num;
    updateSettings(next);
    setEditingThreshold(null);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 style={{ fontFamily: SERIF, fontSize: "1.3rem", fontWeight: 600 }}>Insights</h2>
          <p className="text-sm mt-1" style={{ color: MUTED }}>Total days per country, {periodLabel}</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <TypeFilter value={typeFilter} onChange={setTypeFilter} />
          <div className="flex gap-1 rounded-lg p-1" style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}` }}>
            {[{ id: "calendar", label: "Calendar year" }, { id: "rolling", label: "Rolling 365d" }].map((o) => (
              <button key={o.id} onClick={() => setWindowMode(o.id)}
                className="rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors"
                style={{ background: windowMode === o.id ? SURFACE : "transparent", color: windowMode === o.id ? ACCENT : MUTED }}>
                {o.label}
              </button>
            ))}
          </div>
          {windowMode === "calendar" && (
            <div className="relative">
              <select value={year} onChange={(e) => setYear(Number(e.target.value))}
                className="appearance-none rounded-lg pl-4 pr-9 py-2.5 text-sm font-medium outline-none cursor-pointer"
                style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: TEXT }}>
                {years.map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
              <ChevronDown size={15} className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
            </div>
          )}
          <button onClick={() => setSettingsOpen((s) => !s)} title="Home country & thresholds"
            className="flex items-center justify-center rounded-lg p-2.5"
            style={{ background: settingsOpen ? SURFACE : SURFACE_RAISED, border: `1px solid ${BORDER}`, color: settingsOpen ? ACCENT : MUTED }}>
            <Settings2 size={15} />
          </button>
        </div>
      </div>

      {settingsOpen && (
        <div className="rounded-xl p-4 flex flex-wrap items-center gap-3" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
          <label className="flex items-center gap-2 text-sm" style={{ color: MUTED }}>
            Home country:
            <select
              value={settings.homeCountry}
              onChange={(e) => updateSettings({ ...settings, homeCountry: e.target.value })}
              className="rounded-lg px-2.5 py-1.5 text-sm outline-none"
              style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}`, color: TEXT }}>
              <option value="">None</option>
              {COUNTRY_LIST.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </label>
          <span className="text-xs" style={{ color: MUTED }}>
            Your home country never triggers the threshold warning. Click any day count below to set a custom threshold for that country (default {RESIDENCY_THRESHOLD}).
          </span>
        </div>
      )}

      {overThreshold.length > 0 && (
        <div className="flex items-start gap-3 rounded-xl px-4 py-3" style={{ background: DANGER_DIM, border: `1px solid ${DANGER}55` }}>
          <AlertTriangle size={18} style={{ color: DANGER }} className="shrink-0 mt-0.5" />
          <div className="text-sm" style={{ color: TEXT }}>
            <span style={{ fontWeight: 600 }}>Threshold exceeded</span> in {periodLabel}: {overThreshold.map((c) => `${c.country} (${thresholdFor(c.country)}d)`).join(", ")}.
          </div>
        </div>
      )}

      {byCountry.length === 0 ? (
        <EmptyState text={windowMode === "rolling" ? "No trips in the last 365 days." : `No trips recorded for ${year}.`} />
      ) : (
        <div className="rounded-2xl p-6" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
          <div className="text-sm mb-5" style={{ color: MUTED }}>
            {totalDays} total travel {totalDays === 1 ? "day" : "days"} across {byCountry.length} {byCountry.length === 1 ? "country" : "countries"}
          </div>
          <div className="space-y-4">
            {byCountry.map(({ country, days, personal, business }) => {
              const threshold = thresholdFor(country);
              const home = isHome(country);
              const over = !home && days > threshold;
              const personalPct = Math.min(100, (personal / threshold) * 100);
              const businessPct = Math.min(100 - personalPct, (business / threshold) * 100);
              const editing = editingThreshold === country;
              return (
                <div key={country}>
                  <div className="flex items-center justify-between text-sm mb-1.5">
                    <span style={{ fontFamily: SANS, fontWeight: 500, color: TEXT }} className="flex items-center gap-1.5">
                      <span aria-hidden="true">{countryFlag(country)}</span>
                      {country}
                      {home && <span className="text-xs rounded-full px-1.5 py-0.5" style={{ background: ACCENT_DIM, color: ACCENT }}>home</span>}
                      {over && <AlertTriangle size={13} style={{ color: DANGER }} />}
                    </span>
                    {editing ? (
                      <input
                        type="number"
                        autoFocus
                        defaultValue={threshold}
                        onBlur={(e) => saveThreshold(country, e.target.value)}
                        onKeyDown={(e) => { if (e.key === "Enter") saveThreshold(country, e.target.value); if (e.key === "Escape") setEditingThreshold(null); }}
                        className="w-16 rounded-md px-1.5 py-0.5 text-xs text-right outline-none"
                        style={inputStyle}
                      />
                    ) : (
                      <button onClick={() => setEditingThreshold(country)} className="text-right" style={{ color: over ? DANGER : MUTED, fontWeight: over ? 600 : 400 }}>
                        {days}d{typeFilter === "all" && personal > 0 && business > 0 && (
                          <span style={{ color: MUTED, fontWeight: 400 }}> ({personal}p / {business}b)</span>
                        )}
                        <span style={{ color: MUTED, fontWeight: 400 }}> / {threshold}</span>
                      </button>
                    )}
                  </div>
                  <div className="h-2 rounded-full overflow-hidden flex" style={{ background: SURFACE_RAISED }}>
                    <div className="h-full transition-all duration-500" style={{ width: `${personalPct}%`, background: over ? DANGER : TEAL }} />
                    <div className="h-full transition-all duration-500" style={{ width: `${businessPct}%`, background: over ? `${DANGER}AA` : BUSINESS }} />
                  </div>
                </div>
              );
            })}
          </div>
          <div className="mt-5 pt-4 flex flex-wrap items-center gap-x-4 gap-y-1.5 text-xs" style={{ borderTop: `1px solid ${BORDER}`, color: MUTED }}>
            <span className="flex items-center gap-1.5"><span className="inline-block h-2 w-2 rounded-full" style={{ background: TEAL }} /> personal</span>
            <span className="flex items-center gap-1.5"><span className="inline-block h-2 w-2 rounded-full" style={{ background: BUSINESS }} /> business</span>
            <span className="flex items-center gap-1.5"><span className="inline-block h-2 w-2 rounded-full" style={{ background: DANGER }} /> over threshold</span>
            <span>Click a day count to set a custom threshold.</span>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Map / Stamp Wall tab                                                     */
/* ---------------------------------------------------------------------- */

function MapTab({ trips, onEdit }) {
  const [selected, setSelected] = useState(null);
  const [typeFilter, setTypeFilter] = useState("all");

  const filteredTrips = useMemo(() => trips.filter((t) => typeFilter === "all" || (t.type || "personal") === typeFilter), [trips, typeFilter]);

  const byCountry = useMemo(() => {
    const map = {};
    filteredTrips.forEach((t) => {
      const d = inclusiveDays(t.start, t.end);
      if (!map[t.country]) map[t.country] = { country: t.country, days: 0, trips: [] };
      map[t.country].days += d;
      map[t.country].trips.push(t);
    });
    return map;
  }, [filteredTrips]);

  useEffect(() => {
    if (selected && byCountry[selected.country]) setSelected(byCountry[selected.country]);
    else if (selected) setSelected(null);
  }, [byCountry]); // eslint-disable-line

  const grouped = useMemo(() => {
    const g = {};
    Object.values(byCountry).forEach((entry) => { (g[continentOf(entry.country)] = g[continentOf(entry.country)] || []).push(entry); });
    Object.values(g).forEach((arr) => arr.sort((a, b) => b.days - a.days));
    return g;
  }, [byCountry]);

  const maxDays = Math.max(1, ...Object.values(byCountry).map((e) => e.days));
  const visitedCount = Object.keys(byCountry).length;

  if (trips.length === 0) return <EmptyState text="Log a trip to start filling your stamp wall." />;

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 style={{ fontFamily: SERIF, fontSize: "1.3rem", fontWeight: 600 }}>Stamp wall</h2>
          <p className="text-sm mt-1" style={{ color: MUTED }}>{visitedCount} {visitedCount === 1 ? "country" : "countries"} stamped — larger and brighter means more days spent there.</p>
        </div>
        <TypeFilter value={typeFilter} onChange={setTypeFilter} />
      </div>

      {visitedCount === 0 ? (
        <EmptyState text="No trips match this filter." />
      ) : (
        <div className="space-y-8">
          {CONTINENT_ORDER.filter((c) => grouped[c]?.length).map((continent) => (
            <div key={continent}>
              <div className="flex items-center gap-2 mb-3">
                <span className="h-2 w-2 rounded-full" style={{ background: CONTINENT_ACCENT[continent] }} />
                <h3 className="text-sm font-medium" style={{ color: MUTED, fontFamily: SANS }}>{continent}</h3>
              </div>
              <div className="flex flex-wrap gap-3">
                {grouped[continent].map((entry) => {
                  const intensity = 0.35 + 0.65 * (entry.days / maxDays);
                  const size = 84 + 46 * (entry.days / maxDays);
                  const color = CONTINENT_ACCENT[continent];
                  return (
                    <button key={entry.country} onClick={() => setSelected(entry)}
                      className="relative flex flex-col items-center justify-center rounded-2xl transition-transform hover:-translate-y-0.5"
                      style={{ width: size, height: size, background: SURFACE, border: `2px dashed ${color}${Math.round(intensity * 255).toString(16).padStart(2, "0")}` }}>
                      <Stamp size={18} style={{ color, opacity: intensity }} />
                      <span className="mt-1.5 text-xs font-medium text-center px-2 leading-tight" style={{ color: TEXT }}>{entry.country}</span>
                      <span className="text-[10px] mt-0.5" style={{ color: MUTED }}>{entry.days}d</span>
                      {entry.days > RESIDENCY_THRESHOLD && (
                        <span className="absolute -top-1.5 -right-1.5 flex h-5 w-5 items-center justify-center rounded-full" style={{ background: DANGER }}>
                          <AlertTriangle size={11} color="#1A0A0C" />
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}

      {selected && (
        <CountryTripsModal entry={selected} onClose={() => setSelected(null)} onEdit={(t) => { setSelected(null); onEdit(t); }} />
      )}
    </div>
  );
}

function CountryTripsModal({ entry, onClose, onEdit }) {
  return (
    <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center p-4" style={{ background: "#00000099" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full max-w-md rounded-2xl p-6" style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}` }}>
        <div className="flex items-start justify-between">
          <div>
            <h3 style={{ fontFamily: SERIF, fontSize: "1.2rem", fontWeight: 600 }}>{entry.country}</h3>
            <p className="text-sm mt-1" style={{ color: MUTED }}>{entry.days} total days · {entry.trips.length} {entry.trips.length === 1 ? "trip" : "trips"}</p>
          </div>
          <button onClick={onClose} style={{ color: MUTED }}><X size={18} /></button>
        </div>
        <ul className="mt-4 space-y-2 max-h-64 overflow-y-auto">
          {[...entry.trips].sort((a, b) => toDate(b.start) - toDate(a.start)).map((t) => (
            <li key={t.id} className="flex items-center justify-between gap-2 text-sm rounded-lg px-3 py-2" style={{ background: SURFACE }}>
              <span className="flex items-center gap-2 min-w-0">
                <TypeBadge type={t.type} />
                <span style={{ color: MUTED }} className="truncate">{formatRange(t.start, t.end)}</span>
              </span>
              <span className="flex items-center gap-2 shrink-0">
                <span style={{ color: TEAL, fontWeight: 500 }}>{inclusiveDays(t.start, t.end)}d</span>
                <button onClick={() => onEdit(t)} style={{ color: MUTED }}><Pencil size={14} /></button>
              </span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* World Map tab                                                            */
/* ---------------------------------------------------------------------- */

function WorldMapTab({ trips, onEdit }) {
  const [selected, setSelected] = useState(null);
  const [typeFilter, setTypeFilter] = useState("all");

  const filteredTrips = useMemo(() => trips.filter((t) => typeFilter === "all" || (t.type || "personal") === typeFilter), [trips, typeFilter]);

  const byCountry = useMemo(() => {
    const map = {};
    filteredTrips.forEach((t) => {
      if (!COUNTRY_LATLON[t.country]) return;
      const d = inclusiveDays(t.start, t.end);
      if (!map[t.country]) map[t.country] = { country: t.country, days: 0, trips: [] };
      map[t.country].days += d;
      map[t.country].trips.push(t);
    });
    return map;
  }, [filteredTrips]);

  useEffect(() => {
    if (selected && byCountry[selected.country]) setSelected(byCountry[selected.country]);
    else if (selected) setSelected(null);
  }, [byCountry]); // eslint-disable-line

  const entries = Object.values(byCountry);
  const maxDays = Math.max(1, ...entries.map((e) => e.days));
  const unmapped = [...new Set(filteredTrips.filter((t) => !COUNTRY_LATLON[t.country]).map((t) => t.country))];
  const continentsShown = [...new Set(entries.map((e) => continentOf(e.country)))];

  if (trips.length === 0) return <EmptyState text="Log a trip to see it appear on the map." />;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 style={{ fontFamily: SERIF, fontSize: "1.3rem", fontWeight: 600 }}>World map</h2>
          <p className="text-sm mt-1" style={{ color: MUTED }}>Dot size and glow track days spent — click a country to see its trips.</p>
        </div>
        <TypeFilter value={typeFilter} onChange={setTypeFilter} />
      </div>

      {entries.length === 0 ? (
        <EmptyState text="No trips match this filter." />
      ) : (
        <div className="rounded-2xl p-4 sm:p-6" style={{ background: SURFACE, border: `1px solid ${BORDER}` }}>
          <ComposableMap projectionConfig={{ scale: 148 }} style={{ width: "100%", height: "auto" }}>
            <Geographies geography={WORLD_GEO_URL}>
              {({ geographies }) =>
                geographies.map((geo) => (
                  <Geography
                    key={geo.rsmKey}
                    geography={geo}
                    fill={SURFACE_RAISED}
                    stroke={BORDER}
                    strokeWidth={0.5}
                    style={{
                      default: { outline: "none" },
                      hover: { outline: "none", fill: SURFACE_RAISED },
                      pressed: { outline: "none" },
                    }}
                  />
                ))
              }
            </Geographies>

            {entries.map((entry) => {
              const [lat, lon] = COUNTRY_LATLON[entry.country];
              const ratio = entry.days / maxDays;
              const r = 5 + 9 * ratio;
              const over = entry.days > RESIDENCY_THRESHOLD;
              const color = CONTINENT_ACCENT[continentOf(entry.country)];
              return (
                <Marker key={entry.country} coordinates={[lon, lat]} onClick={() => setSelected(entry)} style={{ cursor: "pointer" }}>
                  <title>{`${entry.country} — ${entry.days}d`}</title>
                  <circle r={r + 5} fill={color} opacity={0.15} />
                  <circle r={r} fill={color} opacity={0.45 + 0.55 * ratio} stroke={over ? DANGER : "none"} strokeWidth={over ? 1.5 : 0} />
                </Marker>
              );
            })}
          </ComposableMap>

          <div className="mt-4 flex flex-wrap items-center gap-x-4 gap-y-1.5 text-xs" style={{ color: MUTED }}>
            {continentsShown.map((c) => (
              <span key={c} className="flex items-center gap-1.5">
                <span className="inline-block h-2 w-2 rounded-full" style={{ background: CONTINENT_ACCENT[c] }} /> {c}
              </span>
            ))}
            <span className="flex items-center gap-1.5">
              <span className="inline-block h-2 w-2 rounded-full" style={{ background: DANGER }} /> over {RESIDENCY_THRESHOLD} days
            </span>
          </div>
        </div>
      )}

      {unmapped.length > 0 && (
        <p className="text-xs" style={{ color: MUTED }}>
          Not shown on the map yet (no coordinates on file): {unmapped.join(", ")}
        </p>
      )}

      {selected && (
        <CountryTripsModal entry={selected} onClose={() => setSelected(null)} onEdit={(t) => { setSelected(null); onEdit(t); }} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Shared bits                                                              */
/* ---------------------------------------------------------------------- */

function EmptyState({ text }) {
  return (
    <div className="flex flex-col items-center justify-center text-center rounded-2xl py-16 px-6" style={{ background: SURFACE, border: `1px dashed ${BORDER}`, color: MUTED }}>
      <Plane size={22} className="mb-3" style={{ color: ACCENT }} />
      <p className="text-sm max-w-xs">{text}</p>
    </div>
  );
}

function TypeBadge({ type }) {
  const meta = TRIP_TYPES[type] || TRIP_TYPES.personal;
  const Icon = meta.icon;
  return (
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-medium shrink-0" style={{ background: `${meta.color}22`, color: meta.color }}>
      <Icon size={11} /> {meta.label}
    </span>
  );
}

function TypeFilter({ value, onChange }) {
  const options = [{ id: "all", label: "All" }, { id: "personal", label: "Personal", icon: Home }, { id: "business", label: "Business", icon: Briefcase }];
  return (
    <div className="flex gap-1 rounded-lg p-1" style={{ background: SURFACE_RAISED, border: `1px solid ${BORDER}` }}>
      {options.map((o) => {
        const active = value === o.id;
        const Icon = o.icon;
        const color = o.id === "personal" ? TEAL : o.id === "business" ? BUSINESS : ACCENT;
        return (
          <button key={o.id} onClick={() => onChange(o.id)} className="flex items-center gap-1 rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors"
            style={{ background: active ? SURFACE : "transparent", color: active ? color : MUTED }}>
            {Icon && <Icon size={12} />} {o.label}
          </button>
        );
      })}
    </div>
  );
}
