import React, { useState } from "react";

const selectStyle = {
  background: "#0f2419",
  border: "1px solid #2c4a3a",
  color: "#eaf3ec",
  borderRadius: 8,
  padding: "8px 10px",
  fontSize: 13,
  fontFamily: "ui-monospace, monospace",
  outline: "none",
  width: "100%",
};
const labelStyle = {
  fontSize: 10,
  letterSpacing: 2,
  color: "#6f8c7c",
  fontWeight: 700,
  fontFamily: "ui-monospace, monospace",
  display: "block",
  marginBottom: 4,
};
function chipStyle(active) {
  return {
    padding: "6px 10px",
    borderRadius: 999,
    border: `1px solid ${active ? "#f0c95a" : "#2c4a3a"}`,
    background: active ? "rgba(240,201,90,0.12)" : "transparent",
    color: active ? "#f0c95a" : "#9fb8a8",
    fontSize: 11,
    fontWeight: 700,
    cursor: "pointer",
    whiteSpace: "nowrap",
  };
}

const STALL_PRESETS = [
  { label: "1.5 stall", value: 1.5 },
  { label: "1 box", value: 3 },
];

function emptyHorse() {
  return { name: "", stallSpace: 1.5 };
}

// Used for both "+ New Order" (initial is null) and editing an existing
// app order (initial is that order's current fields) -- same form either
// way, just a different onSave target (createAppOrder vs updateAppOrder)
// wired up by the caller.
export default function OrderForm({
  initial,
  defaultTakenBy,
  origins,
  destinations,
  onSave,
  onClose,
  saving,
  errorText,
  extraActions,
}) {
  const [requestedBy, setRequestedBy] = useState(initial?.requestedBy || "");
  const [takenBy, setTakenBy] = useState(initial?.takenBy || defaultTakenBy || "");
  const [requestedDates, setRequestedDates] = useState(initial?.requestedDates || "");
  const [origin, setOrigin] = useState(initial?.origin || "");
  const [originTrack, setOriginTrack] = useState(initial?.originTrack || "");
  const [originTrainer, setOriginTrainer] = useState(initial?.originTrainer || "");
  const [originFarm, setOriginFarm] = useState(initial?.originFarm || "");
  const [destination, setDestination] = useState(initial?.destination || "");
  const [track, setTrack] = useState(initial?.track || "");
  const [destinationTrainer, setDestinationTrainer] = useState(initial?.destinationTrainer || "");
  const [destinationFarm, setDestinationFarm] = useState(initial?.destinationFarm || "");
  const [horses, setHorses] = useState(
    initial?.horses?.length ? initial.horses.map((h) => ({ name: h.name || "", stallSpace: h.stallSpace ?? 1.5 })) : [emptyHorse()]
  );
  const [notes, setNotes] = useState(initial?.notes || "");

  function updateHorse(i, patch) {
    setHorses((prev) => prev.map((h, idx) => (idx === i ? { ...h, ...patch } : h)));
  }
  function addHorse() {
    setHorses((prev) => [...prev, emptyHorse()]);
  }
  function removeHorse(i) {
    setHorses((prev) => (prev.length > 1 ? prev.filter((_, idx) => idx !== i) : prev));
  }

  const canSave = origin.trim() && destination.trim() && horses.some((h) => h.stallSpace !== "" && h.stallSpace != null);

  function handleSave() {
    if (!canSave) return;
    onSave({
      requestedBy: requestedBy.trim(),
      takenBy: takenBy.trim(),
      requestedDates: requestedDates.trim() || "date flexible",
      dateTaken: initial?.dateTaken || new Date().toISOString().slice(0, 10),
      origin: origin.trim().toUpperCase(),
      originTrack: originTrack.trim(),
      originTrainer: originTrainer.trim(),
      originFarm: originFarm.trim(),
      destination: destination.trim().toUpperCase(),
      track: track.trim(),
      destinationTrainer: destinationTrainer.trim(),
      destinationFarm: destinationFarm.trim(),
      horses: horses
        .filter((h) => h.stallSpace !== "" && h.stallSpace != null)
        .map((h) => ({ name: h.name.trim() || null, stallSpace: Number(h.stallSpace) })),
      notes: notes.trim(),
      tripDate: initial?.tripDate || "",
      drivers: initial?.drivers || "",
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
        zIndex: 70,
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%",
          maxWidth: 660,
          maxHeight: "90vh",
          overflowY: "auto",
          background: "#0b1c14",
          border: "1px solid #22402f",
          borderRadius: 16,
          boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
        }}
      >
        <div
          style={{
            padding: "22px 26px",
            borderBottom: "1px solid #1d3527",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
          }}
        >
          <div>
            <div style={{ fontSize: 11, letterSpacing: 3, color: "#7f9c8c", fontFamily: "ui-monospace, monospace" }}>
              {initial ? "EDIT ORDER" : "+ NEW ORDER"}
            </div>
            <div style={{ fontSize: 20, fontWeight: 900, marginTop: 4 }}>
              {initial ? "Edit transport request" : "Add a transport request"}
            </div>
          </div>
          <button
            onClick={onClose}
            style={{
              background: "transparent",
              border: "1px solid #2c4a3a",
              color: "#9fb8a8",
              borderRadius: 8,
              width: 32,
              height: 32,
              fontSize: 18,
              cursor: "pointer",
            }}
          >
            ×
          </button>
        </div>

        <div style={{ padding: "22px 26px", display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
            <div style={{ flex: "1 1 160px" }}>
              <label style={labelStyle}>REQUESTED BY</label>
              <input style={selectStyle} value={requestedBy} onChange={(e) => setRequestedBy(e.target.value)} placeholder="name" />
            </div>
            <div style={{ flex: "1 1 160px" }}>
              <label style={labelStyle}>TAKEN BY</label>
              <input style={selectStyle} value={takenBy} onChange={(e) => setTakenBy(e.target.value)} placeholder="you" />
            </div>
          </div>

          <div>
            <label style={labelStyle}>REQUESTED DATE(S)</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
              {["NEXT AVAILABLE", "ASAP", "READY"].map((chip) => (
                <div key={chip} style={chipStyle(requestedDates === chip)} onClick={() => setRequestedDates(chip)}>
                  {chip}
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <input
                style={{ ...selectStyle, flex: "1 1 220px" }}
                value={requestedDates}
                onChange={(e) => setRequestedDates(e.target.value)}
                placeholder="e.g. mid September, or pick a date →"
              />
              <input
                type="date"
                style={{ ...selectStyle, width: "auto", flex: "0 0 auto" }}
                onChange={(e) => {
                  if (!e.target.value) return;
                  const d = new Date(e.target.value + "T00:00:00");
                  setRequestedDates(
                    d.toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" }).toUpperCase()
                  );
                }}
              />
            </div>
          </div>

          <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
            <div style={{ flex: "1 1 260px", border: "1px solid #1d3527", borderRadius: 10, padding: 12 }}>
              <div style={{ fontSize: 11, letterSpacing: 2, color: "#7be0a8", fontWeight: 800, marginBottom: 8 }}>ORIGIN</div>
              <label style={labelStyle}>REGION</label>
              <input
                style={{ ...selectStyle, marginBottom: 8 }}
                value={origin}
                onChange={(e) => setOrigin(e.target.value)}
                placeholder="e.g. KENTUCKY"
                list="sallee-origins"
              />
              <datalist id="sallee-origins">
                {origins.map((o) => (
                  <option key={o} value={o} />
                ))}
              </datalist>
              <label style={labelStyle}>TRACK (optional)</label>
              <input style={{ ...selectStyle, marginBottom: 8 }} value={originTrack} onChange={(e) => setOriginTrack(e.target.value)} />
              <label style={labelStyle}>TRAINER</label>
              <input
                style={{ ...selectStyle, marginBottom: 8 }}
                value={originTrainer}
                onChange={(e) => setOriginTrainer(e.target.value)}
              />
              <label style={labelStyle}>FARM</label>
              <input style={selectStyle} value={originFarm} onChange={(e) => setOriginFarm(e.target.value)} />
            </div>

            <div style={{ flex: "1 1 260px", border: "1px solid #1d3527", borderRadius: 10, padding: 12 }}>
              <div style={{ fontSize: 11, letterSpacing: 2, color: "#eaf3ec", fontWeight: 800, marginBottom: 8 }}>DESTINATION</div>
              <label style={labelStyle}>REGION</label>
              <input
                style={{ ...selectStyle, marginBottom: 8 }}
                value={destination}
                onChange={(e) => setDestination(e.target.value)}
                placeholder="e.g. OCALA"
                list="sallee-destinations"
              />
              <datalist id="sallee-destinations">
                {destinations.map((d) => (
                  <option key={d} value={d} />
                ))}
              </datalist>
              <label style={labelStyle}>TRACK (optional)</label>
              <input style={{ ...selectStyle, marginBottom: 8 }} value={track} onChange={(e) => setTrack(e.target.value)} />
              <label style={labelStyle}>TRAINER</label>
              <input
                style={{ ...selectStyle, marginBottom: 8 }}
                value={destinationTrainer}
                onChange={(e) => setDestinationTrainer(e.target.value)}
              />
              <label style={labelStyle}>FARM</label>
              <input style={selectStyle} value={destinationFarm} onChange={(e) => setDestinationFarm(e.target.value)} />
            </div>
          </div>

          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <label style={labelStyle}>HORSES</label>
              <button
                onClick={addHorse}
                style={{
                  background: "transparent",
                  border: "1px solid #2c4a3a",
                  color: "#7be0a8",
                  borderRadius: 8,
                  padding: "4px 10px",
                  fontSize: 11,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                + add horse
              </button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {horses.map((h, i) => (
                <div key={i} style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <input
                    style={{ ...selectStyle, flex: "2 1 160px" }}
                    value={h.name}
                    onChange={(e) => updateHorse(i, { name: e.target.value })}
                    placeholder="horse name (optional)"
                  />
                  <div style={{ display: "flex", gap: 4 }}>
                    {STALL_PRESETS.map((p) => (
                      <div key={p.value} style={chipStyle(h.stallSpace === p.value)} onClick={() => updateHorse(i, { stallSpace: p.value })}>
                        {p.label}
                      </div>
                    ))}
                  </div>
                  <input
                    type="number"
                    step="0.5"
                    min="0"
                    style={{ ...selectStyle, flex: "0 0 70px" }}
                    value={h.stallSpace}
                    onChange={(e) => updateHorse(i, { stallSpace: e.target.value === "" ? "" : Number(e.target.value) })}
                  />
                  {horses.length > 1 && (
                    <button
                      onClick={() => removeHorse(i)}
                      style={{
                        background: "transparent",
                        border: "1px solid #4a2c2c",
                        color: "#ff9d9d",
                        borderRadius: 6,
                        width: 28,
                        height: 28,
                        cursor: "pointer",
                      }}
                    >
                      ×
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div>
            <label style={labelStyle}>NOTES (optional)</label>
            <textarea
              style={{ ...selectStyle, minHeight: 60, resize: "vertical" }}
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
            />
          </div>

          {errorText && (
            <div
              style={{
                padding: "10px 14px",
                background: "rgba(255,107,107,0.1)",
                border: "1px solid #ff6b6b",
                borderRadius: 8,
                color: "#ff9d9d",
                fontSize: 12,
              }}
            >
              {errorText}
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginTop: 4, flexWrap: "wrap" }}>
            <div style={{ display: "flex", gap: 8 }}>{extraActions}</div>
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
              disabled={!canSave || saving}
              style={{
                background: !canSave || saving ? "#2c4a3a" : "#f0c95a",
                color: !canSave || saving ? "#6f8c7c" : "#0a2417",
                border: "none",
                borderRadius: 8,
                padding: "10px 20px",
                fontWeight: 800,
                fontSize: 13,
                cursor: !canSave || saving ? "not-allowed" : "pointer",
              }}
            >
              {saving ? "Saving…" : initial ? "Save changes" : "Add order to board"}
            </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
