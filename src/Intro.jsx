import React, { useEffect, useRef, useState } from "react";

// Plays once per browser tab session (not once ever — a fresh tab plays it
// again), matching the shop board / dispatch board intros.
const SESSION_KEY = "sallee-load-board:intro-played";
const TOTAL_MS = 4200;

export function shouldPlayIntro() {
  try {
    return !sessionStorage.getItem(SESSION_KEY);
  } catch {
    return true;
  }
}

function markIntroPlayed() {
  try {
    sessionStorage.setItem(SESSION_KEY, "1");
  } catch {
    // private browsing etc -- worst case the intro replays, harmless
  }
}

// Ramp drops -> horse walks up in a green/gold sheet -> ramp shuts -> van
// drives off. Pure CSS keyframes (timed via animation-delay) driving a
// small inline SVG, so there's no animation library dependency. Tap
// anywhere, or "Skip →", jumps straight to onDone.
export default function Intro({ onDone }) {
  const firedRef = useRef(false);
  const [exiting, setExiting] = useState(false);

  useEffect(() => {
    const t = setTimeout(finish, TOTAL_MS);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function finish() {
    if (firedRef.current) return;
    firedRef.current = true;
    markIntroPlayed();
    setExiting(true);
    setTimeout(onDone, 220);
  }

  return (
    <div
      onClick={finish}
      style={{
        position: "fixed",
        inset: 0,
        background: "radial-gradient(ellipse at center, #0f2a1c 0%, #05100b 75%)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 999,
        cursor: "pointer",
        overflow: "hidden",
        opacity: exiting ? 0 : 1,
        transition: "opacity 0.2s ease",
      }}
    >
      <button
        onClick={(e) => {
          e.stopPropagation();
          finish();
        }}
        style={{
          position: "absolute",
          top: 20,
          right: 24,
          background: "transparent",
          border: "1px solid #2c4a3a",
          color: "#9fb8a8",
          borderRadius: 8,
          padding: "8px 14px",
          fontSize: 13,
          fontWeight: 700,
          cursor: "pointer",
          zIndex: 2,
        }}
      >
        Skip →
      </button>

      <div style={{ textAlign: "center" }}>
        <svg width="340" height="190" viewBox="0 0 340 190" style={{ overflow: "visible" }}>
          <g className="sallee-intro-van">
            <rect x="70" y="40" width="190" height="90" rx="10" fill="#123a26" stroke="#2c4a3a" strokeWidth="2" />
            <rect x="248" y="58" width="42" height="58" rx="6" fill="#0c2217" stroke="#2c4a3a" strokeWidth="2" />
            <circle cx="110" cy="140" r="14" fill="#06110b" stroke="#2c4a3a" strokeWidth="3" />
            <circle cx="230" cy="140" r="14" fill="#06110b" stroke="#2c4a3a" strokeWidth="3" />
            <text
              x="165"
              y="90"
              textAnchor="middle"
              fontFamily="ui-monospace, monospace"
              fontSize="13"
              fontWeight="800"
              fill="#f0c95a"
              letterSpacing="1.5"
            >
              SALLEE
            </text>
            <rect
              className="sallee-intro-ramp"
              x="4"
              y="126"
              width="66"
              height="7"
              rx="2"
              fill="#2c4a3a"
              style={{ transformOrigin: "70px 130px" }}
            />
          </g>

          <g className="sallee-intro-horse">
            <ellipse cx="0" cy="4" rx="24" ry="12" fill="#f0c95a" stroke="#0a2417" strokeWidth="1.5" />
            <path d="M 20 -2 Q 32 -7 34 3 Q 32 9 22 7 Z" fill="#3a2a1a" />
            <rect x="-18" y="12" width="5" height="15" fill="#3a2a1a" />
            <rect x="-5" y="12" width="5" height="15" fill="#3a2a1a" />
            <rect x="7" y="12" width="5" height="15" fill="#3a2a1a" />
            <rect x="18" y="12" width="5" height="15" fill="#3a2a1a" />
          </g>
        </svg>

        <div
          style={{
            marginTop: 14,
            fontFamily: "ui-monospace, monospace",
            fontSize: 13,
            letterSpacing: 3,
            color: "#7f9c8c",
            fontWeight: 700,
          }}
        >
          LOADING…
        </div>
      </div>

      <style>{`
        .sallee-intro-van {
          animation: sallee-intro-drive 1.0s ease-in 3.1s forwards;
        }
        .sallee-intro-ramp {
          animation:
            sallee-intro-ramp-down 0.6s ease-out 0.2s forwards,
            sallee-intro-ramp-up 0.6s ease-in 2.4s forwards;
        }
        .sallee-intro-horse {
          opacity: 0;
          animation: sallee-intro-horse-walk 1.6s ease-in-out 0.9s forwards;
        }
        @keyframes sallee-intro-ramp-down {
          from { transform: rotate(0deg); }
          to   { transform: rotate(52deg); }
        }
        @keyframes sallee-intro-ramp-up {
          from { transform: rotate(52deg); }
          to   { transform: rotate(0deg); }
        }
        @keyframes sallee-intro-horse-walk {
          0%   { opacity: 0;   transform: translate(6px, 148px) scale(0.92); }
          12%  { opacity: 1; }
          78%  { opacity: 1;   transform: translate(150px, 72px) scale(0.92); }
          100% { opacity: 0;   transform: translate(172px, 66px) scale(0.86); }
        }
        @keyframes sallee-intro-drive {
          from { transform: translateX(0); }
          to   { transform: translateX(460px); }
        }
      `}</style>
    </div>
  );
}
