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
            {/* tail */}
            <path
              d="M -24 -4 Q -37 0 -34 10 Q -39 17 -32 23"
              fill="none"
              stroke="#2a160a"
              strokeWidth="3"
              strokeLinecap="round"
            />
            {/* hind legs */}
            <rect x="-16" y="7" width="4" height="17" rx="1" fill="#2a160a" />
            <rect x="-9" y="9" width="4" height="15" rx="1" fill="#2a160a" />
            {/* body */}
            <ellipse cx="-2" cy="0" rx="22" ry="9" fill="#6b4226" stroke="#2a160a" strokeWidth="1.5" />
            {/* front legs */}
            <rect x="8" y="9" width="4" height="15" rx="1" fill="#2a160a" />
            <rect x="15" y="7" width="4" height="17" rx="1" fill="#2a160a" />
            {/* arched neck + long head/muzzle */}
            <path
              d="M 12 -6
                 C 15 -17, 21 -26, 29 -30
                 C 33 -32, 37 -31, 39 -28
                 L 47 -25
                 C 51 -23.5, 51 -19.5, 47 -18.5
                 L 41 -17.5
                 C 41.5 -15, 40.5 -12, 38 -10.5
                 L 35 -2
                 C 34.3 1, 30.5 1, 30 -2
                 L 29 -8
                 C 24 -9.5, 16 -9, 12 -6
                 Z"
              fill="#6b4226"
              stroke="#2a160a"
              strokeWidth="1.5"
            />
            {/* nostril */}
            <ellipse cx="44" cy="-21.5" rx="1.6" ry="1" fill="#2a160a" />
            {/* ear */}
            <path d="M 31 -29.5 L 29 -37 L 36 -31 Z" fill="#2a160a" />
            {/* mane */}
            <path
              d="M 13 -8 L 17 -14 L 20 -10 L 24 -19 L 27 -15 L 31 -24 L 34 -20"
              fill="none"
              stroke="#2a160a"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
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
