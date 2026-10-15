import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import SalleeLoadBoard from "../SalleeLoadBoard.jsx";
import Intro, { shouldPlayIntro } from "./Intro.jsx";
import { DEMO_MODE, initAuth, signIn, signOut } from "./auth.js";

const screenWrap = {
  minHeight: "100vh",
  display: "flex",
  flexDirection: "column",
  alignItems: "center",
  justifyContent: "center",
  gap: 20,
  background: "radial-gradient(ellipse at top, #0f2a1c 0%, #05100b 60%)",
  color: "#eaf3ec",
  fontFamily: "Arial, Helvetica, sans-serif",
  textAlign: "center",
  padding: 24,
};

function greeting() {
  const hour = new Date().getHours();
  if (hour < 12) return "morning";
  if (hour < 18) return "afternoon";
  return "evening";
}

function Root() {
  const [introDone, setIntroDone] = useState(!shouldPlayIntro());
  const [authReady, setAuthReady] = useState(DEMO_MODE);
  const [account, setAccount] = useState(null);
  const [greeted, setGreeted] = useState(false);

  useEffect(() => {
    if (DEMO_MODE) return;
    let cancelled = false;
    initAuth()
      .then((acc) => {
        if (cancelled) return;
        setAccount(acc || null);
        setAuthReady(true);
      })
      .catch((err) => {
        console.error("Auth init failed", err);
        if (!cancelled) setAuthReady(true);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (account && !greeted) {
      const t = setTimeout(() => setGreeted(true), 1300);
      return () => clearTimeout(t);
    }
  }, [account, greeted]);

  if (!introDone) {
    return <Intro onDone={() => setIntroDone(true)} />;
  }

  if (DEMO_MODE) {
    return <SalleeLoadBoard demoMode account={null} onSignOut={null} />;
  }

  if (!authReady) {
    return (
      <div style={screenWrap}>
        <div style={{ fontFamily: "ui-monospace, monospace", color: "#8fae9c", fontSize: 13, letterSpacing: 2 }}>
          LOADING…
        </div>
      </div>
    );
  }

  if (!account) {
    return (
      <div style={screenWrap}>
        <div style={{ fontSize: 12, letterSpacing: 6, color: "#e8c96a", fontWeight: 800 }}>SALLEE HORSE VANS</div>
        <div style={{ fontSize: 30, fontWeight: 900 }}>LOAD BOARD</div>
        <button
          onClick={() => {
            try {
              signIn();
            } catch (err) {
              alert(err.message);
            }
          }}
          style={{
            background: "#f0c95a",
            color: "#0a2417",
            border: "none",
            borderRadius: 10,
            padding: "14px 24px",
            fontWeight: 800,
            fontSize: 14,
            cursor: "pointer",
          }}
        >
          Sign in with Microsoft
        </button>
      </div>
    );
  }

  if (!greeted) {
    const first = (account.name || "").split(" ")[0] || "there";
    return (
      <div style={screenWrap}>
        <div style={{ fontSize: 12, letterSpacing: 6, color: "#e8c96a", fontWeight: 800 }}>SALLEE HORSE VANS</div>
        <div style={{ fontSize: 28, fontWeight: 900 }}>
          Good {greeting()}, {first}
        </div>
      </div>
    );
  }

  return <SalleeLoadBoard demoMode={false} account={account} onSignOut={signOut} />;
}

const root = createRoot(document.getElementById("root"));
root.render(<Root />);
