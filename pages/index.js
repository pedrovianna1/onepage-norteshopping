import { useState, useEffect } from "react";
import dynamic from "next/dynamic";

// O componente usa window.storage/localStorage — precisa carregar só no browser
const OnePageDashboard = dynamic(() => import("../components/OnePageDashboard"), { ssr: false });

const SENHA = "norteshopping2026"; // troque aqui antes de publicar

export default function Home() {
  const [ok, setOk] = useState(false);
  const [input, setInput] = useState("");
  const [checked, setChecked] = useState(false);

  useEffect(() => {
    if (typeof window !== "undefined") {
      setOk(window.sessionStorage.getItem("onepage_auth") === "1");
    }
    setChecked(true);
  }, []);

  const entrar = () => {
    if (input === SENHA) {
      window.sessionStorage.setItem("onepage_auth", "1");
      setOk(true);
    } else {
      alert("Senha incorreta.");
    }
  };

  if (!checked) return null;

  if (!ok) {
    return (
      <div style={{
        minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
        background: "#0D454A", fontFamily: "'Segoe UI', Arial, sans-serif",
      }}>
        <div style={{ background: "#12474D", padding: 32, borderRadius: 14, width: 300, textAlign: "center" }}>
          <div style={{ color: "#fff", fontWeight: 700, marginBottom: 14 }}>One Page NorteShopping</div>
          <input
            type="password"
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => e.key === "Enter" && entrar()}
            placeholder="Senha de acesso"
            style={{ width: "100%", padding: 10, borderRadius: 8, border: "none", marginBottom: 12, boxSizing: "border-box" }}
          />
          <button onClick={entrar} style={{
            width: "100%", padding: 10, borderRadius: 8, border: "none",
            background: "#EDA754", fontWeight: 700, cursor: "pointer",
          }}>Entrar</button>
        </div>
      </div>
    );
  }

  return <OnePageDashboard />;
}
