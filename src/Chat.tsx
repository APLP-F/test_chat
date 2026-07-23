import { useEffect, useState } from "react";
import { PublicClientApplication, AccountInfo } from "@azure/msal-browser";
import { Components } from "botframework-webchat";
import {
  CopilotStudioClient,
  CopilotStudioWebChat,
} from "@microsoft/agents-copilotstudio-client";

import { msalConfig, copilotLoginRequest } from "./authConfig";
import { settings } from "./settings";
import "./Chat.css";

const { BasicWebChat, Composer } = Components;

const msalInstance = new PublicClientApplication(msalConfig);

function Chat() {
  const [connection, setConnection] = useState<any>(null);
  const [usuario, setUsuario] = useState("");
  const [mensaje, setMensaje] = useState("Preparando Puerto Emplea...");
  const [necesitaLogin, setNecesitaLogin] = useState(false);
  const [cargando, setCargando] = useState(false);

  const conectarConCuenta = async (cuenta: AccountInfo) => {
    setCargando(true);

    try {
      setUsuario(cuenta.username);
      setMensaje("Obteniendo permisos para Copilot Studio...");

      let tokenResult;

      try {
        tokenResult = await msalInstance.acquireTokenSilent({
          ...copilotLoginRequest,
          account: cuenta,
        });
      } catch {
        tokenResult = await msalInstance.acquireTokenPopup({
          ...copilotLoginRequest,
          account: cuenta,
        });
      }

      setMensaje("Conectando con Puerto Emplea...");

      const client = new CopilotStudioClient(
        settings,
        tokenResult.accessToken
      );

      const nuevaConexion =
        CopilotStudioWebChat.createConnection(client, {
          showTyping: true,
        });

      setConnection(nuevaConexion);
      setNecesitaLogin(false);
    } catch (error) {
      console.error("Error conectando con el chat:", error);
      setMensaje("Error cargando el chat");
      setNecesitaLogin(true);
    } finally {
      setCargando(false);
    }
  };

  const iniciarSesion = async () => {
    setCargando(true);

    try {
      await msalInstance.initialize();

      const loginResult = await msalInstance.loginPopup({
        ...copilotLoginRequest,
      });

      if (loginResult.account) {
        await conectarConCuenta(loginResult.account);
      } else {
        setMensaje("No se pudo obtener la cuenta del usuario.");
        setNecesitaLogin(true);
      }
    } catch (error) {
      console.error("Error iniciando sesión:", error);
      setMensaje("No se pudo iniciar sesión.");
      setNecesitaLogin(true);
    } finally {
      setCargando(false);
    }
  };

  useEffect(() => {
    const cargarChat = async () => {
      try {
        await msalInstance.initialize();

        const cuentas = msalInstance.getAllAccounts();

        if (cuentas.length === 0) {
          setMensaje("Inicia sesión para usar Puerto Emplea.");
          setNecesitaLogin(true);
          return;
        }

        await conectarConCuenta(cuentas[0]);
      } catch (error) {
        console.error("Error inicializando MSAL:", error);
        setMensaje("Inicia sesión para usar Puerto Emplea.");
        setNecesitaLogin(true);
      }
    };

    cargarChat();
  }, []);

  const styleOptions = {
    rootHeight: "100%",
    rootWidth: "100%",
    primaryFont: "Inter, Segoe UI, sans-serif",
    hideUploadButton: false,
  };

  if (!connection) {
    return (
      <div className="loading">
        <div className="loading-card">
          <h1>Puerto Emplea</h1>

          <p>{mensaje}</p>

          {usuario && (
            <p>
              Usuario conectado: <strong>{usuario}</strong>
            </p>
          )}

          {necesitaLogin && (
            <button
              onClick={iniciarSesion}
              disabled={cargando}
              style={{
                marginTop: "20px",
                padding: "12px 22px",
                borderRadius: "10px",
                border: "none",
                background: "#0d3b66",
                color: "#ffffff",
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {cargando
                ? "Conectando..."
                : "Iniciar sesión con Microsoft"}
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="app-layout">
      <aside className="sidebar">
        <div className="brand">
          <img
            src={`${import.meta.env.BASE_URL}logo.png`}
            alt="Puerto Emplea"
            className="logo"
          />

          <div className="brand-subtitle">
            Asistente IA
          </div>
        </div>

        <button className="new-chat-btn">
          + Nueva conversación
        </button>

        <div className="history-title">
          Conversaciones
        </div>

        <div className="history-item active">
          Conversación actual
        </div>

        <div className="user-info">
          {usuario}
        </div>
      </aside>

      <main className="chat-area">
        <div className="chat-wrapper">
          <Composer
            directLine={connection}
            styleOptions={styleOptions}
          >
            <BasicWebChat />
          </Composer>
        </div>
      </main>
    </div>
  );
}

export default Chat;