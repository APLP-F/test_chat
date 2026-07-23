import { useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
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
  const [mensaje, setMensaje] = useState("Conectando con Puerto Emplea...");

  useEffect(() => {
    const cargarChat = async () => {
      try {
        await msalInstance.initialize();
        await msalInstance.handleRedirectPromise();

        const cuentas = msalInstance.getAllAccounts();

        if (cuentas.length === 0) {
          await msalInstance.loginRedirect({
            ...copilotLoginRequest,
          });

          return;
        }

        const cuenta = cuentas[0];
        setUsuario(cuenta.username);

        let tokenResult;

        try {
          tokenResult = await msalInstance.acquireTokenSilent({
            ...copilotLoginRequest,
            account: cuenta,
          });
        } catch {
          await msalInstance.acquireTokenRedirect({
            ...copilotLoginRequest,
            account: cuenta,
          });

          return;
        }

        const client = new CopilotStudioClient(
          settings,
          tokenResult.accessToken
        );

        const nuevaConexion = CopilotStudioWebChat.createConnection(client, {
          showTyping: true,
        });

        setConnection(nuevaConexion);
      } catch (error) {
        console.error(error);
        setMensaje("Error cargando el chat");
      }
    };

    cargarChat();
  }, []);

  const styleOptions = {
    rootHeight: "100%",
    rootWidth: "100%",

    primaryFont: "Inter, Segoe UI, sans-serif",

    botAvatarInitials: "",
    userAvatarInitials: "",

    bubbleMaxWidth: 920,
    bubbleMinWidth: 260,
    bubbleBorderRadius: 10,
    bubbleBorderColor: "#e5e7eb",
    bubbleBorderStyle: "solid",
    bubbleBorderWidth: 1,
    bubbleBackground: "#ffffff",
    bubbleTextColor: "#111827",
    bubbleNubSize: 0,

    bubbleFromUserMaxWidth: 720,
    bubbleFromUserBorderRadius: 10,
    bubbleFromUserBorderColor: "#e5e7eb",
    bubbleFromUserBorderStyle: "solid",
    bubbleFromUserBorderWidth: 1,
    bubbleFromUserBackground: "#ffffff",
    bubbleFromUserTextColor: "#111827",
    bubbleFromUserNubSize: 0,

    sendBoxHeight: 58,
    sendBoxBackground: "#ffffff",
    sendBoxBorderTop: "1px solid #e5e7eb",
    sendBoxTextColor: "#111827",
    sendBoxPlaceholderColor: "#777777",

    hideUploadButton: false,
  };

  if (!connection) {
    return (
      <div className="loading">
        <div className="loading-card">
          <h1>Puerto Emplea</h1>
          <p>{mensaje}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="app-layout">
      <aside className="sidebar">
        <div className="brand">
          <img
            src="/logo.png"
            alt="Puerto Emplea"
            className="logo"
          />

          <div className="brand-title">
            Puerto Emplea
          </div>

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