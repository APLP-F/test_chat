import { useEffect, useMemo, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import type { AccountInfo } from "@azure/msal-browser";
import { Components, createStore } from "botframework-webchat";
import {
  CopilotStudioClient,
  CopilotStudioWebChat,
} from "@microsoft/agents-copilotstudio-client";

import {
  msalConfig,
  copilotLoginRequest,
  appRedirectUri,
  popupRedirectUri,
} from "./authConfig";

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
  const [msalListo, setMsalListo] = useState(false);

  const estaEnIframe = window.self !== window.top;

  const store = useMemo(
    () =>
      createStore({}, () => (next: any) => (action: any) => {
        if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
          const activity = action.payload?.activity;

          if (
            activity?.from?.role === "bot" &&
            activity?.type === "message" &&
            activity?.text
          ) {
            window.parent.postMessage(
              {
                type: "BOT_RESPONSE",
                text: activity.text,
                conversationId: activity.conversation?.id || "",
              },
              "*"
            );
          }
        }

        if (action.type === "WEB_CHAT/SEND_MESSAGE") {
          window.parent.postMessage(
            {
              type: "USER_MESSAGE",
              text: action.payload?.text || "",
            },
            "*"
          );
        }

        return next(action);
      }),
    []
  );

  const conectarConToken = async (
    username: string,
    accessToken: string
  ): Promise<void> => {
    setCargando(true);

    try {
      setUsuario(username);
      setMensaje("Conectando con Puerto Emplea...");

      const client = new CopilotStudioClient(settings, accessToken);

      const nuevaConexion = await CopilotStudioWebChat.createConnection(
        client,
        {
          showTyping: true,
        }
      );

      setConnection(nuevaConexion);
      setNecesitaLogin(false);
    } catch (error) {
      console.error("Error conectando con Copilot Studio:", error);
      setMensaje("No se pudo conectar con Puerto Emplea.");
      setNecesitaLogin(true);
    } finally {
      setCargando(false);
    }
  };

  const conectarConCuenta = async (cuenta: AccountInfo): Promise<void> => {
    setCargando(true);

    try {
      setUsuario(cuenta.username);
      setMensaje("Obteniendo permisos para Copilot Studio...");

      const tokenResult = await msalInstance.acquireTokenSilent({
        ...copilotLoginRequest,
        account: cuenta,
      });

      await conectarConToken(cuenta.username, tokenResult.accessToken);
    } catch (error) {
      console.warn("No se pudo obtener token silenciosamente:", error);

      setMensaje("Inicia sesión para usar Puerto Emplea.");
      setNecesitaLogin(true);
    } finally {
      setCargando(false);
    }
  };

  const iniciarSesion = async (): Promise<void> => {
    if (!msalListo || cargando) {
      return;
    }

    setCargando(true);
    setMensaje("Abriendo inicio de sesión...");

    try {
      if (estaEnIframe) {
        const authUrl = `${popupRedirectUri}?login=1`;

        const popup = window.open(
          authUrl,
          "PuertoEmpleaAuth",
          "width=520,height=720,menubar=no,toolbar=no,location=yes,status=no,resizable=yes,scrollbars=yes"
        );

        if (!popup) {
          setMensaje("El navegador bloqueó la ventana de inicio de sesión.");
          setNecesitaLogin(true);
          setCargando(false);
        }

        return;
      }

      await msalInstance.loginRedirect({
        ...copilotLoginRequest,
        redirectUri: appRedirectUri,
      });
    } catch (error) {
      console.error("Error iniciando sesión:", error);
      setMensaje("No se pudo iniciar sesión.");
      setNecesitaLogin(true);
      setCargando(false);
    }
  };

  useEffect(() => {
    const recibirMensajeAuth = async (event: MessageEvent) => {
      if (event.origin !== window.location.origin) {
        return;
      }

      const data = event.data;

      if (!data || typeof data !== "object") {
        return;
      }

      if (data.type === "PUERTO_EMPLEA_AUTH_SUCCESS") {
        if (!data.accessToken) {
          setMensaje("No se recibió token de acceso.");
          setNecesitaLogin(true);
          setCargando(false);
          return;
        }

        await conectarConToken(
          data.username || "Usuario autenticado",
          data.accessToken
        );

        return;
      }

      if (data.type === "PUERTO_EMPLEA_AUTH_ERROR") {
        console.error("Error desde auth.html:", data.error);
        setMensaje("No se pudo iniciar sesión.");
        setNecesitaLogin(true);
        setCargando(false);
      }
    };

    window.addEventListener("message", recibirMensajeAuth);

    return () => {
      window.removeEventListener("message", recibirMensajeAuth);
    };
  }, []);

  useEffect(() => {
    let cancelado = false;

    const cargarChat = async (): Promise<void> => {
      try {
        await msalInstance.initialize();

        if (cancelado) {
          return;
        }

        if (!estaEnIframe) {
          const redirectResult = await msalInstance.handleRedirectPromise();

          if (redirectResult?.account && redirectResult.accessToken) {
            await conectarConToken(
              redirectResult.account.username,
              redirectResult.accessToken
            );
            setMsalListo(true);
            return;
          }

          if (redirectResult?.account) {
            await conectarConCuenta(redirectResult.account);
            setMsalListo(true);
            return;
          }
        }

        setMsalListo(true);

        const cuentas = msalInstance.getAllAccounts();

        if (cuentas.length === 0) {
          setMensaje("Inicia sesión para usar Puerto Emplea.");
          setNecesitaLogin(true);
          return;
        }

        await conectarConCuenta(cuentas[0]);
      } catch (error) {
        console.error("Error inicializando MSAL:", error);

        if (!cancelado) {
          setMensaje("Inicia sesión para usar Puerto Emplea.");
          setNecesitaLogin(true);
          setMsalListo(true);
        }
      }
    };

    cargarChat();

    return () => {
      cancelado = true;
    };
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
              disabled={cargando || !msalListo}
              style={{
                marginTop: "20px",
                padding: "12px 22px",
                borderRadius: "10px",
                border: "none",
                background: "#0d3b66",
                color: "#ffffff",
                fontWeight: 600,
                cursor: cargando || !msalListo ? "not-allowed" : "pointer",
              }}
            >
              {cargando ? "Conectando..." : "Iniciar sesión con Microsoft"}
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

          <div className="brand-subtitle">Asistente IA</div>
        </div>

        <button className="new-chat-btn">+ Nueva conversación</button>

        <div className="history-title">Conversaciones</div>

        <div className="history-item active">Conversación actual</div>

        <div className="user-info">{usuario}</div>
      </aside>

      <main className="chat-area">
        <div className="chat-wrapper">
          <Composer
            directLine={connection}
            store={store}
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