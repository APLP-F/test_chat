import { useEffect, useMemo, useRef, useState } from "react";
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

const STORAGE_KEY = "puerto_emplea_chat_conversations_v1";

type ChatRole = "user" | "bot";

interface SavedMessage {
  id: string;
  role: ChatRole;
  text: string;
  createdAt: string;
}

interface SavedConversation {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  messages: SavedMessage[];
}

function createId(): string {
  if (window.crypto && "randomUUID" in window.crypto) {
    return window.crypto.randomUUID();
  }

  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createEmptyConversation(): SavedConversation {
  const now = new Date().toISOString();

  return {
    id: createId(),
    title: "Conversación actual",
    createdAt: now,
    updatedAt: now,
    messages: [],
  };
}

function loadStoredConversations(): SavedConversation[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);

    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw) as unknown;

    if (!Array.isArray(parsed)) {
      return [];
    }

    return [...parsed] as SavedConversation[];
  } catch {
    return [];
  }
}

function saveStoredConversations(conversations: SavedConversation[]): void {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(conversations));
}

function getInitialChatData(): {
  conversations: SavedConversation[];
  activeConversationId: string;
} {
  const stored = loadStoredConversations();

  if (stored.length > 0) {
    return {
      conversations: stored,
      activeConversationId: stored[0].id,
    };
  }

  const firstConversation = createEmptyConversation();

  saveStoredConversations([firstConversation]);

  return {
    conversations: [firstConversation],
    activeConversationId: firstConversation.id,
  };
}

const initialChatData = getInitialChatData();

function Chat() {
  const [connection, setConnection] = useState<any>(null);
  const [usuario, setUsuario] = useState("");
  const [mensaje, setMensaje] = useState("Preparando Puerto Emplea...");
  const [necesitaLogin, setNecesitaLogin] = useState(false);
  const [cargando, setCargando] = useState(false);
  const [msalListo, setMsalListo] = useState(false);

  const [accessTokenActual, setAccessTokenActual] = useState("");

  const [conversations, setConversations] = useState<SavedConversation[]>(
    initialChatData.conversations
  );

  const [activeConversationId, setActiveConversationId] = useState(
    initialChatData.activeConversationId
  );

  const activeConversationIdRef = useRef(initialChatData.activeConversationId);
  const liveConversationIdRef = useRef(initialChatData.activeConversationId);

  const [webChatKey, setWebChatKey] = useState(createId());
  const [modoHistorial, setModoHistorial] = useState(false);

  const estaEnIframe = window.self !== window.top;

  function appendMessageToConversation(
    conversationId: string,
    role: ChatRole,
    text: string
  ): void {
    const cleanText = text?.trim();

    if (!cleanText) {
      return;
    }

    const now = new Date().toISOString();

    const newMessage: SavedMessage = {
      id: createId(),
      role,
      text: cleanText,
      createdAt: now,
    };

    setConversations((previous) => {
      const next = previous.map((conversation) => {
        if (conversation.id !== conversationId) {
          return conversation;
        }

        const isFirstUserMessage =
          role === "user" && conversation.messages.length === 0;

        return {
          ...conversation,
          title: isFirstUserMessage
            ? cleanText.length > 42
              ? `${cleanText.slice(0, 42)}...`
              : cleanText
            : conversation.title,
          updatedAt: now,
          messages: [...conversation.messages, newMessage],
        };
      });

      saveStoredConversations(next);

      return next;
    });
  }

  function createWebChatStore() {
    return createStore({}, () => (next: any) => (action: any) => {
      const currentLiveConversationId = liveConversationIdRef.current;

      if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
        const activity = action.payload?.activity;

        if (
          activity?.from?.role === "bot" &&
          activity?.type === "message" &&
          activity?.text
        ) {
          appendMessageToConversation(
            currentLiveConversationId,
            "bot",
            activity.text
          );

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
        const userText = action.payload?.text || "";

        appendMessageToConversation(
          currentLiveConversationId,
          "user",
          userText
        );

        window.parent.postMessage(
          {
            type: "USER_MESSAGE",
            text: userText,
          },
          "*"
        );
      }

      return next(action);
    });
  }

  const [store, setStore] = useState<any>(() => createWebChatStore());

  const activeConversation = useMemo(() => {
    return (
      conversations.find(
        (conversation) => conversation.id === activeConversationId
      ) || conversations[0]
    );
  }, [conversations, activeConversationId]);

  const conectarConToken = async (
    username: string,
    accessToken: string
  ): Promise<void> => {
    setCargando(true);

    try {
      setUsuario(username);
      setAccessTokenActual(accessToken);
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

  const crearNuevaConversacion = async (): Promise<void> => {
    if (!accessTokenActual || !usuario) {
      setMensaje("Inicia sesión para crear una nueva conversación.");
      setNecesitaLogin(true);
      return;
    }

    const newConversation = createEmptyConversation();

    setConversations((previous) => {
      const next = [newConversation, ...previous];

      saveStoredConversations(next);

      return next;
    });

    setActiveConversationId(newConversation.id);

    activeConversationIdRef.current = newConversation.id;
    liveConversationIdRef.current = newConversation.id;

    setModoHistorial(false);

    const newStore = createWebChatStore();

    setStore(newStore);
    setWebChatKey(createId());
    setConnection(null);
    setMensaje("Creando nueva conversación...");

    try {
      const client = new CopilotStudioClient(settings, accessTokenActual);

      const nuevaConexion = await CopilotStudioWebChat.createConnection(
        client,
        {
          showTyping: true,
        }
      );

      setConnection(nuevaConexion);
      setNecesitaLogin(false);
      setMensaje("");
    } catch (error) {
      console.error("Error creando nueva conversación:", error);
      setMensaje("No se pudo crear una nueva conversación.");
      setNecesitaLogin(false);
    }
  };

  const seleccionarConversacion = (conversationId: string): void => {
    setActiveConversationId(conversationId);
    activeConversationIdRef.current = conversationId;

    if (conversationId === liveConversationIdRef.current) {
      setModoHistorial(false);
    } else {
      setModoHistorial(true);
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

  if (!connection && !modoHistorial) {
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

        <button className="new-chat-btn" onClick={crearNuevaConversacion}>
          + Nueva conversación
        </button>

        <div className="history-title">Conversaciones</div>

        <div
          style={{
            backgroundColor: "red",
            color: "white",
            padding: "12px",
            fontWeight: "bold",
            textAlign: "center",
            marginBottom: "10px",
            borderRadius: "8px"
          }}
        >
          BOTON DE PRUEBA
        </div>

        <div className="conversation-list">
          {conversations.map((conversation) => (
            <button
              key={conversation.id}
              className={
                conversation.id === activeConversation?.id
                  ? "history-item active"
                  : "history-item"
              }
              onClick={() => seleccionarConversacion(conversation.id)}
              title={conversation.title}
            >
              {conversation.title}
            </button>
          ))}
        </div>

        <div className="user-info">{usuario}</div>
      </aside>

      <main className="chat-area">
        <div className="chat-wrapper">
          {modoHistorial ? (
            <div className="saved-history">
              <div className="saved-history-header">
                <strong>{activeConversation?.title || "Conversación"}</strong>
                <span>Historial guardado</span>
              </div>

              <div className="saved-history-messages">
                {activeConversation?.messages.length ? (
                  activeConversation.messages.map((message) => (
                    <div
                      key={message.id}
                      className={
                        message.role === "user"
                          ? "saved-message saved-message-user"
                          : "saved-message saved-message-bot"
                      }
                    >
                      <div className="saved-message-role">
                        {message.role === "user" ? "Tú" : "Puerto Emplea"}
                      </div>

                      <div className="saved-message-text">{message.text}</div>
                    </div>
                  ))
                ) : (
                  <div className="saved-history-empty">
                    Esta conversación todavía no tiene mensajes guardados.
                  </div>
                )}
              </div>

              <button
                className="new-chat-btn"
                onClick={crearNuevaConversacion}
                style={{ marginTop: "16px" }}
              >
                + Nueva conversación
              </button>
            </div>
          ) : (
            <Composer
              key={webChatKey}
              directLine={connection}
              store={store}
              styleOptions={styleOptions}
            >
              <BasicWebChat />
            </Composer>
          )}
        </div>
      </main>
    </div>
  );
}

export default Chat;