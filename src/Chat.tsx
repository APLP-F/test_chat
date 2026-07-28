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

const DEFAULT_STORAGE_KEY = "puerto_emplea_chat_conversations_v1";
let activeStorageKey = DEFAULT_STORAGE_KEY;

const DATAVERSE_FLOW_URL = "https://6b8fc4584a99e825afb8ecbd16a97c.53.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/18/workflows/78ff7b170f4d4d3dbe0c175ba4a5568e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=66ilhkEvJXjx1ZeSaBfNCgmGZDvmVqsFr_jK8Nh1hME";
const DEFAULT_CONVERSATION_TITLE = "Conversación actual";

function createUserStorageKey(username: string): string {
  const safeUsername = username
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._@-]/g, "_");

  return `puerto_emplea_chat_conversations_v1_${safeUsername || "anonimo"}`;
}

type ChatRole = "user" | "bot";

interface DataverseConversation {
  conversationRowId?: string;
  localConversationId?: string;
  titulo?: string;
  usuarioEmail?: string;
  copilotConversationId?: string;
  fechaCreacion?: string;
}

interface DataverseMessage {
  rol?: string;
  mensaje?: string;
  fecha?: string;
}

interface DataverseFlowResponse {
  ok?: boolean;
  conversationRowId?: string;
  conversaciones?: DataverseConversation[];
  mensajes?: DataverseMessage[];
  mensaje?: string;
  error?: string;
}

interface SavedMessage {
  id: string;
  role: ChatRole;
  text: string;
  createdAt: string;
}

interface SavedConversation {
  id: string;
  dataverseConversationRowId?: string;
  copilotConversationId?: string;
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

function createConversationTitle(text: string): string {
  const cleanText = text.trim().replace(/\s+/g, " ");

  if (!cleanText) {
    return DEFAULT_CONVERSATION_TITLE;
  }

  if (cleanText.length > 42) {
    return `${cleanText.slice(0, 42)}...`;
  }

  return cleanText;
}

function cleanConversationTitle(title: string): string {
  const cleanTitle = title
    .replace(/^(Continuación:\s*)+/i, "")
    .trim();

  return cleanTitle || DEFAULT_CONVERSATION_TITLE;
}

function getTitleFromFirstUserMessage(messages: SavedMessage[]): string {
  const firstUserMessage = messages.find(
    (message) => message.role === "user" && message.text?.trim()
  );

  if (!firstUserMessage) {
    return DEFAULT_CONVERSATION_TITLE;
  }

  return createConversationTitle(firstUserMessage.text);
}

function normalizeConversation(
  conversation: Partial<SavedConversation>
): SavedConversation {
  const now = new Date().toISOString();

  const safeMessages = Array.isArray(conversation.messages)
    ? conversation.messages
    : [];

  const currentTitle = cleanConversationTitle(
    conversation.title || DEFAULT_CONVERSATION_TITLE
  );

  const normalizedTitle =
    currentTitle === DEFAULT_CONVERSATION_TITLE
      ? getTitleFromFirstUserMessage(safeMessages)
      : currentTitle;

  return {
    id: conversation.id || createId(),
    dataverseConversationRowId: conversation.dataverseConversationRowId,
    copilotConversationId: conversation.copilotConversationId,
    title: normalizedTitle,
    createdAt: conversation.createdAt || now,
    updatedAt: conversation.updatedAt || conversation.createdAt || now,
    messages: safeMessages,
  };
}

function createEmptyConversation(): SavedConversation {
  const now = new Date().toISOString();

  return {
    id: createId(),
    dataverseConversationRowId: undefined,
    copilotConversationId: undefined,
    title: DEFAULT_CONVERSATION_TITLE,
    createdAt: now,
    updatedAt: now,
    messages: [],
  };
}


function mapDataverseConversationToSavedConversation(
  conversation: DataverseConversation
): SavedConversation {
  const now = new Date().toISOString();
  const rowId = conversation.conversationRowId || createId();
  const title = cleanConversationTitle(
    conversation.titulo || DEFAULT_CONVERSATION_TITLE
  );

  return {
    id: conversation.localConversationId || rowId,
    dataverseConversationRowId: conversation.conversationRowId || rowId,
    copilotConversationId: conversation.copilotConversationId || undefined,
    title,
    createdAt: conversation.fechaCreacion || now,
    updatedAt: conversation.fechaCreacion || now,
    messages: [],
  };
}

function normalizeRoleFromDataverse(role?: string): ChatRole {
  const normalizedRole = (role || "").trim().toLowerCase();

  if (normalizedRole === "bot" || normalizedRole === "puerto emplea") {
    return "bot";
  }

  return "user";
}

function mapDataverseMessageToSavedMessage(message: DataverseMessage): SavedMessage {
  return {
    id: createId(),
    role: normalizeRoleFromDataverse(message.rol),
    text: message.mensaje || "",
    createdAt: message.fecha || new Date().toISOString(),
  };
}

function isPuertoEmpleaGreeting(text: string): boolean {
  const normalizedText = text
    .trim()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase();

  return (
    normalizedText.includes("hola, soy puerto emplea") &&
    normalizedText.includes("en que puedo ayudarle")
  );
}

function loadStoredConversations(): SavedConversation[] {
  try {
    const raw = window.localStorage.getItem(activeStorageKey);

    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw) as unknown;

    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map((conversation) =>
      normalizeConversation(conversation as Partial<SavedConversation>)
    );

    saveStoredConversations(normalized);

    return normalized;
  } catch {
    return [];
  }
}

function saveStoredConversations(conversations: SavedConversation[]): void {
  window.localStorage.setItem(activeStorageKey, JSON.stringify(conversations));
}

function activateStorageForUser(username: string): SavedConversation[] {
  activeStorageKey = createUserStorageKey(username);

  const userConversations = loadStoredConversations();

  if (userConversations.length > 0) {
    return userConversations;
  }

  try {
    const legacyRaw = window.localStorage.getItem(DEFAULT_STORAGE_KEY);

    if (!legacyRaw) {
      return [];
    }

    const legacyParsed = JSON.parse(legacyRaw) as unknown;

    if (!Array.isArray(legacyParsed)) {
      return [];
    }

    const legacyConversations = legacyParsed.map((conversation) =>
      normalizeConversation(conversation as Partial<SavedConversation>)
    );

    if (legacyConversations.length > 0) {
      saveStoredConversations(legacyConversations);
      window.localStorage.removeItem(DEFAULT_STORAGE_KEY);
    }

    return legacyConversations;
  } catch {
    return [];
  }
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
  const usuarioRef = useRef("");
  const [mensaje, setMensaje] = useState("Preparando Puerto Emplea...");
  const [necesitaLogin, setNecesitaLogin] = useState(false);
  const [cargando, setCargando] = useState(false);
  const [msalListo, setMsalListo] = useState(false);

  const [accessTokenActual, setAccessTokenActual] = useState("");

  const [conversations, setConversations] = useState<SavedConversation[]>(
    initialChatData.conversations
  );

  const conversationsRef = useRef<SavedConversation[]>(
    initialChatData.conversations
  );

  const [activeConversationId, setActiveConversationId] = useState(
    initialChatData.activeConversationId
  );

  const activeConversationIdRef = useRef<string>(
    initialChatData.activeConversationId
  );

  const liveConversationIdRef = useRef<string | null>(
    initialChatData.activeConversationId
  );

  const [liveConversationId, setLiveConversationId] = useState<string | null>(
    initialChatData.activeConversationId
  );

  const ocultarSaludoInicialRef = useRef(false);

  const [webChatKey, setWebChatKey] = useState(createId());
  const [modoHistorial, setModoHistorial] = useState(false);
  const [, setResumedConversationId] = useState<string | null>(null);

  const estaEnIframe = window.self !== window.top;
  async function callDataverseFlow(
    payload: Record<string, unknown>
  ): Promise<DataverseFlowResponse | undefined> {
    if (!DATAVERSE_FLOW_URL) {
  console.warn("DATAVERSE_FLOW_URL todavía no está configurada.");
  return undefined;
}

    try {
      const response = await fetch(DATAVERSE_FLOW_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const responseText = await response.text();

      if (!response.ok) {
        console.error("Error desde Power Automate:", response.status, responseText);
        return undefined;
      }

      if (!responseText) {
        return { ok: true };
      }

      return JSON.parse(responseText) as DataverseFlowResponse;
    } catch (error) {
      console.error("No se pudo llamar al flujo de Dataverse:", error);
      return undefined;
    }
  }

  async function cargarConversacionesDesdeDataverse(
    username: string
  ): Promise<SavedConversation[]> {
    const result = await callDataverseFlow({
      accion: "listarConversaciones",
      usuarioEmail: username,
    });

    if (!Array.isArray(result?.conversaciones)) {
      return [];
    }

    return result.conversaciones
      .map(mapDataverseConversationToSavedConversation)
      .filter((conversation) => conversation.dataverseConversationRowId);
  }

  async function cargarMensajesDesdeDataverse(
    conversation: SavedConversation
  ): Promise<SavedMessage[]> {
    if (!conversation.dataverseConversationRowId) {
      return conversation.messages;
    }

    const result = await callDataverseFlow({
      accion: "obtenerMensajes",
      conversationRowId: conversation.dataverseConversationRowId,
      usuarioEmail: usuarioRef.current,
    });

    if (!Array.isArray(result?.mensajes)) {
      return conversation.messages;
    }

    return result.mensajes
      .map(mapDataverseMessageToSavedMessage)
      .filter((message) => message.text.trim());
  }

  function updateConversationMessages(
    localConversationId: string,
    messages: SavedMessage[]
  ): void {
    setConversations((previous) => {
      const next = previous.map((conversation) => {
        if (conversation.id !== localConversationId) {
          return conversation;
        }

        return {
          ...conversation,
          messages,
          updatedAt: messages[messages.length - 1]?.createdAt || conversation.updatedAt,
        };
      });

      conversationsRef.current = next;
      saveStoredConversations(next);

      return next;
    });
  }

  function updateDataverseConversationRowId(
    localConversationId: string,
    dataverseConversationRowId?: string
  ): void {
    if (!dataverseConversationRowId) {
      return;
    }

    setConversations((previous) => {
      let changed = false;

      const next = previous.map((conversation) => {
        if (conversation.id !== localConversationId) {
          return conversation;
        }

        if (conversation.dataverseConversationRowId === dataverseConversationRowId) {
          return conversation;
        }

        changed = true;

        return {
          ...conversation,
          dataverseConversationRowId,
        };
      });

      if (changed) {
        conversationsRef.current = next;
        saveStoredConversations(next);
      }

      return next;
    });
  }

  async function registrarConversacionEnDataverse(
    conversation: SavedConversation,
    username: string,
    copilotConversationId?: string
  ): Promise<void> {
    const result = await callDataverseFlow({
      accion: "crearConversacion",
      localConversationId: conversation.id,
      titulo: conversation.title,
      usuario: username,
      usuarioEmail: username,
      conversationRowId: "",
      copilotConversationId: copilotConversationId || conversation.copilotConversationId || "",
      rol: "",
      mensaje: "",
    });

    if (result?.conversationRowId) {
      updateDataverseConversationRowId(conversation.id, result.conversationRowId);
    }
  }


  async function registrarMensajeEnDataverse(
    localConversationId: string,
    role: ChatRole,
    text: string,
    copilotConversationId?: string
  ): Promise<void> {
    const cleanText = text?.trim();

    if (!cleanText) {
      return;
    }

    const conversation = conversationsRef.current.find(
      (currentConversation) => currentConversation.id === localConversationId
    );

    if (!conversation?.dataverseConversationRowId) {
      console.warn(
        "No se pudo guardar el mensaje en Dataverse porque la conversación no tiene GUID de Dataverse.",
        {
          localConversationId,
          role,
          cleanText,
        }
      );
      return;
    }

    await callDataverseFlow({
      accion: "guardarMensaje",
      localConversationId,
      titulo: conversation.title,
      usuario: usuarioRef.current,
      usuarioEmail: usuarioRef.current,
      conversationRowId: conversation.dataverseConversationRowId,
      copilotConversationId:
        copilotConversationId || conversation.copilotConversationId || "",
      rol: role,
      mensaje: cleanText,
    });
  }

  function updateCopilotConversationId(
    localConversationId: string,
    copilotConversationId?: string
  ): void {
    if (!copilotConversationId) {
      return;
    }

    setConversations((previous) => {
      let changed = false;

      const next = previous.map((conversation) => {
        if (conversation.id !== localConversationId) {
          return conversation;
        }

        if (conversation.copilotConversationId === copilotConversationId) {
          return conversation;
        }

        changed = true;

        return {
          ...conversation,
          copilotConversationId,
        };
      });

      if (changed) {
        conversationsRef.current = next;
        saveStoredConversations(next);
      }

      return next;
    });
  }

  function setLiveConversation(conversationId: string | null): void {
    liveConversationIdRef.current = conversationId;
    setLiveConversationId(conversationId);
  }

  function createAndActivateLocalConversation(): SavedConversation {
    const newConversation = createEmptyConversation();

    setConversations((previous) => {
      const next = [newConversation, ...previous];
      conversationsRef.current = next;
      saveStoredConversations(next);
      return next;
    });

    setActiveConversationId(newConversation.id);
    activeConversationIdRef.current = newConversation.id;
    setLiveConversation(newConversation.id);

    setModoHistorial(false);
    setResumedConversationId(null);

    return newConversation;
  }

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

        const hasPreviousUserMessage = conversation.messages.some(
          (message) => message.role === "user" && message.text?.trim()
        );

        const shouldRenameConversation =
          role === "user" &&
          !hasPreviousUserMessage &&
          conversation.title === DEFAULT_CONVERSATION_TITLE;

        return {
          ...conversation,
          title: shouldRenameConversation
            ? createConversationTitle(cleanText)
            : conversation.title,
          updatedAt: now,
          messages: [...conversation.messages, newMessage],
        };
      });

      conversationsRef.current = next;
      saveStoredConversations(next);

      return next;
    });
  }

  function createWebChatStore() {
    return createStore({}, () => (next: any) => (action: any) => {
      const currentLiveConversationId = liveConversationIdRef.current;

      if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
        const activity = action.payload?.activity;
        const realCopilotConversationId = activity?.conversation?.id;

        if (currentLiveConversationId && realCopilotConversationId) {
          updateCopilotConversationId(
            currentLiveConversationId,
            realCopilotConversationId
          );
        }

        if (
          currentLiveConversationId &&
          activity?.from?.role === "bot" &&
          activity?.type === "message" &&
          activity?.text
        ) {
          if (ocultarSaludoInicialRef.current) {
            ocultarSaludoInicialRef.current = false;

            if (isPuertoEmpleaGreeting(activity.text)) {
              return;
            }
          }

          appendMessageToConversation(
            currentLiveConversationId,
            "bot",
            activity.text
          );

          void registrarMensajeEnDataverse(
            currentLiveConversationId,
            "bot",
            activity.text,
            realCopilotConversationId
          );

          window.parent.postMessage(
            {
              type: "BOT_RESPONSE",
              text: activity.text,
              conversationId: realCopilotConversationId || "",
            },
            "*"
          );
        }
      }

      if (action.type === "WEB_CHAT/SEND_MESSAGE") {
        const userText = action.payload?.text || "";

        if (currentLiveConversationId) {
          appendMessageToConversation(
            currentLiveConversationId,
            "user",
            userText
          );

          void registrarMensajeEnDataverse(
            currentLiveConversationId,
            "user",
            userText
          );
        }

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

  async function createCopilotConnectionForConversation(
    localConversationId: string,
    accessToken: string,
    loadingMessage: string,
    shouldResumeExistingConversation: boolean
  ): Promise<void> {
    const conversationToOpen = conversationsRef.current.find(
      (conversation) => conversation.id === localConversationId
    );

    const realCopilotConversationId = shouldResumeExistingConversation
      ? conversationToOpen?.copilotConversationId
      : undefined;

    setActiveConversationId(localConversationId);
    activeConversationIdRef.current = localConversationId;
    setLiveConversation(localConversationId);
    setModoHistorial(false);

    const newStore = createWebChatStore();

    setStore(newStore);
    setWebChatKey(createId());
    setConnection(null);
    setMensaje(loadingMessage);

    if (shouldResumeExistingConversation) {
      ocultarSaludoInicialRef.current = true;
    }

    const client = new CopilotStudioClient(settings, accessToken);

    const nuevaConexion = await CopilotStudioWebChat.createConnection(client, {
      showTyping: true,
      ...(realCopilotConversationId
        ? { conversationId: realCopilotConversationId }
        : {}),
    });

    if (nuevaConexion.conversationId) {
      updateCopilotConversationId(
        localConversationId,
        nuevaConexion.conversationId
      );
    }

    setConnection(nuevaConexion);
    setNecesitaLogin(false);
    setModoHistorial(false);
    setMensaje("");
  }

  const conectarConToken = async (
    username: string,
    accessToken: string
  ): Promise<void> => {
    setCargando(true);

    try {
      setUsuario(username);
      usuarioRef.current = username;
      setAccessTokenActual(accessToken);
      setMensaje("Conectando con Puerto Emplea...");

      const localUserConversations = activateStorageForUser(username);
      const dataverseUserConversations =
        await cargarConversacionesDesdeDataverse(username);

      const storedUserConversations =
        dataverseUserConversations.length > 0
          ? dataverseUserConversations
          : localUserConversations;

      const newConversation = createEmptyConversation();
      const nextConversations = [newConversation, ...storedUserConversations];

      conversationsRef.current = nextConversations;
      saveStoredConversations(nextConversations);
      setConversations(nextConversations);
      setActiveConversationId(newConversation.id);
      activeConversationIdRef.current = newConversation.id;
      setLiveConversation(newConversation.id);
      setModoHistorial(false);
      setResumedConversationId(null);

      await registrarConversacionEnDataverse(newConversation, username);

      await createCopilotConnectionForConversation(
        newConversation.id,
        accessToken,
        "Conectando con Puerto Emplea...",
        false
      );

      setResumedConversationId(null);
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
      usuarioRef.current = cuenta.username;
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
    if (!accessTokenActual || !usuarioRef.current) {
      setMensaje("Inicia sesión para crear una nueva conversación.");
      setNecesitaLogin(true);
      return;
    }

    setCargando(true);

    try {
      const newConversation = createAndActivateLocalConversation();

      await registrarConversacionEnDataverse(newConversation, usuarioRef.current);

      await createCopilotConnectionForConversation(
        newConversation.id,
        accessTokenActual,
        "Creando nueva conversación...",
        false
      );

      setResumedConversationId(null);
    } catch (error) {
      console.error("Error creando nueva conversación:", error);
      setMensaje("No se pudo crear una nueva conversación.");
      setNecesitaLogin(false);
    } finally {
      setCargando(false);
    }
  };

  const seleccionarConversacion = async (
    conversationId: string
  ): Promise<void> => {
    setActiveConversationId(conversationId);
    activeConversationIdRef.current = conversationId;

    if (conversationId === liveConversationIdRef.current && connection) {
      setModoHistorial(false);
      return;
    }

    const selectedConversation = conversationsRef.current.find(
      (conversation) => conversation.id === conversationId
    );

    let selectedConversationWithMessages = selectedConversation;

    if (selectedConversation?.dataverseConversationRowId) {
      setCargando(true);

      try {
        const dataverseMessages = await cargarMensajesDesdeDataverse(
          selectedConversation
        );

        updateConversationMessages(selectedConversation.id, dataverseMessages);

        selectedConversationWithMessages = {
          ...selectedConversation,
          messages: dataverseMessages,
        };
      } catch (error) {
        console.error("Error cargando mensajes desde Dataverse:", error);
      } finally {
        setCargando(false);
      }
    }

    if (!selectedConversationWithMessages?.copilotConversationId) {
      if (liveConversationIdRef.current && conversationId !== liveConversationIdRef.current) {
        setLiveConversation(null);
        setConnection(null);
        setWebChatKey(createId());
        setResumedConversationId(null);
      }

      setModoHistorial(true);
      return;
    }

    if (!accessTokenActual || !usuarioRef.current) {
      setModoHistorial(true);
      return;
    }

    setCargando(true);

    try {
      await createCopilotConnectionForConversation(
        conversationId,
        accessTokenActual,
        "Abriendo conversación...",
        true
      );

      setResumedConversationId(conversationId);
    } catch (error) {
      console.error("Error abriendo conversación:", error);
      setModoHistorial(true);
      setResumedConversationId(null);
      setMensaje("No se pudo abrir la conversación para continuar.");
    } finally {
      setCargando(false);
    }
  };

  const eliminarConversacion = (conversationId: string): void => {
    const conversationToDelete = conversations.find(
      (conversation) => conversation.id === conversationId
    );

    const titleToShow = conversationToDelete?.title || "esta conversación";

    const confirmar = window.confirm(
      `¿Desea eliminar la conversación "${titleToShow}"?`
    );

    if (!confirmar) {
      return;
    }

    const nextConversations = conversations.filter(
      (conversation) => conversation.id !== conversationId
    );

    if (nextConversations.length === 0) {
      const newConversation = createEmptyConversation();

      conversationsRef.current = [newConversation];
      saveStoredConversations([newConversation]);

      setConversations([newConversation]);
      setActiveConversationId(newConversation.id);

      activeConversationIdRef.current = newConversation.id;
      setLiveConversation(null);
      setResumedConversationId(null);

      setModoHistorial(true);
      setConnection(null);
      setWebChatKey(createId());

      return;
    }

    conversationsRef.current = nextConversations;
    saveStoredConversations(nextConversations);
    setConversations(nextConversations);

    const deletedActiveConversation = conversationId === activeConversationId;
    const deletedLiveConversation = conversationId === liveConversationIdRef.current;

    if (deletedLiveConversation) {
      setLiveConversation(null);
      setResumedConversationId(null);
      setConnection(null);
      setWebChatKey(createId());
    }

    if (!deletedActiveConversation) {
      return;
    }

    const nextActiveConversation = nextConversations[0];

    setActiveConversationId(nextActiveConversation.id);
    activeConversationIdRef.current = nextActiveConversation.id;

    setModoHistorial(true);
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

        const username = data.username || "Usuario autenticado";
        usuarioRef.current = username;

        await conectarConToken(
          username,
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
            usuarioRef.current = redirectResult.account.username;
            await conectarConToken(
              redirectResult.account.username,
              redirectResult.accessToken
            );
            setMsalListo(true);
            return;
          }

          if (redirectResult?.account) {
            usuarioRef.current = redirectResult.account.username;
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

        usuarioRef.current = cuentas[0].username;
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
      <aside
        className="sidebar"
        style={{
          width: "220px",
          minWidth: "220px",
          maxWidth: "220px",
        }}
      >
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
          className="conversation-list"
          style={{
            flex: 1,
            minHeight: 0,
            overflowY: "auto",
            paddingRight: "4px",
          }}
        >
          {conversations.map((conversation) => {
            const isLiveConversation = conversation.id === liveConversationId;
            const isSelectedConversation = conversation.id === activeConversation?.id;

            return (
              <div
                key={conversation.id}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  marginBottom: "6px",
                }}
              >
                <button
                  className={
                    isSelectedConversation ? "history-item active" : "history-item"
                  }
                  style={{
                    flex: 1,
                    marginBottom: 0,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    gap: "8px",
                  }}
                  onClick={() => seleccionarConversacion(conversation.id)}
                  title={conversation.title}
                >
                  <span
                    style={{
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {conversation.title}
                  </span>

                  {isLiveConversation && (
                    <span
                      style={{
                        padding: "2px 6px",
                        borderRadius: "999px",
                        background: "#16a34a",
                        color: "#ffffff",
                        fontSize: "10px",
                        fontWeight: 700,
                        flexShrink: 0,
                      }}
                    >
                      ACTIVA
                    </span>
                  )}
                </button>

                <button
                  type="button"
                  title="Eliminar conversación"
                  onClick={() => eliminarConversacion(conversation.id)}
                  style={{
                    width: "34px",
                    minWidth: "34px",
                    height: "34px",
                    border: "none",
                    borderRadius: "8px",
                    background: "#dc2626",
                    color: "#ffffff",
                    cursor: "pointer",
                    fontWeight: "bold",
                    flexShrink: 0,
                  }}
                >
                  ×
                </button>
              </div>
            );
          })}
        </div>

        <div className="user-info">{usuario}</div>
      </aside>

      <main
        className="chat-area"
        style={{
          flex: 1,
          minWidth: 0,
        }}
      >
        <div className="chat-wrapper">
          {modoHistorial ? (
            <div className="saved-history">
              <div className="saved-history-header">
                <strong>{activeConversation?.title || "Conversación"}</strong>
                <span>Historial guardado · solo lectura</span>
              </div>

              <div
                style={{
                  marginBottom: "12px",
                  padding: "10px 12px",
                  borderRadius: "10px",
                  background: "#eef6ff",
                  color: "#0d3b66",
                  fontSize: "13px",
                  border: "1px solid #d7e8fb",
                }}
              >
                El historial de esta conversación se muestra en el panel derecho.
                {activeConversation?.copilotConversationId
                  ? " Puedes reabrirla desde el botón del panel derecho."
                  : " Esta conversación no tiene identificador real de Copilot guardado, por lo que se mantiene como solo lectura."}
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

      {activeConversation?.messages?.length ? (
        <aside
          style={{
            width: "280px",
            minWidth: "280px",
            height: "100%",
            borderLeft: "1px solid #e5edf5",
            background: "#f8fbff",
            display: "flex",
            flexDirection: "column",
            overflow: "hidden",
          }}
        >
          <div
            style={{
              padding: "12px",
              borderBottom: "1px solid #e5edf5",
              background: "#ffffff",
            }}
          >
            <div
              style={{
                fontWeight: 800,
                color: "#0d3b66",
                marginBottom: "4px",
              }}
            >
              Historial de la conversación
            </div>

            <div
              style={{
                fontSize: "12px",
                color: "#6b7280",
                marginBottom: "12px",
              }}
            >
              {activeConversation.title}
            </div>

            {activeConversation.id === liveConversationId && !modoHistorial ? (
              <span
                style={{
                  display: "inline-block",
                  padding: "4px 8px",
                  borderRadius: "999px",
                  background: "#16a34a",
                  color: "#ffffff",
                  fontSize: "11px",
                  fontWeight: 800,
                }}
              >
                ACTIVA
              </span>
            ) : (
              <button
                type="button"
                onClick={() =>
                  activeConversation && seleccionarConversacion(activeConversation.id)
                }
                disabled={
                  cargando ||
                  !activeConversation?.copilotConversationId ||
                  !accessTokenActual ||
                  !usuarioRef.current
                }
                style={{
                  width: "100%",
                  border: "none",
                  borderRadius: "10px",
                  background:
                    activeConversation?.copilotConversationId && accessTokenActual && usuarioRef.current
                      ? "#0d3b66"
                      : "#9ca3af",
                  color: "#ffffff",
                  padding: "10px 12px",
                  cursor:
                    activeConversation?.copilotConversationId && accessTokenActual && usuarioRef.current
                      ? "pointer"
                      : "not-allowed",
                  fontWeight: 700,
                }}
              >
                🔄 Continuar conversación
              </button>
            )}

            {!activeConversation?.copilotConversationId && (
              <div
                style={{
                  marginTop: "10px",
                  padding: "9px 10px",
                  borderRadius: "9px",
                  background: "#fff7ed",
                  color: "#9a3412",
                  fontSize: "12px",
                  border: "1px solid #fed7aa",
                }}
              >
                Esta conversación fue creada antes de guardar el identificador real
                de Copilot. Puedes verla como historial, pero no reabrirla con
                contexto real.
              </div>
            )}
          </div>

          <div
            style={{
              flex: 1,
              minHeight: 0,
              overflowY: "auto",
              padding: "14px",
              display: "flex",
              flexDirection: "column",
              gap: "10px",
            }}
          >
            {activeConversation.messages.map((message) => (
              <div
                key={message.id}
                className={
                  message.role === "user"
                    ? "saved-message saved-message-user"
                    : "saved-message saved-message-bot"
                }
                style={{ maxWidth: "100%" }}
              >
                <div className="saved-message-role">
                  {message.role === "user" ? "Tú" : "Puerto Emplea"}
                </div>

                <div className="saved-message-text">{message.text}</div>
              </div>
            ))}
          </div>
        </aside>
      ) : null}
    </div>
  );
}

export default Chat;
