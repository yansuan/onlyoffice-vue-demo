export type AppConfig = {
  documentServerUrl: string
  callbackBaseUrl: string
  wsBaseUrl: string
  jwtSecret: string
  documentPath: string
}

/** Prefer runtime /config.js (Docker -e); fall back to Vite .env for local dev. */
export function getAppConfig(): AppConfig {
  const runtime = window.__APP_CONFIG__
  return {
    documentServerUrl:
      runtime?.VITE_DOCUMENT_SERVER_URL || import.meta.env.VITE_DOCUMENT_SERVER_URL,
    callbackBaseUrl: runtime?.VITE_CALLBACK_BASE_URL || import.meta.env.VITE_CALLBACK_BASE_URL,
    wsBaseUrl: runtime?.VITE_WS_BASE_URL || import.meta.env.VITE_WS_BASE_URL,
    jwtSecret: runtime?.VITE_ONLYOFFICE_JWT_SECRET || import.meta.env.VITE_ONLYOFFICE_JWT_SECRET,
    documentPath: runtime?.VITE_DOCUMENT_PATH || import.meta.env.VITE_DOCUMENT_PATH,
  }
}

/** Build typed WebSocket URL from VITE_WS_BASE_URL, e.g. type=vue | type=plugin */
export function buildWsUrl(type: 'vue' | 'plugin', wsBaseUrl = getAppConfig().wsBaseUrl): string {
  const base = wsBaseUrl.replace(/\/$/, '').replace(/\?.*$/, '')
  return `${base}?type=${type}`
}
