export type AppConfig = {
  documentServerUrl: string
  callbackBaseUrl: string
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
    jwtSecret: runtime?.VITE_ONLYOFFICE_JWT_SECRET || import.meta.env.VITE_ONLYOFFICE_JWT_SECRET,
    documentPath: runtime?.VITE_DOCUMENT_PATH || import.meta.env.VITE_DOCUMENT_PATH,
  }
}
