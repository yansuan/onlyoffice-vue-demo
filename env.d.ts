/// <reference types="vite/client" />

interface AppRuntimeConfig {
  VITE_DOCUMENT_SERVER_URL?: string
  VITE_CALLBACK_BASE_URL?: string
  VITE_ONLYOFFICE_JWT_SECRET?: string
  VITE_DOCUMENT_PATH?: string
}

interface ImportMetaEnv {
  readonly VITE_DOCUMENT_SERVER_URL: string
  readonly VITE_CALLBACK_BASE_URL: string
  readonly VITE_ONLYOFFICE_JWT_SECRET: string
  readonly VITE_DOCUMENT_PATH: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}

interface Window {
  __APP_CONFIG__?: AppRuntimeConfig | null
}
