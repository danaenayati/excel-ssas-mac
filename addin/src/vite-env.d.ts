/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** Install token for `npm run dev`, when the page is not served by the helper. */
  readonly VITE_CUBE_TOKEN?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
