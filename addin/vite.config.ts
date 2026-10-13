import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

// In production the helper serves dist/ itself, so the API is same-origin. In
// development the pane runs on the Vite server and /api is proxied to the helper.
// The helper only accepts its own Origin, so the proxy forces that header.
const HELPER = "https://localhost:47613";

export default defineConfig({
  base: "./",
  plugins: [react()],
  build: {
    target: "es2019",
    cssTarget: "safari13",
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: false,
  },
  server: {
    port: 5173,
    strictPort: true,
    proxy: {
      "/api": {
        target: HELPER,
        secure: false,
        changeOrigin: true,
        // Lower-case on purpose: it must replace the browser's own "origin" header.
        headers: { origin: HELPER },
        configure(proxy) {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.setHeader("origin", HELPER);
          });
        },
      },
    },
  },
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
  },
});
