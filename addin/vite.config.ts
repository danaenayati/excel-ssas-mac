import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig, type Plugin } from "vitest/config";
import react from "@vitejs/plugin-react";

// In production the helper serves dist/ itself, so the API is same-origin. In
// development the pane runs on the Vite server and /api is proxied to the helper.
// The helper only accepts its own Origin, so the proxy forces that header.
const HELPER = "https://localhost:47613";

// `vite build --mode functions` (the second half of `npm run build`) builds the cube
// functions into dist/functions.js, functions.html and functions.json, next to the pane
// and without touching it. The script is one classic file with a fixed name: Excel runs
// it in a JavaScript-only runtime that loads exactly the URL in the manifest and has no
// modules, and the helper puts the install token into that one file when serving it.
const FUNCTIONS_DIR = "src/functions/";

function functionsFiles(): Plugin {
  return {
    name: "cube-functions-files",
    generateBundle() {
      for (const name of ["functions.html", "functions.json"]) {
        const path = fileURLToPath(new URL(FUNCTIONS_DIR + name, import.meta.url));
        this.emitFile({ type: "asset", fileName: name, source: readFileSync(path, "utf8") });
      }
    },
  };
}

export default defineConfig(({ mode }) =>
  mode === "functions"
    ? {
        publicDir: false,
        plugins: [functionsFiles()],
        build: {
          target: "es2019",
          outDir: "dist",
          emptyOutDir: false, // the pane was built into dist/ just before
          sourcemap: false,
          lib: {
            entry: fileURLToPath(new URL(FUNCTIONS_DIR + "functions.ts", import.meta.url)),
            formats: ["iife"],
            name: "CubeExplorerFunctions",
            fileName: () => "functions.js",
          },
        },
      }
    : {
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
      },
);
