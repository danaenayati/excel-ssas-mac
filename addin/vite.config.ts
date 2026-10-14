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

// `vite build --mode commands` (the third part) builds the right-click commands the same
// way: dist/commands.js, a classic script with a fixed name, loaded by commands.html (the
// manifest's FunctionFile), and message.html, the small page that shows what a command
// has to say. The helper puts the install token into commands.html, as into every page.
const COMMANDS_DIR = "src/commands/";

/** Copy files from `dir` into dist/ as they are. */
function copiedFiles(dir: string, names: string[]): Plugin {
  return {
    name: "cube-copied-files",
    generateBundle() {
      for (const name of names) {
        const path = fileURLToPath(new URL(dir + name, import.meta.url));
        this.emitFile({ type: "asset", fileName: name, source: readFileSync(path, "utf8") });
      }
    },
  };
}

/** One classic script, `<name>.js`, built from `dir/<name>.ts` next to the pane in dist/. */
function scriptBuild(dir: string, name: string, global: string, files: string[]) {
  return {
    publicDir: false as const,
    plugins: [copiedFiles(dir, files)],
    build: {
      target: "es2019",
      outDir: "dist",
      emptyOutDir: false, // the pane was built into dist/ just before
      sourcemap: false,
      lib: {
        entry: fileURLToPath(new URL(dir + name + ".ts", import.meta.url)),
        formats: ["iife" as const],
        name: global,
        fileName: () => name + ".js",
      },
    },
  };
}

export default defineConfig(({ mode }) =>
  mode === "functions"
    ? scriptBuild(FUNCTIONS_DIR, "functions", "CubeExplorerFunctions", ["functions.html", "functions.json"])
    : mode === "commands"
    ? scriptBuild(COMMANDS_DIR, "commands", "CubeExplorerCommands", ["commands.html", "message.html"])
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
