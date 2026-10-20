import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

const getHttpsOptions = () => {
  const keyPath = resolve(".certs/localhost.key");
  const certPath = resolve(".certs/localhost.crt");
  if (!existsSync(keyPath) || !existsSync(certPath)) {
    throw new Error("Missing HTTPS certificate. Run: npm run certs:create");
  }
  return { key: readFileSync(keyPath), cert: readFileSync(certPath) };
};

export default defineConfig(async ({ command, mode }) => ({
  base: command === "build" ? "/screenshot_inserter/" : "/",
  plugins: [react()],
  server: {
    host: "localhost",
    port: 3000,
    strictPort: true,
    https: command === "serve" && mode !== "test" ? getHttpsOptions() : undefined,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  preview: {
    host: "localhost",
    port: 3000,
    strictPort: true,
  },
  build: {
    target: "es2022",
    outDir: "docs",
    emptyOutDir: true,
    sourcemap: false,
    rollupOptions: {
      input: resolve("taskpane.html"),
    },
  },
  test: {
    environment: "jsdom",
    setupFiles: ["./src/test/setup.ts"],
    coverage: {
      reporter: ["text", "html"],
    },
  },
}));
