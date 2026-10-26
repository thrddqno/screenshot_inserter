import { defineConfig } from "vite";
import type { Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { createHash } from "node:crypto";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { relative, resolve } from "node:path";

const VERSION_INPUTS = ["src", "taskpane.html", "package.json", "package-lock.json", "vite.config.ts"];
const repositoryPath = (path: string): string => relative(process.cwd(), path).replaceAll("\\", "/");

const isRuntimeVersionInput = (path: string): boolean => {
  const normalized = repositoryPath(path);
  return !normalized.startsWith("src/test/") && !/\.(?:test|spec)\.[cm]?[jt]sx?$/.test(normalized);
};

const collectFiles = (path: string): string[] => {
  if (!statSync(path).isDirectory()) return [path];
  return readdirSync(path)
    .flatMap((entry) => collectFiles(resolve(path, entry)))
    .sort();
};

const createBuildVersion = (): string => {
  const hash = createHash("sha256");
  const inputs = VERSION_INPUTS
    .flatMap((path) => collectFiles(resolve(path)))
    .filter(isRuntimeVersionInput)
    .sort();
  for (const input of inputs) {
    hash.update(repositoryPath(input));
    hash.update("\0");
    hash.update(readFileSync(input));
    hash.update("\0");
  }
  return hash.digest("hex").slice(0, 12);
};

const appVersionPlugin = (version: string): Plugin => ({
  name: "app-version",
  configureServer(server) {
    server.middlewares.use((request, response, next) => {
      if (request.url?.split("?", 1)[0] !== "/version.json") return next();
      response.statusCode = 200;
      response.setHeader("Content-Type", "application/json; charset=utf-8");
      response.setHeader("Cache-Control", "no-store");
      response.end(JSON.stringify({ version }));
    });
  },
  generateBundle() {
    this.emitFile({
      type: "asset",
      fileName: "version.json",
      source: `${JSON.stringify({ version })}\n`,
    });
  },
});

const getHttpsOptions = () => {
  const keyPath = resolve(".certs/localhost.key");
  const certPath = resolve(".certs/localhost.crt");
  if (!existsSync(keyPath) || !existsSync(certPath)) {
    throw new Error("Missing HTTPS certificate. Run: npm run certs:create");
  }
  return { key: readFileSync(keyPath), cert: readFileSync(certPath) };
};

export default defineConfig(({ command, mode }) => {
  const appVersion = createBuildVersion();
  return {
    base: command === "build" ? "/screenshot_inserter/" : "/",
    define: {
      __APP_VERSION__: JSON.stringify(appVersion),
    },
    plugins: [react(), appVersionPlugin(appVersion)],
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
        input: {
          docs: resolve("index.html"),
          taskpane: resolve("taskpane.html"),
        },
      },
    },
    test: {
      environment: "jsdom",
      setupFiles: ["./src/test/setup.ts"],
      coverage: {
        reporter: ["text", "html"],
      },
    },
  };
});
