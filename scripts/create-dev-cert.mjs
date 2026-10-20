import { mkdirSync } from "node:fs";
import { spawnSync } from "node:child_process";

mkdirSync(".certs", { recursive: true, mode: 0o700 });

const result = spawnSync(
  "openssl",
  [
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-sha256",
    "-days",
    "825",
    "-nodes",
    "-keyout",
    ".certs/localhost.key",
    "-out",
    ".certs/localhost.crt",
    "-subj",
    "/CN=localhost",
    "-addext",
    "subjectAltName=DNS:localhost,IP:127.0.0.1",
  ],
  { stdio: "inherit" },
);

if (result.error) throw result.error;
if (result.status !== 0) process.exit(result.status ?? 1);

console.log("Created .certs/localhost.crt and .certs/localhost.key");
console.log("Trust localhost.crt in Windows Trusted Root Certification Authorities before sideloading.");
