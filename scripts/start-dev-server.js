#!/usr/bin/env node
/* eslint-disable no-console */
const { spawn } = require("child_process");
const path = require("path");

// Launch the installed webpack CLI directly. Recent Node versions on Windows
// can fail to spawn npx.cmd with EINVAL.
const command = process.execPath;
const args = [
  path.resolve(__dirname, "..", "node_modules", "webpack", "bin", "webpack.js"),
  "serve",
  "--mode",
  "development",
];

let sawAddrInUse = false;

const child = spawn(command, args, {
  stdio: ["inherit", "pipe", "pipe"],
  env: process.env,
});

function handleChunk(chunk, target) {
  const text = chunk.toString();
  target.write(text);
  if (text.includes("EADDRINUSE") && text.includes("4001")) {
    sawAddrInUse = true;
  }
}

child.stdout.on("data", (chunk) => handleChunk(chunk, process.stdout));
child.stderr.on("data", (chunk) => handleChunk(chunk, process.stderr));

child.on("error", (err) => {
  console.error("[dev-server] Failed to start:", err.message);
  process.exit(1);
});

child.on("close", (code) => {
  if (sawAddrInUse) {
    console.log(
      "[dev-server] Port 4001 already in use. Reusing existing dev server (existing env flags stay active)."
    );
    console.log(
      "[dev-server] If desktop verbose logs are missing, stop the running dev server and start it again."
    );
    process.exit(0);
  }
  process.exit(code || 0);
});

