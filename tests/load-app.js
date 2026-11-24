/**
 * Loads an app.js (the trial copy by default) outside the browser and returns
 * the pure functions it exposes as `ZeekrTracker`. The file runs in its own
 * sandbox with a fake address and empty storage, so nothing is fetched or saved.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

function loadApp(relativePath = "beta/app.js") {
  const file = path.join(__dirname, "..", relativePath);
  const memory = new Map();
  const fakeStorage = {
    getItem: (k) => (memory.has(k) ? memory.get(k) : null),
    setItem: (k, v) => memory.set(k, String(v)),
    removeItem: (k) => memory.delete(k),
  };
  const sandbox = {
    console,
    location: { origin: "https://example.test", pathname: "/Zeekr_Tracker/beta/" },
    localStorage: fakeStorage,
    sessionStorage: fakeStorage,
    URLSearchParams,
    TextEncoder,
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(file, "utf8"), sandbox, { filename: file });
  if (!sandbox.ZeekrTracker) throw new Error("app.js did not expose ZeekrTracker");
  return sandbox.ZeekrTracker;
}

module.exports = { loadApp };
