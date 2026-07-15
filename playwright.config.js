const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  timeout: 30_000,
  use: {
    baseURL: "http://127.0.0.1:8787",
    locale: "he-IL",
    channel: process.env.PLAYWRIGHT_CHANNEL || undefined,
  },
  webServer: {
    command: "python -m http.server 8787",
    url: "http://127.0.0.1:8787",
    reuseExistingServer: true,
  },
});
