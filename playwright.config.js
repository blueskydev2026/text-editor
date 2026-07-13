const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  timeout: 30_000,
  use: { baseURL: "http://127.0.0.1:8787", locale: "he-IL" },
  webServer: {
    command: "python3 -m http.server 8787",
    url: "http://127.0.0.1:8787",
    reuseExistingServer: true,
  },
});
