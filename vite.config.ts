import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import mkcert from "vite-plugin-mkcert";

const base = process.env.VITE_BASE_URL ?? "/";

export default defineConfig({
  base,
  plugins: [react(), mkcert()],
  server: {
    port: 3000,
    https: true,
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: "src/taskpane/index.html",
      },
    },
  },
});
