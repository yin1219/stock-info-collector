import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "node:path";
export default defineConfig({
  root: path.resolve(import.meta.dirname, "src/demo"),
  plugins: [react()],
  server: { port: 5178, strictPort: true, host: "127.0.0.1" },
  build: {
    outDir: path.resolve(import.meta.dirname, "artifacts/ui-demo/dist"),
    emptyOutDir: true,
  },
});
