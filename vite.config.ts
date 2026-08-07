import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";

export default defineConfig({
  plugins: [react(), viteSingleFile()],
  build: {
    rollupOptions: {
      input: "views/form-shell.html",
    },
    outDir: "dist",
    emptyOutDir: false,
  },
});
