import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";
import path from "node:path";

export default defineConfig({
  base: "/senai-cba/",
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["favicon.svg"],
      manifest: {
        name: "SENAI · Gestão de Estoque LIS",
        short_name: "Estoque LIS",
        description:
          "Sistema profissional de gestão de estoque com sincronização Excel Online via Power Automate.",
        theme_color: "#00529b",
        background_color: "#0b1220",
        display: "standalone",
        start_url: "/senai-cba/",
        scope: "/senai-cba/",
        icons: [
          {
            src: "favicon.svg",
            sizes: "any",
            type: "image/svg+xml",
            purpose: "any maskable",
          },
        ],
      },
      workbox: {
        navigateFallback: "/senai-cba/index.html",
        runtimeCaching: [
          {
            urlPattern: ({ url }) =>
              url.origin.includes("powerplatform") ||
              url.origin.includes("logic.azure"),
            handler: "NetworkOnly",
          },
        ],
      },
    }),
  ],
});
