import React from "react";
import { createRoot } from "react-dom/client";
import App from "../citas-peluqueria.jsx";

// Sustituto local de window.storage (API de los artifacts de Claude) usando localStorage
if (!window.storage) {
  window.storage = {
    async get(key) {
      const value = localStorage.getItem(key);
      return value === null ? null : { key, value };
    },
    async set(key, value) {
      localStorage.setItem(key, value);
      return { key, value };
    },
    async delete(key) {
      localStorage.removeItem(key);
      return { key, deleted: true };
    },
  };
}

createRoot(document.getElementById("root")).render(<App />);
