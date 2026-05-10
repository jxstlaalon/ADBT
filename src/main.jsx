import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

window.storage = {
  async get(key, shared) {
    return localStorage.getItem(key) || null;
  },
  async set(key, value, shared) {
    localStorage.setItem(key, value);
  },
  async delete(key, shared) {
    localStorage.removeItem(key);
  },
  async list(prefix, shared) {
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key.startsWith(prefix)) {
        keys.push(key);
      }
    }
    return keys;
  }
};

import './index.css'

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)