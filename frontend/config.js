// Nova PDF Configuration
const CONFIG = {
  API_BASE: "https://file-converter-0a2q.onrender.com/api", // Prod
  // API_BASE: "http://localhost:5000/api", // Local Dev
};

// Handle module exports for potential Node/Build usage
if (typeof module !== 'undefined' && module.exports) {
  module.exports = CONFIG;
}

// Global exposure for browser script tags
window.CONFIG = CONFIG;