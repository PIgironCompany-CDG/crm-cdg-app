// Configurazione pubblica dell'app CRM CDG. Nessun segreto qui.
// L'app parla SOLO con il backend (Cloudflare Worker), che autentica e custodisce i dati.
window.CRM_CONFIG = {
  // URL del backend Cloudflare Worker
  workerBase: "https://crm-cdg-api.j-scarparo.workers.dev",
  appTitle: "CRM CDG"
};
