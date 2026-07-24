// Configurazione pubblica dell'app CRM CDG.
// NON contiene segreti: l'app key Dropbox è pubblica (flusso PKCE senza secret).
// Il refresh token e l'app secret NON vanno mai qui: stanno nei GitHub Secrets (lato automazione).
window.CRM_CONFIG = {
  // App key dell'app Dropbox "CRM CDG - GitHub"
  dropboxAppKey: "jeihle42vxibgiv",

  // Deve combaciare ESATTAMENTE con un Redirect URI registrato nell'app Dropbox.
  // In locale usa http://localhost:8080/ (o simile); in produzione l'indirizzo GitHub Pages.
  redirectUri: "https://pigironcompany-cdg.github.io/crm-cdg-app/",

  // Percorsi dentro la App Folder di Dropbox (Apps/CRM CDG - GitHub)
  snapshotPath: "/snapshot.json",
  editsPath: "/edits.jsonl",

  // Titolo mostrato in intestazione
  appTitle: "CRM CDG"
};
