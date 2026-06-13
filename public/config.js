/*
 * Presentation Editor — host UI configuration.
 *
 * Edit this file on your server to show or hide editor UI for embedders. It is a
 * plain static file loaded before the app, so NO REBUILD is needed — change a
 * value, redeploy this one file (or edit it in place), reload the editor.
 *
 * Set any flag to false to hide that piece of UI. Omit a flag to keep its
 * default (shown). You can also override per-embed from the URL, e.g.:
 *   /?ui=fileMenu:0,present:0,tabs.insert:0
 */
window.presentationEditorConfig = {
  ui: {
    ribbon: true,            // the whole top toolbar
    fileMenu: true,          // the File tab + menu
    save: true,              // Save / Download .pptx (toolbar button + menu item)
    open: true,              // Open a .pptx
    export: true,            // Export as PDF / PNG menu items
    importPattern: true,     // Import pattern (JSON)
    newPresentation: true,   // Create new
    present: true,           // the Present (slideshow) button
    leftRail: true,          // left icon rail (find / thumbnails / about)
    slidePanel: true,        // slide thumbnail panel
    rightPanel: true,        // the format / settings panel
    statusBar: true,         // bottom status bar
    notesBar: true,          // speaker-notes bar
    docTitle: true,          // the editable presentation-title field
    tabs: {                  // individual ribbon tabs
      home: true,
      insert: true,
      design: true,
      transitions: true,
      view: true,
    },
  },

  // Authoritative permissions the EDITOR enforces (not just UI visibility).
  // With export:false, the scripting API methods exportPDF / exportSlidePNG /
  // exportPNGZip REJECT on every path — UI, keyboard, and pe:invoke from the
  // console — instead of returning a file. (Override per-embed with ?perm=export:0.)
  permissions: {
    export: true,
  },

  // Optional: per-export server authorization. When set, the editor POSTs to this
  // URL (credentials included) before each export and proceeds ONLY on a 2xx
  // response — so your backend can approve/deny each export per user. Leave unset
  // to skip. Set it here (server-served), not in the URL.
  // exportAuthUrl: "https://your-app.example.com/api/authorize-export",
};
