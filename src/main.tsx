import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { installApi } from "./util/api";
import "./styles.css";

// expose the public scripting API on window for same-origin host apps
installApi(window);

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
