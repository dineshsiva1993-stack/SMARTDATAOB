import React from "react";

/**
 * Operation Bulletin — Cost Sheet Builder
 * -----------------------------------------------------------------
 * This app is built as a single, self-contained HTML/CSS/JS file
 * (ob-app.html) — not as React components. Rewriting 7000+ lines of
 * vanilla DOM code into JSX/React state would be extremely risky
 * (near-certain to introduce subtle bugs) for very little benefit,
 * since the app already works perfectly on its own.
 *
 * Instead, this component embeds the working HTML file as-is inside
 * an <iframe>. This gives you 100% identical behavior to the HTML
 * file, with zero risk of CSS/JS clashing with the rest of your
 * React app, and it drops straight into any existing React project.
 *
 * SETUP (one-time):
 * 1. Copy `ob-app.html` into your React project's `public/` folder
 *    (Create React App / Vite / Next.js `public/` all work the same
 *    way) — e.g. `public/ob-app.html`.
 * 2. Copy this file in as `src/OBApp.jsx` (or wherever you keep
 *    components) and use it like any other component:
 *
 *      import OBApp from "./OBApp";
 *      <OBApp />
 *
 * 3. Whenever you get an updated ob-app.html from Claude, just
 *    replace the file in `public/` — no code changes needed here.
 *
 * Note on data storage: ob-app.html already includes a fallback for
 * `window.storage` (the Claude.ai-only API it was originally built
 * against) — it uses Firebase if you've configured it inside the
 * HTML file, otherwise it falls back to the browser's localStorage.
 * That fallback runs automatically once this is hosted outside
 * Claude.ai, so no changes are needed for it to keep saving data.
 */
export default function OBApp({ src = "/ob-app.html", title = "Operation Bulletin" }) {
  return (
    <iframe
      src={src}
      title={title}
      style={{
        width: "100%",
        height: "100vh",
        border: "none",
        display: "block",
      }}
    />
  );
}
