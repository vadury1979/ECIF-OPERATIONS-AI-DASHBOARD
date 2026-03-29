async function loadData() {
  const url = window.__ECIF_DATA_URL__ || "./data.json";
  const res = await fetch(url, { cache: "no-store" });
  return await res.json();
}

function initTeamsIfAvailable() {
  try {
    if (window.microsoftTeams?.app) {
      window.microsoftTeams.app.initialize().then(async () => {

        document.body.classList.add("in-teams");

        const ctx = await window.microsoftTeams.app.getContext();
        const theme = ctx?.app?.theme || "default";

        document.body.classList.toggle("teams-dark", theme === "dark");
        document.body.classList.toggle("teams-contrast", theme === "contrast");

        window.microsoftTeams.app.registerOnThemeChangeHandler((t) => {
          document.body.classList.toggle("teams-dark", t === "dark");
          document.body.classList.toggle("teams-contrast", t === "contrast");
        });

      });
    }
  } catch (e) {
    console.log("Not inside Teams");
  }
}

function renderShell() {
  document.getElementById("app-root").innerHTML = `
      <div class="main">
          <div class="card">
              <h2>✅ ECIF Dashboard Loaded</h2>
              <p>If you are seeing this inside Teams – integration is working.</p>
          </div>
      </div>
  `;
}

(async function main() {
  initTeamsIfAvailable();
  renderShell();

  try {
    const data = await loadData(); 
    console.log("Data Loaded:", data);
  } catch (err) {
    console.log("No data.json yet");
  }
})();
