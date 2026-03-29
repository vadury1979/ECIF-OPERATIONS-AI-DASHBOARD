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

function renderDashboard(data) {

  document.getElementById("app-root").innerHTML = `
    <div class="main">

      <div class="card">
        <h2>Total Cases</h2>
        <h1>${data.totalCases}</h1>
      </div>

      <div class="card">
        <h2>Active Cases</h2>
        <h1>${data.activeCases}</h1>
      </div>

      <div class="card">
        <h2>Completed</h2>
        <h1>${data.completedCases}</h1>
      </div>

      <div class="card">
        <h2>Cancelled</h2>
        <h1>${data.cancelledCases}</h1>
      </div>

      <div class="card">
        <h2>Investment</h2>
        <h1>$${data.investment.toLocaleString()}</h1>
      </div>

    </div>
  `;
}

(async function main() {

  initTeamsIfAvailable();

  try {

    const data = await loadData();

    renderDashboard(data);

    console.log("✅ Data Loaded");

  }
  catch(err) {

    document.getElementById("app-root").innerHTML =
      "<h2>❌ Failed to load data.json</h2>";

  }

})();
