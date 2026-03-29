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

let regionHTML = '';

  data.regionData.forEach(r => {

    regionHTML += `
      <div class="card">
        <h3>${r.region}</h3>
        <p>Cases: ${r.cases}</p>
        <p>Investment: $${r.investment.toLocaleString()}</p>
      </div>
    `;

  });

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

      <h2>🌍 Regional Snapshot</h2>
      
<div class="card">
  <h3>🔄 Cases with IO</h3>
  <p>${data.ioPercent}%</p>
</div>

<div class="card">
  <h3>📋 Cases with PO</h3>
  <p>${data.poPercent}%</p>
</div>

<div class="card">
  <h3>🧾 Cases with Invoice</h3>
  <p>${data.invoicePercent}%</p>
</div>

<div class="card">
  <h3>✅ Invoice Approved</h3>
  <p>${data.invoiceApprovedPercent}%</p>
</div>

<div class="card">
  <h3>⏱ Avg E2E Cycle</h3>
  <p>${data.e2eCycle} Days</p>
</div>


      ${regionHTML}

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
