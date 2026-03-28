const state = {
  status: "all",
  search: "",
  records: [],
  activeKey: null,
};

const summaryEls = {
  all: document.getElementById("stat-all"),
  open: document.getElementById("stat-open"),
  closed: document.getElementById("stat-closed"),
};

const tableBody = document.getElementById("claims-table-body");
const detailList = document.getElementById("detail-list");
const searchInput = document.getElementById("search-input");
const refreshButton = document.getElementById("refresh-button");
const tabButtons = [...document.querySelectorAll(".tab-button")];

async function fetchClaims() {
  const params = new URLSearchParams({
    status: state.status,
    search: state.search,
  });
  const response = await fetch(`/api/claims?${params.toString()}`);
  if (!response.ok) {
    throw new Error("Failed to load claims");
  }
  return response.json();
}

function renderSummary(summary) {
  summaryEls.all.textContent = summary.all ?? 0;
  summaryEls.open.textContent = summary.open ?? 0;
  summaryEls.closed.textContent = summary.closed ?? 0;
}

function renderTable(records) {
  tableBody.innerHTML = "";
  if (!records.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6">No claims found.</td>`;
    tableBody.appendChild(tr);
    return;
  }

  records.forEach((record) => {
    const tr = document.createElement("tr");
    if (record.key === state.activeKey) {
      tr.classList.add("active");
    }
    tr.innerHTML = `
      <td>${escapeHtml(record.claim_id || "-")}</td>
      <td>${escapeHtml(record.customer_name || record.title || "-")}</td>
      <td>${escapeHtml(record.insurance_company || "-")}</td>
      <td>${escapeHtml(record.claim_number || "-")}</td>
      <td>${escapeHtml(record.town || "-")}</td>
      <td>${escapeHtml(record.status || "-")}</td>
    `;
    tr.addEventListener("click", () => {
      state.activeKey = record.key;
      renderTable(records);
      renderDetails(record);
    });
    tableBody.appendChild(tr);
  });
}

function renderDetails(record) {
  if (!record) {
    detailList.innerHTML = `<div class="empty-state">Select a claim to view details.</div>`;
    return;
  }

  const fields = [
    ["Claim ID", record.claim_id],
    ["Customer", record.customer_name || record.title],
    ["Insurance", record.insurance_company],
    ["Claim #", record.claim_number],
    ["Status", record.status],
    ["Type", record.claim_type],
    ["Date of Loss", record.date_of_loss],
    ["Town", record.town],
    ["Owner Address", record.owner_address],
    ["Vehicle", record.vehicle],
    ["VIN", record.vin],
    ["Shop", record.shop_name],
    ["Phone", record.contact_phone],
    ["Email", record.contact_email],
    ["Office Progress", record.office_progress_status],
    ["Inspection", record.office_appt_when],
    ["Paperwork", record.office_waiting_for_paperwork],
    ["Appraisal Notes", record.assignment_claim_notes],
    ["Claim Notes", record.notes],
    ["Folder", record.source_path],
  ];

  detailList.innerHTML = fields
    .map(([label, value]) => {
      return `
        <div class="detail-item">
          <span class="detail-label">${escapeHtml(label)}</span>
          <div class="detail-value">${escapeHtml(value || "-")}</div>
        </div>
      `;
    })
    .join("");
}

async function refreshData() {
  const payload = await fetchClaims();
  state.records = payload.records || [];
  renderSummary(payload.summary || {});

  const activeStillExists = state.records.find((record) => record.key === state.activeKey);
  if (!activeStillExists) {
    state.activeKey = state.records[0]?.key || null;
  }

  renderTable(state.records);
  renderDetails(state.records.find((record) => record.key === state.activeKey) || null);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

tabButtons.forEach((button) => {
  button.addEventListener("click", async () => {
    state.status = button.dataset.status || "all";
    tabButtons.forEach((entry) => entry.classList.toggle("active", entry === button));
    await refreshData();
  });
});

searchInput.addEventListener("input", async (event) => {
  state.search = event.target.value.trim();
  await refreshData();
});

refreshButton.addEventListener("click", async () => {
  await refreshData();
});

refreshData().catch((error) => {
  detailList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
});
