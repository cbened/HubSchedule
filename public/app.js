const statusEl = document.getElementById("status");
const listEl = document.getElementById("submissionList");
const refreshButton = document.getElementById("refreshButton");
const template = document.getElementById("submissionTemplate");
const formMeta = document.getElementById("formMeta");

const REFRESH_MS = 60_000;
let refreshTimer;

Office.onReady(async () => {
  attachEvents();
  await loadSubmissions();
  refreshTimer = setInterval(loadSubmissions, REFRESH_MS);
});

function attachEvents() {
  refreshButton.addEventListener("click", () => {
    loadSubmissions();
  });
}

async function loadSubmissions() {
  setStatus("Fetching latest submissions…");
  refreshButton.disabled = true;

  try {
    const response = await fetch("/api/submissions");
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || "Failed to fetch submissions.");
    }

    renderMeta(payload.meta);
    renderSubmissions(payload.results || []);
    setStatus(`Loaded ${payload.results.length} submission(s). Last updated ${new Date().toLocaleTimeString()}.`);
  } catch (error) {
    setStatus(error.message || "Unexpected error.", true);
  } finally {
    refreshButton.disabled = false;
  }
}

function renderMeta(meta) {
  const portalText = meta.portalId ? `Portal ${meta.portalId}` : "Portal not set";
  formMeta.textContent = `Form ${meta.formId} · ${portalText}`;
}

function renderSubmissions(results) {
  listEl.innerHTML = "";

  if (!results.length) {
    listEl.innerHTML = `<p class="muted">No submissions found for this form yet.</p>`;
    return;
  }

  for (const submission of results) {
    const node = template.content.firstElementChild.cloneNode(true);
    const nameEl = node.querySelector("[data-name]");
    const dateEl = node.querySelector("[data-date]");
    const fieldsEl = node.querySelector("[data-fields]");

    const props = submission.values || {};
    const name = [props.firstname, props.lastname].filter(Boolean).join(" ") || props.email || "Unknown Contact";

    nameEl.textContent = name;
    dateEl.textContent = new Date(submission.submittedAt).toLocaleString();

    for (const [key, value] of Object.entries(props)) {
      const dt = document.createElement("dt");
      dt.textContent = key;

      const dd = document.createElement("dd");
      dd.textContent = String(value);

      fieldsEl.append(dt, dd);
    }

    listEl.appendChild(node);
  }
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("muted", !isError);
  statusEl.style.color = isError ? "#b91c1c" : "";
}

window.addEventListener("beforeunload", () => {
  if (refreshTimer) {
    clearInterval(refreshTimer);
  }
});
