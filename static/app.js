const picker = document.getElementById("files");
const selected = document.getElementById("selected");
const processButton = document.getElementById("process");
const status = document.getElementById("status");
const results = document.getElementById("results");

let chosen = [];

picker.addEventListener("change", () => {
  chosen = [...picker.files];
  renderSelected();
});

function renderSelected() {
  selected.innerHTML = chosen.map(f =>
    `<div class="file"><span>${escapeHtml(f.name)}</span><span>${formatBytes(f.size)}</span></div>`
  ).join("");
  processButton.disabled = chosen.length === 0;
}

processButton.addEventListener("click", async () => {
  if (!chosen.length) return;

  processButton.disabled = true;
  results.innerHTML = "";
  status.textContent = "Processing PDF(s)…";

  const form = new FormData();
  chosen.forEach(file => form.append("files", file));

  try {
    const response = await fetch("/api/process", { method: "POST", body: form });
    const data = await response.json();

    if (!response.ok) throw new Error(data.error || "Processing failed.");

    status.textContent = `${data.files.length} Excel file(s) created.`;

    data.files.forEach(item => {
      const div = document.createElement("div");
      div.className = "result";
      div.innerHTML =
        `<b>${escapeHtml(item.filename)}</b>` +
        `<div>${item.rows} line item(s)</div>` +
        `<a href="/api/download?path=${encodeURIComponent(item.path)}">Download Excel</a>`;
      results.appendChild(div);
    });
  } catch (err) {
    status.textContent = "Error: " + err.message;
  } finally {
    processButton.disabled = false;
  }
});

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;"
  }[c]));
}
