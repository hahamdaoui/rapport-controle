/* Interface de l'application « Rapport de contrôle » – JavaScript natif, sans dépendance. */
(() => {
  "use strict";
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const MAX_BYTES = 10 * 1024 * 1024;

  const state = { token: null, data: null, mapping: null, scale: "", lastSuggestion: "" };

  const icon = (n) => `<svg class="icon"><use href="${window.SPRITE}#${n}"/></svg>`;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = (v, nd = 2) => (v === null || v === undefined ? "—" : Number(v).toFixed(nd).replace(".", ","));
  const fmtGrade = (v) => String(+Number(v).toFixed(2)).replace(".", ",");
  const fmtPct = (v) => (v === null || v === undefined ? "—" : Number(v).toFixed(1).replace(".", ",") + " %");
  const show = (el, on = true) => el.classList.toggle("hidden", !on);

  /* ---------------------------------------------------------------- Alertes */
  function showAlert(msg) {
    const box = $("#global-alert");
    box.innerHTML = icon("alert") + `<div>${esc(msg)}</div>`;
    show(box);
    box.scrollIntoView({ behavior: "smooth", block: "center" });
  }
  const hideAlert = () => show($("#global-alert"), false);

  function setStep(n) {
    $$(".steps li").forEach((li) => {
      const i = Number(li.dataset.step);
      li.classList.toggle("done", i < n);
      li.classList.toggle("active", i === n);
    });
  }

  function overlay(on, text) {
    if (text) $("#overlay-text").textContent = text;
    show($("#overlay"), on);
  }

  async function api(url, body, method = "POST") {
    let resp;
    try {
      resp = await fetch(url, { method, headers: { "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    } catch (e) {
      throw new Error("Impossible de contacter l'application. Vérifiez qu'elle est toujours lancée.");
    }
    let json = {};
    try { json = await resp.json(); } catch (e) { /* réponse vide */ }
    if (!resp.ok || json.ok === false) throw new Error(json.error || "Une erreur est survenue.");
    return json;
  }

  /* ------------------------------------------------------------- Import */
  const dz = $("#dropzone");
  const fileInput = $("#file-input");
  const openDialog = () => fileInput.click();

  $("#btn-browse").addEventListener("click", (e) => { e.stopPropagation(); openDialog(); });
  $("#btn-replace").addEventListener("click", openDialog);
  dz.addEventListener("click", openDialog);
  dz.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openDialog(); } });
  ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
  dz.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) upload(f); });
  ["dragover", "drop"].forEach((ev) => window.addEventListener(ev, (e) => { if (!dz.contains(e.target)) e.preventDefault(); }));
  fileInput.addEventListener("change", () => { if (fileInput.files[0]) upload(fileInput.files[0]); fileInput.value = ""; });

  function upload(file) {
    hideAlert();
    const ext = (file.name.toLowerCase().match(/\.[^.]+$/) || [""])[0];
    if (![".xlsx", ".xls"].includes(ext)) return showAlert("Format non valide. Seuls les fichiers Excel (.xlsx, .xls) sont acceptés.");
    if (file.size > MAX_BYTES) return showAlert("Le fichier est trop volumineux (maximum 10 Mo).");
    if (file.size === 0) return showAlert("Le fichier est vide.");

    const fd = new FormData();
    fd.append("file", file);
    const xhr = new XMLHttpRequest();
    const wrap = $("#upload-progress"), bar = $("#progress-bar"), txt = $("#progress-text");
    show(wrap); bar.classList.remove("indeterminate"); bar.style.width = "0%"; txt.textContent = "Envoi du fichier…";
    xhr.open("POST", "/api/upload");
    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable) bar.style.width = Math.round((e.loaded / e.total) * 100) + "%";
    };
    xhr.upload.onload = () => { bar.classList.add("indeterminate"); txt.textContent = "Analyse du fichier en cours…"; };
    xhr.onerror = () => { show(wrap, false); showAlert("Impossible de contacter l'application. Vérifiez qu'elle est toujours lancée."); };
    xhr.onload = () => {
      show(wrap, false);
      let json = {};
      try { json = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }
      if (xhr.status !== 200 || !json.ok) return showAlert(json.error || "Échec de l'import du fichier.");
      if (state.token) fetch(`/api/file/${state.token}`, { method: "DELETE" });
      state.token = json.token;
      state.scale = "";
      state.mapping = json.mapping;
      $("#observations").value = "";
      show($("#card-result"), false);
      render(json);
      $("#card-preview").scrollIntoView({ behavior: "smooth", block: "start" });
    };
    xhr.send(fd);
  }

  $("#btn-remove").addEventListener("click", () => {
    if (state.token) fetch(`/api/file/${state.token}`, { method: "DELETE" });
    Object.assign(state, { token: null, data: null, mapping: null, scale: "", lastSuggestion: "" });
    ["#card-preview", "#card-exam", "#card-obs", "#card-generate", "#card-result", "#file-bar"].forEach((s) => show($(s), false));
    show(dz); hideAlert(); setStep(1);
    $$("[data-meta]").forEach((i) => { i.value = ""; delete i.dataset.touched; });
  });

  /* ------------------------------------------------------------- Rendu aperçu */
  function render(data) {
    state.data = data;
    show(dz, false); show($("#file-bar"));
    $("#file-name").textContent = data.filename;
    $("#file-sub").textContent = `Feuille « ${data.sheet} » · ${data.counts.rows} ligne(s) d'élèves détectée(s)` +
      (data.sheets.length > 1 ? ` · ${data.sheets.length} feuilles lues` : "");
    ["#card-preview", "#card-exam", "#card-obs", "#card-generate"].forEach((s) => show($(s)));
    renderIssues(data);
    fillMeta(data.metadata);
    renderMapping(data);
    renderStats(data);
    renderStudents(data);
    $("#btn-generate").disabled = !data.has_valid;
    setStep(data.has_valid ? 3 : 2);
  }

  function renderIssues(data) {
    const box = $("#issues");
    const icons = { error: "alert", warning: "alert", info: "info" };
    const cls = { error: "alert-error", warning: "alert-warn", info: "alert-info" };
    let html = data.issues.map((i) => `<div class="alert ${cls[i.level] || "alert-info"}">${icon(icons[i.level] || "info")}<div>${esc(i.message)}</div></div>`).join("");
    if (data.has_valid && !data.issues.some((i) => i.level === "error")) {
      html = `<div class="alert alert-ok">${icon("check")}<div><strong>${data.counts.valid} élève(s) avec une note valide</strong> seront pris en compte.</div></div>` + html;
    }
    box.innerHTML = html;
  }

  function fillMeta(meta) {
    $$("[data-meta]").forEach((inp) => {
      if (inp.dataset.touched) return; // ne pas écraser une saisie manuelle
      inp.value = meta[inp.dataset.meta] || "";
      inp.placeholder = "Non trouvé dans le fichier";
    });
  }
  $$("[data-meta]").forEach((i) => i.addEventListener("input", () => { i.dataset.touched = "1"; }));

  function colLabel(c) {
    const extra = c.header || (c.samples.length ? "ex. " + c.samples.join(", ") : "");
    return `${c.letter} — ${extra || "(vide)"}`;
  }

  function renderMapping(data) {
    const m = data.mapping;
    $("#map-sheet").innerHTML = data.sheets.map((s) => `<option ${s === data.sheet ? "selected" : ""}>${esc(s)}</option>`).join("");
    const opts = (sel, withNone) =>
      (withNone ? `<option value="">— Aucune —</option>` : `<option value="" disabled ${sel == null ? "selected" : ""}>— Choisir —</option>`) +
      data.columns.map((c) => `<option value="${c.index}" ${c.index === sel ? "selected" : ""}>${esc(colLabel(c))}</option>`).join("");
    $("#map-name").innerHTML = opts(m.name_cols[0], false);
    $("#map-grade").innerHTML = opts(m.grade_col, false);
    $("#map-abs").innerHTML = opts(m.absence_col, true);
    $("#map-first").value = m.first_row || "";
    $("#map-scale").value = state.scale || "";
    const flag = data.needs_mapping;
    show($("#mapping-flag"), flag);
    $("#mapping-panel").open = flag;
  }

  function currentMapping() {
    const prev = state.data.mapping.name_cols;
    const nameSel = $("#map-name").value;
    return {
      sheet: $("#map-sheet").value,
      name_cols: nameSel === "" ? prev : (prev.length > 1 && String(prev[0]) === nameSel ? prev : [Number(nameSel)]),
      grade_col: $("#map-grade").value === "" ? null : Number($("#map-grade").value),
      absence_col: $("#map-abs").value === "" ? "" : Number($("#map-abs").value),
      first_row: Number($("#map-first").value) || null,
    };
  }

  async function reanalyze(mapping, scale) {
    hideAlert();
    overlay(true, "Nouvelle analyse du fichier…");
    try {
      const json = await api("/api/analyze", { token: state.token, filename: state.data.filename, mapping, scale });
      state.mapping = json.mapping;
      state.scale = scale || "";
      render(json);
    } catch (e) { showAlert(e.message); } finally { overlay(false); }
  }

  $("#btn-apply-map").addEventListener("click", () => reanalyze(currentMapping(), $("#map-scale").value));
  $("#map-sheet").addEventListener("change", () => reanalyze({ sheet: $("#map-sheet").value }, $("#map-scale").value));

  /* ------------------------------------------------------------- Statistiques */
  function statCards(stats) {
    return [
      ["", stats.n, "Élèves évalués"],
      ["", fmt(stats.mean), "Moyenne de la classe /20"],
      ["green", fmtGrade(stats.max), "Note la plus haute"],
      ["orange", fmtGrade(stats.min), "Note la plus basse"],
      ["red", stats.support_count, "Élèves nécessitant un soutien"],
      ["green", fmtPct(stats.pass_rate), "Ont la moyenne (≥ 10)"],
    ].map(([c, v, l]) => `<div class="stat ${c}"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join("");
  }

  function levelsTable(stats) {
    const rows = stats.levels.map((l) =>
      `<tr><td class="name"><span class="dot" style="background:${l.color}"></span><strong>${esc(l.label)}</strong></td><td>${esc(l.interval)}</td><td>${l.count}</td><td>${fmtPct(l.percent)}</td><td>${fmt(l.mean)}</td></tr>`).join("");
    return `<thead><tr><th>Niveau de maîtrise</th><th>Intervalle de notes</th><th>Nombre d'élèves</th><th>Pourcentage</th><th>Moyenne du niveau</th></tr></thead>
      <tbody>${rows}<tr class="total"><td class="name">Total / moyenne générale</td><td>0 – 20</td><td>${stats.n}</td><td>100 %</td><td>${fmt(stats.mean)}</td></tr></tbody>`;
  }

  function renderStats(data) {
    const cards = $("#stat-cards"), wrap = $("#levels-preview-wrap");
    if (!data.stats) {
      cards.innerHTML = ""; show(wrap, false);
      return;
    }
    cards.innerHTML = statCards(data.stats);
    $("#levels-preview").innerHTML = levelsTable(data.stats);
    show(wrap);
  }

  const STATUS = { valid: ["Valide", "tag-valid"], absent: ["Absent", "tag-absent"], missing: ["Note manquante", "tag-missing"],
    invalid: ["Invalide", "tag-invalid"], duplicate: ["Doublon", "tag-duplicate"] };

  function renderStudents(data) {
    const only = $("#only-issues").checked;
    const list = data.students.filter((s) => !only || s.status !== "valid");
    $("#students-count").textContent = `(${data.counts.valid} valides sur ${data.counts.rows} lignes)`;
    const rows = list.map((s, i) => {
      const [label, cls] = STATUS[s.status] || [s.status, ""];
      return `<tr><td>${s.row}</td><td class="name" dir="auto">${esc(s.name) || "<em class='muted'>—</em>"}</td>
        <td>${s.status === "valid" ? fmtGrade(s.grade) : esc(s.raw_grade) || "—"}</td>
        <td><span class="tag ${cls}">${label}</span></td><td class="name small muted">${esc(s.message)}</td></tr>`;
    }).join("");
    $("#students-table").innerHTML = `<thead><tr><th>Ligne</th><th>Nom et prénom</th><th>Note</th><th>Statut</th><th>Remarque</th></tr></thead>
      <tbody>${rows || `<tr><td colspan="5" class="muted">Aucune ligne à afficher.</td></tr>`}</tbody>`;
  }
  $("#only-issues").addEventListener("change", () => state.data && renderStudents(state.data));

  /* ------------------------------------------------------------- Observations */
  const obs = $("#observations");
  const countObs = () => { $("#obs-count").textContent = obs.value.length; };
  obs.addEventListener("input", countObs);

  $("#btn-suggest").addEventListener("click", async () => {
    hideAlert();
    if (!state.data || !state.data.has_valid) return showAlert("Aucune note valide : impossible de suggérer une observation.");
    if (obs.value.trim() && obs.value.trim() !== state.lastSuggestion.trim() &&
        !confirm("Remplacer le texte actuel par l'observation suggérée ?")) return;
    overlay(true, "Rédaction de l'observation suggérée…");
    try {
      const json = await api("/api/suggest", { token: state.token, mapping: state.mapping, scale: state.scale });
      obs.value = json.text; state.lastSuggestion = json.text; countObs();
    } catch (e) { showAlert(e.message); } finally { overlay(false); }
  });

  /* ------------------------------------------------------------- Dates */
  const examDate = $("#exam-date"), corrDate = $("#correction-date");
  examDate.addEventListener("change", () => { corrDate.min = examDate.value || ""; });

  function validateDates() {
    const err = $("#date-error");
    [examDate, corrDate].forEach((i) => i.classList.remove("invalid"));
    let msg = "";
    if (!examDate.value) { msg = "Veuillez renseigner la date de l'examen."; examDate.classList.add("invalid"); }
    else if (!corrDate.value) { msg = "Veuillez renseigner la date de correction."; corrDate.classList.add("invalid"); }
    else if (corrDate.value < examDate.value) { msg = "La date de correction ne peut pas être antérieure à la date de l'examen."; corrDate.classList.add("invalid"); }
    err.textContent = msg; show(err, !!msg);
    if (msg) $("#card-exam").scrollIntoView({ behavior: "smooth", block: "center" });
    return !msg;
  }

  /* ------------------------------------------------------------- Génération */
  $("#btn-generate").addEventListener("click", () => generate(false));
  $("#btn-two-pages").addEventListener("click", () => generate(true));

  async function generate(allowMultipage) {
    hideAlert();
    const gerr = $("#generate-error"); show(gerr, false);
    if (!state.token || !state.data || !state.data.has_valid) return showAlert("Importez d'abord un fichier contenant des notes valides.");
    if (!validateDates()) return;
    const metadata = {};
    $$("[data-meta]").forEach((i) => { metadata[i.dataset.meta] = i.value.trim(); });
    overlay(true, "Analyse et génération du rapport…");
    try {
      const json = await api("/api/generate", {
        token: state.token, mapping: state.mapping, scale: state.scale, metadata,
        exam_date: examDate.value, correction_date: corrDate.value, observations: obs.value, allow_multipage: allowMultipage,
      });
      renderResult(json);
      setStep(4);
    } catch (e) { gerr.textContent = e.message; show(gerr); } finally { overlay(false); }
  }

  function renderResult(r) {
    show($("#card-result"));
    $("#result-sub").textContent = r.pages === 1
      ? `Rapport tenant sur 1 page A4 (mise en page : ${r.layout}).`
      : `Rapport de ${r.pages ?? "?"} pages (mise en page : ${r.layout}).`;
    $("#result-warnings").innerHTML = r.warnings.map((w) => `<div class="alert alert-warn">${icon("alert")}<div>${esc(w)}</div></div>`).join("");
    $("#result-cards").innerHTML = statCards(r.stats);
    $("#result-levels").innerHTML = levelsTable(r.stats);
    $("#dl-docx").href = r.docx_url;
    const pdf = $("#dl-pdf");
    pdf.classList.toggle("disabled", !r.pdf_url);
    pdf.href = r.pdf_url || "#";
    pdf.title = r.pdf_url ? "" : "PDF indisponible : LibreOffice est requis";
    $("#chart1").src = r.chart1_url; $("#chart2").src = r.chart2_url;
    const box = $("#two-page-box");
    show(box, !!r.can_offer_two_pages);
    $("#two-page-text").textContent = "Pour rester lisible, vous pouvez obtenir un rapport sur 2 pages au lieu d'un document comprimé.";
    $("#card-result").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  /* ------------------------------------------------------------- Démarrage */
  fetch("/api/status").then((r) => r.json()).then((s) => show($("#pdf-alert"), !s.pdf_available)).catch(() => {});
  countObs();
})();
