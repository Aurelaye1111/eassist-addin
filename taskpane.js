// ============================================================
// TASKPANE — orchestration de l'interface
// ============================================================

Office.onReady(() => {
  wireUpUI();
});

function wireUpUI() {
  document.getElementById("signInBtn").addEventListener("click", handleSignIn);
  document.getElementById("calendarSelect").addEventListener("change", handleCalendarChange);
  document.getElementById("askBtn").addEventListener("click", handleAsk);
  document.getElementById("refreshBtn").addEventListener("click", () => {
    const calendarId = document.getElementById("calendarSelect").value;
    if (calendarId) renderEvents(calendarId);
  });
}

async function handleSignIn() {
  const errorEl = document.getElementById("signInError");
  errorEl.classList.add("hidden");
  setStatus("Connexion en cours…");

  try {
    await signIn();
    document.getElementById("signInView").classList.add("hidden");
    document.getElementById("mainView").classList.remove("hidden");
    setConnectionState(true);
    await populateCalendars();
    setStatus("Connecté.");
  } catch (err) {
    errorEl.textContent = `Échec de connexion : ${err.message}`;
    errorEl.classList.remove("hidden");
    setStatus("Erreur de connexion.");
  }
}

function setConnectionState(on) {
  const el = document.getElementById("connectionState");
  el.classList.toggle("connection-state--on", on);
  el.classList.toggle("connection-state--off", !on);
  el.querySelector(".label").textContent = on ? "Connecté" : "Non connecté";
}

async function populateCalendars() {
  const select = document.getElementById("calendarSelect");
  select.innerHTML = "";

  const calendars = await listCalendars();
  for (const cal of calendars) {
    const opt = document.createElement("option");
    opt.value = cal.id;
    const rightsLabel = cal.rights === "write" ? "lecture/écriture" : "lecture seule";
    opt.textContent = `${cal.name} (${rightsLabel})`;
    select.appendChild(opt);
  }

  if (calendars.length > 0) {
    select.value = calendars[0].id;
    await renderEvents(calendars[0].id);
  }
}

async function handleCalendarChange(e) {
  await renderEvents(e.target.value);
}

async function renderEvents(calendarId) {
  const list = document.getElementById("eventsList");
  list.innerHTML = `<li class="events-empty">Chargement…</li>`;

  try {
    const events = await listEvents(calendarId);
    if (events.length === 0) {
      list.innerHTML = `<li class="events-empty">Aucun événement sur cette période.</li>`;
      return;
    }
    list.innerHTML = "";
    for (const ev of events) {
      const li = document.createElement("li");
      li.className = "event-card";
      li.innerHTML = `
        <div class="e-title">${escapeHtml(ev.title)}</div>
        <div class="e-time">${formatRange(ev.start, ev.end)}</div>
      `;
      list.appendChild(li);
    }
  } catch (err) {
    list.innerHTML = `<li class="events-empty">Erreur de chargement : ${escapeHtml(err.message)}</li>`;
  }
}

async function handleAsk() {
  const prompt = document.getElementById("promptInput").value.trim();
  if (!prompt) return;

  setStatus("Analyse de la demande…");

  // -----------------------------------------------------------------
  // ⚠️ POINT D'ATTENTION ARCHITECTURE ⚠️
  // Le vrai parsing en langage naturel (ex: "trouve un créneau avec Marc
  // jeudi après-midi") doit être fait par un LLM. On NE PEUT PAS appeler
  // l'API Claude directement depuis ce fichier : ça exposerait la clé API
  // dans le code source du site statique GitHub Pages, visible par
  // n'importe qui via l'inspecteur du navigateur.
  //
  // Il faut un petit proxy backend (Azure Function / Cloudflare Worker /
  // AWS Lambda) qui reçoit le prompt, appelle l'API Claude côté serveur
  // avec la clé gardée secrète, et renvoie le JSON structuré ici.
  //
  // En attendant ce backend, on utilise un stub très basique pour
  // pouvoir tester l'UX de bout en bout.
  // -----------------------------------------------------------------
  const proposals = await stubParsePrompt(prompt);

  renderProposals(proposals);
  setStatus("Propose des actions — à toi de valider.");
}

async function stubParsePrompt(prompt) {
  // Stub de démo : pas d'IA réelle ici, juste pour visualiser le flux
  // propositions -> confirmation -> écriture.
  const calendarId = document.getElementById("calendarSelect").value;
  const now = new Date();
  const start = new Date(now.getTime() + 24 * 3600 * 1000);
  start.setHours(14, 0, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);

  return [
    {
      type: "create",
      calendarId,
      title: `À définir depuis : "${prompt.slice(0, 60)}"`,
      startISO: start.toISOString(),
      endISO: end.toISOString(),
    },
  ];
}

function renderProposals(proposals) {
  const section = document.getElementById("proposalsSection");
  const list = document.getElementById("proposalsList");
  list.innerHTML = "";

  if (proposals.length === 0) {
    section.classList.add("hidden");
    return;
  }

  section.classList.remove("hidden");

  proposals.forEach((p, idx) => {
    const li = document.createElement("li");
    li.className = "proposal-card";
    li.innerHTML = `
      <div class="p-title">Créer : ${escapeHtml(p.title)}</div>
      <div class="p-meta">${formatRange(p.startISO, p.endISO)}</div>
      <div class="proposal-actions">
        <button class="btn btn--confirm" data-idx="${idx}">Valider</button>
        <button class="btn btn--reject" data-idx="${idx}">Ignorer</button>
      </div>
    `;
    list.appendChild(li);
  });

  list.querySelectorAll(".btn--confirm").forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      const idx = Number(e.target.dataset.idx);
      const p = proposals[idx];
      setStatus("Écriture sur le calendrier…");
      await createEvent(p.calendarId, { title: p.title, startISO: p.startISO, endISO: p.endISO });
      await renderEvents(p.calendarId);
      e.target.closest(".proposal-card").remove();
      setStatus("Événement créé.");
    });
  });

  list.querySelectorAll(".btn--reject").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.target.closest(".proposal-card").remove();
    });
  });
}

function setStatus(text) {
  document.getElementById("statusLine").textContent = text;
}

function formatRange(startISO, endISO) {
  const start = new Date(startISO);
  const end = new Date(endISO);
  const dateFmt = new Intl.DateTimeFormat("fr-FR", { weekday: "short", day: "2-digit", month: "short" });
  const timeFmt = new Intl.DateTimeFormat("fr-FR", { hour: "2-digit", minute: "2-digit" });
  return `${dateFmt.format(start)} · ${timeFmt.format(start)}–${timeFmt.format(end)}`;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
