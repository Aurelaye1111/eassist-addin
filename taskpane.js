// ============================================================
// TASKPANE — UI orchestration
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
  setStatus("Signing in…");

  try {
    await signIn();
    document.getElementById("signInView").classList.add("hidden");
    document.getElementById("mainView").classList.remove("hidden");
    setConnectionState(true);
    await populateCalendars();
    setStatus("Connected.");
  } catch (err) {
    errorEl.textContent = `Sign-in failed: ${err.message}`;
    errorEl.classList.remove("hidden");
    setStatus("Sign-in error.");
  }
}

function setConnectionState(on) {
  const el = document.getElementById("connectionState");
  el.classList.toggle("connection-state--on", on);
  el.classList.toggle("connection-state--off", !on);
  el.querySelector(".label").textContent = on ? "Connected" : "Not connected";
}

async function populateCalendars() {
  const select = document.getElementById("calendarSelect");
  select.innerHTML = "";

  const calendars = await listCalendars();
  for (const cal of calendars) {
    const opt = document.createElement("option");
    opt.value = cal.id;
    const rightsLabel = cal.rights === "write" ? "read/write" : "read only";
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
  list.innerHTML = `<li class="events-empty">Loading…</li>`;

  try {
    const events = await listEvents(calendarId);
    if (events.length === 0) {
      list.innerHTML = `<li class="events-empty">No events in this period.</li>`;
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
    list.innerHTML = `<li class="events-empty">Loading error: ${escapeHtml(err.message)}</li>`;
  }
}

async function handleAsk() {
  const prompt = document.getElementById("promptInput").value.trim();
  if (!prompt) return;

  setStatus("Analyzing request…");

  // -----------------------------------------------------------------
  // ⚠️ ARCHITECTURE NOTE ⚠️
  // Real natural-language parsing (e.g. "find a slot with Marc on
  // Thursday afternoon") needs to be done by an LLM. We CANNOT call
  // the Claude API directly from this file: that would expose the API
  // key in the source code of this static GitHub Pages site, visible
  // to anyone via the browser inspector.
  //
  // A small backend proxy is needed (Azure Function / Cloudflare
  // Worker / AWS Lambda) that receives the prompt, calls the Claude
  // API server-side with the key kept secret, and returns structured
  // JSON here.
  //
  // Until that backend exists, we use a very basic stub so we can
  // test the end-to-end UX.
  // -----------------------------------------------------------------
  const proposals = await stubParsePrompt(prompt);

  renderProposals(proposals);
  setStatus("Proposed actions ready — your call to approve.");
}

async function stubParsePrompt(prompt) {
  // Demo stub: no real AI here, just enough to visualize the flow
  // proposals -> confirmation -> write.
  const calendarId = document.getElementById("calendarSelect").value;
  const now = new Date();
  const start = new Date(now.getTime() + 24 * 3600 * 1000);
  start.setHours(14, 0, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);

  return [
    {
      type: "create",
      calendarId,
      title: `To define from: "${prompt.slice(0, 60)}"`,
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
      <div class="p-title">Create: ${escapeHtml(p.title)}</div>
      <div class="p-meta">${formatRange(p.startISO, p.endISO)}</div>
      <div class="proposal-actions">
        <button class="btn btn--confirm" data-idx="${idx}">Approve</button>
        <button class="btn btn--reject" data-idx="${idx}">Dismiss</button>
      </div>
    `;
    list.appendChild(li);
  });

  list.querySelectorAll(".btn--confirm").forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      const idx = Number(e.target.dataset.idx);
      const p = proposals[idx];
      setStatus("Writing to calendar…");
      await createEvent(p.calendarId, { title: p.title, startISO: p.startISO, endISO: p.endISO });
      await renderEvents(p.calendarId);
      e.target.closest(".proposal-card").remove();
      setStatus("Event created.");
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
  const dateFmt = new Intl.DateTimeFormat("en-US", { weekday: "short", day: "2-digit", month: "short" });
  const timeFmt = new Intl.DateTimeFormat("en-US", { hour: "2-digit", minute: "2-digit" });
  return `${dateFmt.format(start)} · ${timeFmt.format(start)}–${timeFmt.format(end)}`;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
