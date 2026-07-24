// ============================================================
// GRAPH — calendar access (yours + any shared with you)
// ============================================================
// Important reminder: Graph never grants more access than what
// already exists on the Outlook side. A "read-only" delegated
// calendar stays read-only here — the agent cannot work around that.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

const MOCK_CALENDARS = [
  { id: "me", name: "My calendar", owner: "Aurelie Plisson", rights: "write" },
  { id: "mock-bjorn", name: "Bjorn Bender", owner: "Bjorn Bender", rights: "read" },
  { id: "mock-marc", name: "Marc (shared personal)", owner: "Marc Plisson", rights: "write" },
];

const MOCK_EVENTS = {
  me: [
    { id: "e1", title: "OKR FY26-27 check-in", start: "2026-07-24T09:00:00", end: "2026-07-24T09:30:00" },
    { id: "e2", title: "Project Orient follow-up", start: "2026-07-24T14:00:00", end: "2026-07-24T15:00:00" },
  ],
  "mock-bjorn": [
    { id: "e3", title: "Exec Team Meeting", start: "2026-07-24T10:00:00", end: "2026-07-24T11:00:00" },
  ],
  "mock-marc": [],
};

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error ${res.status}: ${body}`);
  }
  return res.status === 204 ? null : res.json();
}

// Lists accessible calendars: yours + any already shared/delegated.
async function listCalendars() {
  if (CONFIG.mockMode) return MOCK_CALENDARS;

  // /me/calendars returns your own calendars.
  // Colleagues' calendars that were delegated to you are fetched by
  // listing each known person (e.g. via /me/people, or a list you
  // maintain yourself), then calling /users/{id|upn}/calendar for
  // each one — Graph will refuse if the sharing wasn't done on the
  // Outlook side.
  const mine = await graphFetch("/me/calendars");
  return mine.value.map((c) => ({
    id: c.id,
    name: c.name,
    owner: "Me",
    rights: c.canEdit ? "write" : "read",
  }));
}

async function listEvents(calendarId, { startISO, endISO } = {}) {
  if (CONFIG.mockMode) return MOCK_EVENTS[calendarId] || [];

  const start = startISO || new Date().toISOString();
  const end = endISO || new Date(Date.now() + 7 * 24 * 3600 * 1000).toISOString();

  const isMine = calendarId === "me";
  const path = isMine
    ? `/me/calendarView?startDateTime=${start}&endDateTime=${end}`
    : `/users/${calendarId}/calendarView?startDateTime=${start}&endDateTime=${end}`;

  const data = await graphFetch(path);
  return data.value.map((e) => ({
    id: e.id,
    title: e.subject,
    start: e.start.dateTime,
    end: e.end.dateTime,
  }));
}

async function createEvent(calendarId, { title, startISO, endISO, attendees = [] }) {
  if (CONFIG.mockMode) {
    const newEvent = { id: `mock-${Date.now()}`, title, start: startISO, end: endISO };
    MOCK_EVENTS[calendarId] = MOCK_EVENTS[calendarId] || [];
    MOCK_EVENTS[calendarId].push(newEvent);
    return newEvent;
  }

  const isMine = calendarId === "me";
  const path = isMine ? "/me/events" : `/users/${calendarId}/events`;

  const body = {
    subject: title,
    start: { dateTime: startISO, timeZone: "Europe/Paris" },
    end: { dateTime: endISO, timeZone: "Europe/Paris" },
    attendees: attendees.map((a) => ({
      emailAddress: { address: a },
      type: "required",
    })),
  };

  return graphFetch(path, { method: "POST", body: JSON.stringify(body) });
}

async function updateEvent(calendarId, eventId, patch) {
  if (CONFIG.mockMode) {
    const list = MOCK_EVENTS[calendarId] || [];
    const ev = list.find((e) => e.id === eventId);
    if (ev) Object.assign(ev, patch);
    return ev;
  }

  const isMine = calendarId === "me";
  const path = isMine ? `/me/events/${eventId}` : `/users/${calendarId}/events/${eventId}`;
  return graphFetch(path, { method: "PATCH", body: JSON.stringify(patch) });
}
