// ============================================================
// GRAPH — accès calendriers (le tien + ceux partagés avec toi)
// ============================================================
// Rappel important : Graph ne donne jamais plus d'accès que ce
// qui existe déjà côté Outlook. Un calendrier délégué "lecture seule"
// reste lecture seule ici — l'agent ne peut pas contourner ça.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

const MOCK_CALENDARS = [
  { id: "me", name: "Mon calendrier", owner: "Aurélie Plisson", rights: "write" },
  { id: "mock-bjorn", name: "Björn Bender", owner: "Björn Bender", rights: "read" },
  { id: "mock-marc", name: "Marc (perso partagé)", owner: "Marc Plisson", rights: "write" },
];

const MOCK_EVENTS = {
  me: [
    { id: "e1", title: "Point OKR FY26-27", start: "2026-07-24T09:00:00", end: "2026-07-24T09:30:00" },
    { id: "e2", title: "Suivi Project Orient", start: "2026-07-24T14:00:00", end: "2026-07-24T15:00:00" },
  ],
  "mock-bjorn": [
    { id: "e3", title: "CODIR", start: "2026-07-24T10:00:00", end: "2026-07-24T11:00:00" },
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

// Liste les calendriers accessibles : le tien + ceux déjà partagés/délégués.
async function listCalendars() {
  if (CONFIG.mockMode) return MOCK_CALENDARS;

  // /me/calendars renvoie tes propres calendriers.
  // Les calendriers de collègues qui t'ont donné un accès délégué
  // s'obtiennent en listant chaque personne connue (ex: via /me/people
  // ou une liste que tu maintiens toi-même), puis en appelant
  // /users/{id|upn}/calendar pour chacune — Graph refusera si le
  // partage n'a pas été fait côté Outlook.
  const mine = await graphFetch("/me/calendars");
  return mine.value.map((c) => ({
    id: c.id,
    name: c.name,
    owner: "Moi",
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
