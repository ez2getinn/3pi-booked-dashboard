module.exports = async function (context, req) {
  const sheet = (req.query.sheet || "Sep").toString();

  // These headers MUST match your frontend order:
  // Email, Name, Date, Start Time, End Time, Booked, Site, Account, Ticket, Shift4 MID
  const headers = [
    "Email",
    "Name",
    "Date",
    "Start Time",
    "End Time",
    "Booked",
    "Site",
    "Account",
    "Ticket",
    "Shift4 MID",
  ];

  // ✅ Mock rows for now (we will connect Microsoft Excel later)
  const rows = [
    [
      "tech1@shift4.com",
      "John Doe",
      "1/10/2026",
      "8:00 AM",
      "5:00 PM",
      "BOOKED",
      "Las Vegas",
      "Shift4",
      "000123",
      "009900",
    ],
    [
      "tech2@shift4.com",
      "Jane Smith",
      "1/12/2026",
      "9:00 AM",
      "6:00 PM",
      "BOOKED",
      "New York",
      "Shift4",
      "000456",
      "008800",
    ],
  ];

  // IMPORTANT: frontend expects ms aligned with each row
  const ms = rows.map((r) => {
    // date string is at index 2
    const d = new Date(r[2]);
    const dateMs = isNaN(d.getTime()) ? null : d.getTime();

    // fake start/end on same day
    const start = new Date(d);
    start.setHours(8, 0, 0, 0);

    const end = new Date(d);
    end.setHours(17, 0, 0, 0);

    return {
      dateMs: dateMs,
      startMs: isNaN(start.getTime()) ? null : start.getTime(),
      endMs: isNaN(end.getTime()) ? null : end.getTime(),
    };
  });

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: { sheet, headers, rows, ms },
  };
};
