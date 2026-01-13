module.exports = async function (context, req) {
  // Temporary mock response (we replace with real Google Sheets counts later)
  const body = [
    { name: "Jan", count: 0 },
    { name: "Feb", count: 0 },
    { name: "Mar", count: 0 },
    { name: "Apr", count: 0 },
    { name: "May", count: 0 },
    { name: "Jun", count: 0 },
    { name: "Jul", count: 0 },
    { name: "Aug", count: 0 },
    { name: "Sep", count: 10 },
    { name: "Oct", count: 5 },
    { name: "Nov", count: 2 },
    { name: "Dec", count: 0 },
  ];

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body,
  };
};
