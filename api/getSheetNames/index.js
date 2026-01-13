module.exports = async function (context, req) {
  // Temporary mock response (we will replace with Google Sheets API in later steps)
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: months,
  };
};
