module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  };
};
