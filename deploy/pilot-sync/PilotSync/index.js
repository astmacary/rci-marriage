const { run } = require("../scripts/pilot-sync.cjs");

module.exports = async function (context) {
  try {
    await run(message => context.log(message));
  } catch (error) {
    context.log.error("Pilot sync failed:", error.message);
    throw new Error("Pilot sync failed. Check the invocation log.");
  }
};
