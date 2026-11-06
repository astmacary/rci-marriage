const { run } = require("../scripts/couple-onboarding.cjs");

module.exports = async function (context) {
  try {
    await run(message => context.log(message));
  } catch (error) {
    context.log.error("Couple onboarding failed:", error.message);
    throw new Error("Couple onboarding failed. Check the invocation log.");
  }
};
