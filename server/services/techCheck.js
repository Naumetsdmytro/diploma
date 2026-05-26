const User = require("../models/User");

async function recordTechCheck(userId, step, { passed, device = null }) {
  const prefix = `techCheck.${step}`;
  const update = {
    [`${prefix}.passed`]: passed,
    [`${prefix}.checkedAt`]: new Date(),
    [`${prefix}.device`]: device,
    [step]: passed,
  };

  if (!passed) {
    update.overallStatus = "failed";
  }

  return User.findOneAndUpdate(
    { id: userId },
    { $set: update },
    { returnDocument: "after" }
  );
}

async function markTechCheckPassed(userId) {
  return User.findOneAndUpdate(
    { id: userId },
    {
      $set: {
        camera: true,
        microphone: true,
        audio: true,
        overallStatus: "passed",
        "techCheck.camera.passed": true,
        "techCheck.microphone.passed": true,
        "techCheck.audio.passed": true,
        "techCheck.audio.checkedAt": new Date(),
      },
    },
    { returnDocument: "after" }
  );
}

module.exports = { recordTechCheck, markTechCheckPassed };
