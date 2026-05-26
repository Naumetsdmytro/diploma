const mongoose = require("mongoose");

const checkFieldSchema = {
  passed: { type: Boolean, default: false },
  checkedAt: { type: Date, default: null },
  device: { type: String, enum: ["desktop", "mobile", null], default: null },
};

const userSchema = new mongoose.Schema(
  {
    id: { type: String, required: true, unique: true, index: true },
    name: { type: String, required: true },
    googleName: { type: String, default: " " },
    loginCredential: { type: String, required: true },
    meetingLink: { type: String, required: true },
    failureLink: { type: String, required: true },
    camera: { type: Boolean, default: false },
    microphone: { type: Boolean, default: false },
    audio: { type: Boolean, default: false },
    isPossibleToUsePhone: { type: Boolean, default: true },
    techCheck: {
      camera: checkFieldSchema,
      microphone: checkFieldSchema,
      audio: checkFieldSchema,
    },
    overallStatus: {
      type: String,
      enum: ["in_progress", "passed", "failed"],
      default: "in_progress",
    },
  },
  { timestamps: true }
);

module.exports = mongoose.model("User", userSchema);
