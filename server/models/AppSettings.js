const mongoose = require("mongoose");

const appSettingsSchema = new mongoose.Schema(
  {
    key: { type: String, default: "default", unique: true },
    meetingLink: {
      type: String,
      default: "https://meet.google.com/",
    },
    failureLink: {
      type: String,
      default: "https://meet.google.com/",
    },
  },
  { timestamps: true }
);

appSettingsSchema.statics.getSettings = async function () {
  let settings = await this.findOne({ key: "default" });
  if (!settings) {
    settings = await this.create({
      key: "default",
      meetingLink:
        process.env.DEFAULT_MEETING_LINK || "https://meet.google.com/",
      failureLink:
        process.env.DEFAULT_FAILURE_LINK || "https://meet.google.com/",
    });
  }
  return settings;
};

module.exports = mongoose.model("AppSettings", appSettingsSchema);
