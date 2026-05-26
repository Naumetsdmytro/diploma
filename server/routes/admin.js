const express = require("express");
const AppSettings = require("../models/AppSettings");
const { HttpError } = require("../helpers");
const { adminSettingsSchema } = require("../schemas/admin");

const router = express.Router();

function verifyAdminPassword(password) {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected) {
    throw HttpError(500, "ADMIN_PASSWORD is not configured on the server");
  }
  if (password !== expected) {
    throw HttpError(401, "Invalid admin password");
  }
}

router.get("/settings", async (req, res, next) => {
  try {
    const settings = await AppSettings.getSettings();
    res.json({
      meetingLink: settings.meetingLink,
      failureLink: settings.failureLink,
      updatedAt: settings.updatedAt,
    });
  } catch (error) {
    next(error);
  }
});

router.put("/settings", async (req, res, next) => {
  try {
    const { error, value } = adminSettingsSchema.validate(req.body);
    if (error) {
      throw HttpError(400, error.message);
    }

    verifyAdminPassword(value.password);

    const settings = await AppSettings.findOneAndUpdate(
      { key: "default" },
      {
        meetingLink: value.meetingLink,
        failureLink: value.failureLink,
      },
      { returnDocument: "after", upsert: true, setDefaultsOnInsert: true }
    );

    res.json({
      meetingLink: settings.meetingLink,
      failureLink: settings.failureLink,
      updatedAt: settings.updatedAt,
    });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
