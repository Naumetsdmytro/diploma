const express = require("express");
const AppSettings = require("../models/AppSettings");
const User = require("../models/User");
const { createSchema, updateSchema } = require("../schemas");
const { HttpError } = require("../helpers");
const { markTechCheckPassed } = require("../services/techCheck");

const router = express.Router();

router.get("/", async (req, res, next) => {
  try {
    const users = await User.find().sort({ createdAt: -1 }).limit(100);
    res.json(users);
  } catch (error) {
    next(error);
  }
});

router.get("/:id", async (req, res, next) => {
  try {
    const user = await User.findOne({ id: req.params.id });
    if (!user) {
      throw HttpError(404, "Not found");
    }
    res.status(200).json(user);
  } catch (error) {
    next(error);
  }
});

router.post("/", async (req, res, next) => {
  try {
    const { error, value } = createSchema.validate(req.body);
    if (error) {
      throw HttpError(400, error.message);
    }

    const settings = await AppSettings.getSettings();
    const existing = await User.findOne({ id: value.id });
    if (existing) {
      res.status(200).json(existing);
      return;
    }

    const user = await User.create({
      ...value,
      meetingLink: settings.meetingLink,
      failureLink: settings.failureLink,
    });

    console.log("[POST /users] created user id:", user.id);
    res.status(201).json(user);
  } catch (error) {
    next(error);
  }
});

router.put("/:id", async (req, res, next) => {
  try {
    const { id } = req.params;
    const { error, value } = updateSchema.validate(req.body);
    if (error) {
      throw HttpError(400, error.message);
    }

    const user = await User.findOne({ id });
    if (!user) {
      throw HttpError(404, "Not found");
    }

    if (value.audio === true && value.camera === true && value.microphone === true) {
      const updated = await markTechCheckPassed(id);
      res.status(200).json(updated);
      return;
    }

    const techCheckUpdate = {};
    if (value.camera === true) {
      techCheckUpdate["techCheck.camera"] = {
        passed: true,
        checkedAt: new Date(),
        device: "desktop",
      };
    }
    if (value.microphone === true) {
      techCheckUpdate["techCheck.microphone"] = {
        passed: true,
        checkedAt: new Date(),
        device: "desktop",
      };
    }
    if (value.audio === true) {
      techCheckUpdate["techCheck.audio"] = {
        passed: true,
        checkedAt: new Date(),
        device: "desktop",
      };
    }

    const updated = await User.findOneAndUpdate(
      { id },
      { $set: { ...value, ...techCheckUpdate } },
      { returnDocument: "after" }
    );

    res.status(200).json(updated);
  } catch (error) {
    next(error);
  }
});

module.exports = router;
