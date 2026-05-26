const Joi = require("joi");

const adminSettingsSchema = Joi.object({
  password: Joi.string().required(),
  meetingLink: Joi.string().uri().required(),
  failureLink: Joi.string().uri().required(),
});

module.exports = { adminSettingsSchema };
