const Joi = require("joi");

const createSchema = Joi.object({
  id: Joi.string().required(),
  loginCredential: Joi.string().required(),
  name: Joi.string().required(),
  googleName: Joi.string().required(),
});

const updateSchema = Joi.object({
  camera: Joi.boolean(),
  microphone: Joi.boolean(),
  audio: Joi.boolean(),
  meetingLink: Joi.string().uri(),
  isPossibleToUsePhone: Joi.boolean(),
  overallStatus: Joi.string().valid("in_progress", "passed", "failed"),
}).min(1);

module.exports = {
  updateSchema,
  createSchema,
};
