const Joi = require("joi");

const createSchema = Joi.object({
  id: Joi.string().required(),
  loginCredential: Joi.string().required(),
  name: Joi.string().required(),
  googleName: Joi.string().required(),
  meetingLink: Joi.string().required(),
  mainRoomNumber: Joi.number().required(),
});

const updateSchema = Joi.object({
  camera: Joi.boolean().required(),
  microphone: Joi.boolean().required(),
  audio: Joi.boolean().required(),
  meetingLink: Joi.string().required(),
  isPossibleToUsePhone: Joi.boolean(),
});

module.exports = {
  updateSchema,
  createSchema,
};
