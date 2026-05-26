const Joi = require("joi");

const setDataSchema = Joi.object({
  spreadSheetId: Joi.string().required(),
  data: Joi.array().items(Joi.string().required(), Joi.number()).required(),
  sheetName: Joi.string().required(),
});

module.exports = {
  setDataSchema,
};
