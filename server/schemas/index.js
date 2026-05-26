const { createSchema, updateSchema } = require("./users");
const { adminSettingsSchema } = require("./admin");

module.exports = {
  createSchema,
  updateSchema,
  adminSettingsSchema,
};
