const email = "";
const appPass = "";

const smtpOptions = {
  service: "Outlook365",
  host: "smtp-mail.outlook.com",
  port: 587,
  secure: false,
  auth: {
    user: email,
    pass: appPass,
  },
};

const excel = "data/source.xlsx";

const debug = false;

module.exports = {
  email,
  appPass,
  smtpOptions,
  excel,
  debug,
};
