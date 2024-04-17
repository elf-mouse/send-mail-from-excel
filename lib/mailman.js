const nodemailer = require("nodemailer");
const splitSheets = require("./split-sheets");
const config = require("./config");

function sendMail() {
  // 创建一个传输器对象
  const transporter = nodemailer.createTransport(config.smtpOptions);

  // verify connection configuration
  transporter.verify(function (error, success) {
    if (error) {
      console.log(error);
    } else {
      console.log("Server is ready to take our messages");
      splitSheets(transporter, config.excel);
    }
  });
}

module.exports = sendMail;
