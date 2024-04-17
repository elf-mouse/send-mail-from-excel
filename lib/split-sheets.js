const excel = require("exceljs");
const config = require("./config");

async function splitSheets(server, inputWorkbookPath) {
  // 创建一个工作簿对象
  const workbook = new excel.Workbook();

  // 读取输入工作簿
  await workbook.xlsx.readFile(inputWorkbookPath);

  // 循环每个工作表
  workbook.eachSheet(async (worksheet) => {
    // 创建一个新工作簿并将其重命名为当前工作表名称
    const newWorkbook = new excel.Workbook();
    newWorkbook.addWorksheet(worksheet.name);

    // 将当前工作表复制到新工作簿
    worksheet.eachRow((row) =>
      newWorkbook.getWorksheet(worksheet.name).addRow(row.values)
    );

    // 保存新工作簿
    const newFilePath = `data/${worksheet.name}.xlsx`;
    await newWorkbook.xlsx.writeFile(newFilePath);
    console.log(`${worksheet.name}.xlsx generated`);

    const data = newWorkbook.getWorksheet(worksheet.name);
    const [_, email, title, content] = data.getRow(2).values; // 获取内容可能需要根据实际情况进行微调

    // 创建一封邮件
    const mailOptions = {
      from: config.email,
      to:
        email !== null && typeof email === "object"
          ? email.hyperlink.replace("mailto:", "")
          : email,
      subject: title,
      text: content,
      attachments: {
        path: newFilePath,
      },
    };

    // 发送邮件
    if (!config.debug) {
      await server.sendMail(mailOptions);
      console.log(`${worksheet.name}.xlsx sended`);
    }
  });
}

module.exports = splitSheets;
