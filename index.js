const express = require("express");
const ExcelJS = require("exceljs");
const fs = require("fs");
const telegramBot = require("node-telegram-bot-api");
const dotenv = require("dotenv");
dotenv.config();
const app = express();
const port = process.env.PORT || 3000;

const token = process.env.TOKEN;
const bot = new telegramBot(token, { polling: true });

const fileName = "shartnomalar.xlsx";
let workbook = new ExcelJS.Workbook();
let worksheet;
const fileNameContract = "contract.xlsx";
let workbookContract = new ExcelJS.Workbook();

// Fayl mavjud bo'lsa uni yuklaymiz, bo'lmasa yangi fayl yaratamiz
if (fs.existsSync(fileName)) {
  workbook.xlsx.readFile(fileName).then(() => {
    worksheet = workbook.getWorksheet("shartnoma");
    if (!worksheet) {
      worksheet = workbook.addWorksheet("shartnoma");
      initializeWorksheet();
    }
  });
} else {
  worksheet = workbook.addWorksheet("shartnoma");
  initializeWorksheet();
}

// Ish varag'ini boshlang'ich holatga keltiramiz
function initializeWorksheet() {
  worksheet.columns = [
    { header: "Tartib raqami", key: "id", width: 15 },
    { header: "Sana", key: "date", width: 20 },
    { header: "Bank nomi", key: "recipient", width: 30 },
    { header: "Kimga olgan", key: "contact", width: 30 },
    { header: "Telefon raqam", key: "phone", width: 40 },
  ];
}

// Qatorga yangi shartnomani qo'shish
function addNewRow({ recipient, contact, phone }) {
  const currentDate = new Date().toLocaleDateString();

  // Oxirgi qator raqamini aniqlash
  const lastRowNumber = worksheet.rowCount;
  const nextId = lastRowNumber ? lastRowNumber + 1 : 1;

  // Yangi qator qo'shamiz
  worksheet.addRow({
    id: nextId,
    date: currentDate,
    recipient: recipient,
    contact: contact,
    phone: phone,
  });

  // Excel faylini saqlaymiz
  workbook.xlsx.writeFile(fileName).then(() => {
    console.log(`Yangi shartnoma ${nextId}-qatorga qo'shildi.`);
  });
}

const inlineKeyboard = {
  reply_markup: {
    inline_keyboard: [
      [
        {
          text: "Hamkor bank",
          callback_data: "/hamkor",
        },
        {
          text: "Asaka bank",
          callback_data: "/asaka",
        },
      ],
    ],
  },
};

bot.setMyCommands([
  { command: "/start", description: "Bu bot sizga shartnomalar yuboradi" },
]);

let userPhoneNumber = null;
bot.on("contact", async (msg) => {
  const chatId = msg.chat.id;
  const contact = msg.contact;
  userPhoneNumber = msg.contact.phone_number;
  await bot.sendMessage(
    chatId,
    `Rahmat! Contact: ${contact.first_name} ${contact.phone_number}`
  );
  await bot.sendMessage(
    chatId,
    "Shartnoma qaysi bank uchun kerak?",
    inlineKeyboard
  );
});

bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const opts = {
    reply_markup: {
      keyboard: [
        [
          {
            text: "Kontaktingizni yuboring",
            request_contact: true,
          },
        ],
      ],
      resize_keyboard: true,
      one_time_keyboard: true,
      remove_keyboard: true,
    },
  };
  if (msg.text === "/start") {
    await bot.sendMessage(chatId, "Iltimos, kontaktingizni yuboring.", opts);
  }
});
bot.on("callback_query", async (query) => {
  const chatId = query.message.chat.id;
  if (query.data === "/hamkor" || query.data === "/asaka") {
    bot.sendMessage(chatId, "shartnoma yuborilmoqda...");
    const contractNumber = worksheet.rowCount;
    const today = new Date().toLocaleDateString();

    await workbookContract.xlsx
      .readFile(fileNameContract)
      .then(() => {
        let worksheetContract = workbookContract.getWorksheet("contract");

        if (!worksheetContract) {
          console.error("Worksheet 'contract' not found.");
          return;
        }
        const firstRow = worksheetContract.getRow(1);
        firstRow.getCell(1).value = `Hisob-varaq shartnoma ${contractNumber}`;
        firstRow.commit(); // Commit the changes

        const secondRow = worksheetContract.getRow(2);
        secondRow.getCell(
          1
        ).value = `${today} yil                                                                                                                Chelak shaxri`;
        secondRow.commit(); // Commit the changes
        if (query.data === "/asaka" || query.data === "/hamkor") {
          const thirdRow = worksheetContract.getRow(55);
          thirdRow.getCell(1).value =
            query.data === "/hamkor"
              ? `H.r: 20208000104817335001`
              : `H.r: 20208000504817335002`;
          thirdRow.commit(); // Commit the changes
          const fourthRow = worksheetContract.getRow(56);
          fourthRow.getCell(1).value =
            query.data === "/hamkor"
              ? `ChEKI AT “Hamkor bank”`
              : `"Asaka bank" AJ Bosh ofisi.`;
          fourthRow.commit(); // Commit the changes
          const fivethRow = worksheetContract.getRow(57);
          fivethRow.getCell(1).value =
            query.data === "/hamkor"
              ? `MFO: 00083   STIR: 301409058`
              : `MFO: 00873   STIR: 301409058`;
          fivethRow.commit(); // Commit the changes
        }

        return workbookContract.xlsx.writeFile(fileNameContract);
      })
      .catch((error) => {
        console.error("Error reading file:", error);
      });

    addNewRow({
      recipient: query.data === "/hamkor" ? "Hamkor bank" : "Asaka bank",
      contact: `${query.from.first_name} ${query.from.last_name}`,
      phone: userPhoneNumber,
    });
    await bot.sendDocument(chatId, "./contract.xlsx", {
      caption: `${today} yildagi ${contractNumber}-son shartnoma.`,
    });
    bot.sendMessage(chatId, "Shartnoma yuborildi.");
    await bot.sendMessage(
      chatId,
      "Shartnoma qaysi bank uchun kerak?",
      inlineKeyboard
    );
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
