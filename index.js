const express = require("express");
const telegramBot = require("node-telegram-bot-api");
const ExcelJS = require("exceljs");
const fs = require("fs");
const { pool } = require("./db");
const dotenv = require("dotenv");
dotenv.config();
const app = express();
const port = process.env.PORT || 3000;

const token = process.env.BOT_TOKEN;
const bot = new telegramBot(token, { polling: true });

let worksheet;
const fileNameContract = "contract.xlsx";
let workbookContract = new ExcelJS.Workbook();

function logToFile(message) {
  const timestamp = new Date().toISOString();
  fs.appendFile("bot.log", `[${timestamp}] ${message}\n`, (err) => {
    if (err) console.error("Log yozishda xatolik:", err);
  });
}
// Fayl mavjud bo'lsa uni yuklaymiz, bo'lmasa yangi fayl yaratamiz

const inlineKeyboard = {
  reply_markup: {
    inline_keyboard: [
      [
        {
          text: "Hamkor bank",
          callback_data: "hamkor",
        },
        {
          text: "Asaka bank",
          callback_data: "asaka",
        },
      ],
    ],
  },
};

bot.setMyCommands([
  { command: "/start", description: "Bu bot sizga shartnomalar yuboradi" },
]);

bot.on("contact", async (msg) => {
  const chatId = msg.chat.id;
  const contact = msg.contact;
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
  let lastRow;
  if (query.data === "hamkor" || query.data === "asaka") {
    await bot.sendMessage(chatId, "shartnoma yuborilmoqda...");

    pool.query(
      `INSERT INTO contracts (date, bank_name, contact) VALUES ($1, $2, $3)`,
      [
        new Date().toLocaleDateString("uz-UZ"),
        query.data === "hamkor" ? "Hamkor bank" : "Asaka bank",
        `${query.from.first_name}`,
      ]
    );
    // DB dagi oxirgi qatorni olish
    pool.query(
      "SELECT * FROM contracts ORDER BY id DESC LIMIT 1",
      (err, res) => {
        if (err) {
          console.error("Xatolik yuz berdi:", err);
          return;
        }
        const { id } = res.rows[0];
        lastRow = id;
      }
    );
    const today = new Date().toLocaleDateString("uz-UZ");
    console.log(lastRow);

    await workbookContract.xlsx
      .readFile(fileNameContract)
      .then(() => {
        let worksheetContract = workbookContract.getWorksheet("contract");

        if (!worksheetContract) {
          console.error("Worksheet 'contract' not found.");
          return;
        }
        const firstRow = worksheetContract.getRow(1);
        firstRow.getCell(1).value = `Hisob-varaq shartnoma ${lastRow}`;
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
        logToFile(`Error reading file: ${error}`);
        console.error("Error reading file:", error);
      });

    await bot.sendDocument(chatId, "./contract.xlsx", {
      caption: `${today} yildagi ${lastRow}-son shartnoma.`,
    });
    await bot.sendMessage(chatId, "Shartnoma yuborildi.");
    await bot.sendMessage(
      chatId,
      "Yangi shartnoma qaysi bank uchun kerak?",
      inlineKeyboard
    );
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
