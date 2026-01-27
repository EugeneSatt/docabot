import dotenv from "dotenv";
import { Bot, InputFile, Keyboard } from "grammy";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { readFile } from "node:fs/promises";
import path from "node:path";

dotenv.config();

type Session =
  | { step: "fio" }
  | { step: "date"; fio: string };

const sessions = new Map<number, Session>();

const token = process.env.BOT_TOKEN;
if (!token) {
  throw new Error("BOT_TOKEN is missing in .env");
}

const bot = new Bot(token);

const templatePath = path.resolve(process.cwd(), "templates", "Соглашение.docx");

const normalizeSpaces = (value: string) => value.replace(/\s+/g, " ").trim();

const isValidFio = (value: string) => {
  const normalized = normalizeSpaces(value);
  if (normalized.length === 0 || normalized.length > 120) return false;
  const words = normalized.split(" ");
  return words.length >= 2;
};

const sanitizeFilename = (value: string) => {
  const normalized = normalizeSpaces(value);
  const cleaned = normalized
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  const safe = cleaned.length > 0 ? cleaned : "agreement";
  return safe.slice(0, 100);
};

const renderDocx = async (data: { residentFio: string; agreementDate: string }) => {
  const templateBuffer = await readFile(templatePath);
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
  });
  const fullText = doc.getFullText();
  const missing = [];
  if (!fullText.includes("residentFio")) missing.push("residentFio");
  if (!fullText.includes("agreementDate")) missing.push("agreementDate");
  if (missing.length > 0) {
    throw new Error(`TEMPLATE_MISSING_TAGS:${missing.join(",")}`);
  }
  doc.setData(data);
  doc.render();
  return doc.getZip().generate({ type: "nodebuffer" });
};

const monthNames = [
  "января",
  "февраля",
  "марта",
  "апреля",
  "мая",
  "июня",
  "июля",
  "августа",
  "сентября",
  "октября",
  "ноября",
  "декабря",
];

const monthTokens: Record<string, number> = {
  январь: 1,
  января: 1,
  янв: 1,
  "янв.": 1,
  февраль: 2,
  февраля: 2,
  фев: 2,
  "фев.": 2,
  март: 3,
  марта: 3,
  мар: 3,
  "мар.": 3,
  апрель: 4,
  апреля: 4,
  апр: 4,
  "апр.": 4,
  май: 5,
  мая: 5,
  июнь: 6,
  июня: 6,
  июн: 6,
  "июн.": 6,
  июль: 7,
  июля: 7,
  июл: 7,
  "июл.": 7,
  август: 8,
  августа: 8,
  авг: 8,
  "авг.": 8,
  сентябрь: 9,
  сентября: 9,
  сент: 9,
  сен: 9,
  "сен.": 9,
  "сент.": 9,
  октябрь: 10,
  октября: 10,
  окт: 10,
  "окт.": 10,
  ноябрь: 11,
  ноября: 11,
  ноя: 11,
  "ноя.": 11,
  декабрь: 12,
  декабря: 12,
  дек: 12,
  "дек.": 12,
  january: 1,
  jan: 1,
  "jan.": 1,
  february: 2,
  feb: 2,
  "feb.": 2,
  march: 3,
  mar: 3,
  "mar.": 3,
  april: 4,
  apr: 4,
  "apr.": 4,
  may: 5,
  june: 6,
  jun: 6,
  "jun.": 6,
  july: 7,
  jul: 7,
  "jul.": 7,
  august: 8,
  aug: 8,
  "aug.": 8,
  september: 9,
  sept: 9,
  sep: 9,
  "sep.": 9,
  october: 10,
  oct: 10,
  "oct.": 10,
  november: 11,
  nov: 11,
  "nov.": 11,
  december: 12,
  dec: 12,
  "dec.": 12,
};

const formatRussianDate = (day: number, month: number, year: number) => {
  const safeMonth = Math.min(Math.max(month, 1), 12);
  const monthName = monthNames[safeMonth - 1];
  const dayStr = String(day).padStart(2, "0");
  return `«${dayStr}» ${monthName} ${year} г.`;
};

const isValidDate = (day: number, month: number, year: number) => {
  if (year < 1900 || year > 2100) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
};

const coerceYear = (raw: string) => {
  const value = Number(raw);
  if (raw.length === 4) return value;
  if (raw.length === 3) return 2000 + value;
  if (raw.length === 2) return value <= 50 ? 2000 + value : 1900 + value;
  return value;
};

const tryBuildDate = (day: number, month: number, yearRaw: string) => {
  const year = coerceYear(yearRaw);
  if (!isValidDate(day, month, year)) return null;
  return { day, month, year };
};

const normalizeDateInput = (input: string) =>
  normalizeSpaces(
    input
      .toLowerCase()
      .replace(/ё/g, "е")
      .replace(/[«»"]/g, " ")
      .replace(/\b\d{1,2}:\d{2}(:\d{2})?\b/gi, " ")
      .replace(/\b\d{1,2}:\d{2}(:\d{2})?\s*(am|pm)\b/gi, " ")
      .replace(/\b[+-]\d{2}:?\d{2}\b/g, " ")
      .replace(/t(?=\d{1,2}:\d{2})/gi, " ")
      .replace(/\b(г|г\.|год|года)\b/gi, " ")
      .replace(/[(),]/g, " ")
      .replace(/([a-zа-яё])(\d)/gi, "$1 $2")
      .replace(/(\d)([a-zа-яё])/gi, "$1 $2")
  );

const parseWithMonthName = (input: string) => {
  const tokens = input
    .split(/[\s/.\-:]+/)
    .map((token) => token.replace(/\./g, ""))
    .filter(Boolean);

  const monthIndex = tokens.findIndex((token) => monthTokens[token] !== undefined);
  if (monthIndex === -1) return null;

  const month = monthTokens[tokens[monthIndex]];
  let day: number | null = null;
  let year: number | null = null;

  for (let i = monthIndex - 1; i >= 0; i -= 1) {
    if (/^\d{1,2}$/.test(tokens[i])) {
      const candidate = Number(tokens[i]);
      if (candidate >= 1 && candidate <= 31) {
        day = candidate;
        break;
      }
    }
  }
  if (day === null) {
    for (let i = monthIndex + 1; i < tokens.length; i += 1) {
      if (/^\d{1,2}$/.test(tokens[i])) {
        const candidate = Number(tokens[i]);
        if (candidate >= 1 && candidate <= 31) {
          day = candidate;
          break;
        }
      }
    }
  }

  for (let i = monthIndex + 1; i < tokens.length; i += 1) {
    if (/^\d{2,4}$/.test(tokens[i])) {
      year = coerceYear(tokens[i]);
      break;
    }
  }
  if (year === null) {
    for (let i = monthIndex - 1; i >= 0; i -= 1) {
      if (/^\d{2,4}$/.test(tokens[i])) {
        year = coerceYear(tokens[i]);
        break;
      }
    }
  }

  if (day === null) return null;
  if (year === null) {
    year = getMoscowYmd(0).year;
  }

  return isValidDate(day, month, year) ? { day, month, year } : null;
};

const parseNumericDate = (input: string) => {
  const digitsOnly = input.replace(/\D/g, "");
  if (digitsOnly.length === 8) {
    const y1 = digitsOnly.slice(0, 4);
    const m1 = Number(digitsOnly.slice(4, 6));
    const d1 = Number(digitsOnly.slice(6, 8));
    const ymd = tryBuildDate(d1, m1, y1);
    if (ymd) return ymd;
    const d2 = Number(digitsOnly.slice(0, 2));
    const m2 = Number(digitsOnly.slice(2, 4));
    const y2 = digitsOnly.slice(4, 8);
    const dmy = tryBuildDate(d2, m2, y2);
    if (dmy) return dmy;
  }

  if (digitsOnly.length === 6) {
    const d1 = Number(digitsOnly.slice(0, 2));
    const m1 = Number(digitsOnly.slice(2, 4));
    const y1 = digitsOnly.slice(4, 6);
    const dmy = tryBuildDate(d1, m1, y1);
    if (dmy) return dmy;
    const y2 = digitsOnly.slice(0, 2);
    const m2 = Number(digitsOnly.slice(2, 4));
    const d2 = Number(digitsOnly.slice(4, 6));
    const ymd = tryBuildDate(d2, m2, y2);
    if (ymd) return ymd;
  }

  const matches = [...input.matchAll(/\d{1,4}/g)].map((match) => ({
    raw: match[0],
    value: Number(match[0]),
  }));

  for (let i = 0; i <= matches.length - 3; i += 1) {
    const a = matches[i];
    const b = matches[i + 1];
    const c = matches[i + 2];
    const candidates = [];

    if (a.raw.length === 4) {
      candidates.push(tryBuildDate(c.value, b.value, a.raw));
      candidates.push(tryBuildDate(b.value, c.value, a.raw));
    }
    if (c.raw.length === 4) {
      candidates.push(tryBuildDate(a.value, b.value, c.raw));
      candidates.push(tryBuildDate(b.value, a.value, c.raw));
    }
    if (b.raw.length === 4) {
      candidates.push(tryBuildDate(a.value, c.value, b.raw));
      candidates.push(tryBuildDate(c.value, a.value, b.raw));
    }

    if (c.raw.length <= 2) {
      candidates.push(tryBuildDate(a.value, b.value, c.raw));
      candidates.push(tryBuildDate(b.value, a.value, c.raw));
    }
    if (a.raw.length <= 2) {
      candidates.push(tryBuildDate(c.value, b.value, a.raw));
      candidates.push(tryBuildDate(b.value, c.value, a.raw));
    }

    const valid = candidates.find((item) => item);
    if (valid) return valid;
  }

  const two = matches.filter((item) => item.raw.length <= 2);
  if (two.length >= 2) {
    const year = getMoscowYmd(0).year;
    const dmy = isValidDate(two[0].value, two[1].value, year)
      ? { day: two[0].value, month: two[1].value, year }
      : null;
    if (dmy) return dmy;
    const mdy = isValidDate(two[1].value, two[0].value, year)
      ? { day: two[1].value, month: two[0].value, year }
      : null;
    if (mdy) return mdy;
  }

  return null;
};

const parseDateInput = (input: string) => {
  const cleaned = normalizeDateInput(input);
  if (!cleaned) return null;

  const withMonthName = parseWithMonthName(cleaned);
  if (withMonthName) return withMonthName;

  return parseNumericDate(cleaned);
};

const getMoscowYmd = (offsetDays = 0) => {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("ru-RU", {
    timeZone: "Europe/Moscow",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const year = Number(parts.find((p) => p.type === "year")?.value ?? now.getUTCFullYear());
  const month = Number(parts.find((p) => p.type === "month")?.value ?? now.getUTCMonth() + 1);
  const day = Number(parts.find((p) => p.type === "day")?.value ?? now.getUTCDate());
  const baseUtc = Date.UTC(year, month - 1, day);
  const shifted = new Date(baseUtc + offsetDays * 86400000);
  return {
    day: shifted.getUTCDate(),
    month: shifted.getUTCMonth() + 1,
    year: shifted.getUTCFullYear(),
  };
};

const dateKeyboard = new Keyboard()
  .text("Сегодня")
  .text("Вчера")
  .text("Завтра")
  .resized()
  .oneTime();

const askForFio = async (chatId: number) => {
  sessions.set(chatId, { step: "fio" });
  await bot.api.sendMessage(chatId, "Введите ФИО (минимум 2 слова):");
};

bot.command("start", async (ctx) => {
  const chatId = ctx.chat?.id;
  if (!chatId) return;
  await askForFio(chatId);
});

bot.on("message:text", async (ctx) => {
  const chatId = ctx.chat?.id;
  if (!chatId) return;

  const text = ctx.message.text;
  if (text.startsWith("/")) return;

  const session = sessions.get(chatId);
  if (!session) {
    await ctx.reply("Напишите /start чтобы создать документ.");
    return;
  }

  if (session.step === "fio") {
    const fio = normalizeSpaces(text);
    if (!isValidFio(fio)) {
      await ctx.reply("ФИО должно содержать минимум 2 слова и быть не длиннее 120 символов. Попробуйте снова:");
      return;
    }
    sessions.set(chatId, { step: "date", fio });
    await ctx.reply("Введите дату соглашения (в свободной форме):", {
      reply_markup: dateKeyboard,
    });
    return;
  }

  if (session.step === "date") {
    const normalized = normalizeSpaces(text);
    if (normalized.length === 0) {
      await ctx.reply("Дата не может быть пустой. Введите дату соглашения:", {
        reply_markup: dateKeyboard,
      });
      return;
    }

    let ymd: { day: number; month: number; year: number } | null = null;
    const lower = normalized.toLowerCase();
    if (lower === "сегодня") ymd = getMoscowYmd(0);
    if (lower === "вчера") ymd = getMoscowYmd(-1);
    if (lower === "завтра") ymd = getMoscowYmd(1);
    if (!ymd) ymd = parseDateInput(normalized);

    if (!ymd) {
      await ctx.reply(
        "Не удалось распознать дату. Примеры: 27.01.2026, 27 янв 2026, 27 января 2026.",
        { reply_markup: dateKeyboard }
      );
      return;
    }

    const agreementDate = formatRussianDate(ymd.day, ymd.month, ymd.year);

    try {
      const buffer = await renderDocx({
        residentFio: session.fio,
        agreementDate,
      });
      const filename = `agreement_${sanitizeFilename(session.fio)}.docx`;
      await ctx.replyWithDocument(new InputFile(buffer, filename), {
        reply_markup: { remove_keyboard: true },
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error("DOCX generation failed:", error);
      if (message.startsWith("TEMPLATE_MISSING_TAGS")) {
        await ctx.reply("В шаблоне нет полей {{residentFio}} и/или {{agreementDate}}. Обратитесь к администратору.", {
          reply_markup: { remove_keyboard: true },
        });
      } else {
        await ctx.reply("Не удалось создать документ. Попробуйте позже.", {
          reply_markup: { remove_keyboard: true },
        });
      }
    } finally {
      sessions.delete(chatId);
    }
  }
});

bot.catch((err) => {
  console.error("Bot error:", err.error);
});

await bot.start();
