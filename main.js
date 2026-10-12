const ExcelJS = require("exceljs");

const URL_BASE = "https://dchart-api.vndirect.com.vn/dchart/history";
const DATA_FILE_NAME = "./data.xlsx";

const TICKERS = [
  "ACB",
  "BID",
  "CTG",
  "HDB",
  "LPB",
  "MBB",
  "SHB",
  "SSB",
  "STB",
  "TCB",
  "VCB",
  "VIB",
  "VPB",
  "VIC",
  "VHM",
  "VRE",
  "VPL",
  "FPT",
  "MSN",
  "MCH",
  "MWG",
  "VNM",
  "SAB",
  "GAS",
  "BSR",
  "POW",
  "HPG",
  "GVR",
  "SSI",
  "TCX",
  "BVH",
];

const YEARS = [2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026];

const HEADERS = {
  headers: {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
    Referer: "https://dchart.vndirect.com.vn/",
  },
};

const crawl = async (ticker, year) => {
  try {
    const fromTimestamp = Math.floor(new Date(`${year}-01-01`).getTime() / 1000);
    const toTimestamp = Math.floor(new Date(`${year}-12-31`).getTime() / 1000);

    const url = `${URL_BASE}?resolution=D&symbol=${ticker.toUpperCase()}&from=${fromTimestamp}&to=${toTimestamp}`;

    const res = await (await fetch(url, HEADERS)).json();

    if (!res || !res.s === "ok") return {};

    const dates = res.t.map((seconds) => new Date(seconds * 1000).toLocaleDateString("vi-VN"));
    return { dates, prices: [...res.c] };
  } catch (error) {
    console.error(`ERROR at ${ticker}: ${error.message}`);
    return {};
  }
};

const crawlAll = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DATA_FILE_NAME);
  workbook.calcProperties.fullCalcOnLoad = true;

  for (const ticker of TICKERS) {
    const sheet = workbook.getWorksheet(ticker);

    let lineOrder = 1;
    for (const year of YEARS) {
      const { dates, prices } = await crawl(ticker, year);

      sheet.getRow(lineOrder).values = [year, ...dates];
      sheet.getRow(lineOrder + 1).values = ["", ...prices];

      lineOrder += 2;
    }
  }

  await workbook.xlsx.writeFile(DATA_FILE_NAME);
  console.log(`Saved to ${DATA_FILE_NAME}!`);
};

crawlAll();
