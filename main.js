const ExcelJS = require("exceljs");

const URL_BASE = "https://dchart-api.vndirect.com.vn/dchart/history";
const FILE_PATH = "./data.xlsx";

const HEADERS = {
  headers: {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
    Referer: "https://dchart.vndirect.com.vn/",
  },
};

const tickers = [
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

const years = [2019, 2026];

const crawl = async (ticker, yearStart, yearEnd = yearStart) => {
  try {
    const fromTimestamp = Math.floor(new Date(`${yearStart}-01-01`).getTime() / 1000);
    const toTimestamp = Math.floor(new Date(`${yearEnd}-12-31`).getTime() / 1000);

    const url = `${URL_BASE}?resolution=D&symbol=${ticker.toUpperCase()}&from=${fromTimestamp}&to=${toTimestamp}`;

    const res = await (await fetch(url, HEADERS)).json();

    if (res && res.s === "ok") {
      const dates = res.t.map((seconds) => new Date(seconds * 1000).toLocaleDateString("vi-VN"));
      return { dates, prices: [...res.c] };
    }

    console.error("error at res!");
    return {};
  } catch (error) {
    console.error(`ERROR at ${ticker}: ${error.message}`);
    return {};
  }
};

const crawlAll = async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties.fullCalcOnLoad = true;

  for (const i in tickers) {
    const sheet = workbook.addWorksheet(tickers[i]);

    let lineOrder = 1;
    for (let year = Math.min(...years); year <= Math.max(...years); year++) {
      const { dates, prices } = await crawl(tickers[i], year);

      sheet.getRow(lineOrder).values = [year, ...dates];
      sheet.getRow(lineOrder + 1).values = ["", ...prices];

      lineOrder += 2;
    }

    console.log(`crawled ${tickers[i]} - ${+i + 1}`);
  }

  await workbook.xlsx.writeFile(FILE_PATH);
  console.log("Saved to", FILE_PATH);
};

/* const crawlAll = async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties.fullCalcOnLoad = true;

  let isExist = false;
  if (fs.existsSync(FILE_PATH)) {
    isExist = true;
    await workbook.xlsx.readFile(FILE_PATH);
  }

  for (const ticker of tickers) {
    let sheet;
    if (isExist) sheet = workbook.getWorksheet(ticker);
    else sheet = workbook.addWorksheet(ticker);

    let lineOrder = 1;
    for (const year of YEARS) {
      const { dates, prices } = await crawl(ticker, year, year);

      sheet.getRow(lineOrder).values = [year, ...dates];
      sheet.getRow(lineOrder + 1).values = ["", ...prices];

      lineOrder += 2;
    }

    console.log("crawled", ticker);
  }

  await workbook.xlsx.writeFile(FILE_PATH);
  console.log(`Saved to ${FILE_PATH}!`);
}; */

crawlAll();
