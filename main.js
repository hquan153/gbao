const fs = require("fs");
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
  {
    groupName: "Cụm ngân hàng",
    companies: ["ACB", "BID", "CTG", "HDB", "LPB", "MBB", "SHB", "SSB", "STB", "TCB", "VCB", "VIB", "VPB"],
    average: {
      acb: [],
      bid: [],
      ctg: [],
      hdb: [],
      lpb: [],
      mbb: [],
      shb: [],
      ssb: [],
      stb: [],
      tcb: [],
      vcb: [],
      vib: [],
      vpb: [],
    },
    standardDeviation: {
      acb: [],
      bid: [],
      ctg: [],
      hdb: [],
      lpb: [],
      mbb: [],
      shb: [],
      ssb: [],
      stb: [],
      tcb: [],
      vcb: [],
      vib: [],
      vpb: [],
    },
  },
  {
    groupName: "Cụm bất động sản",
    companies: ["VIC", "VHM", "VRE", "VPL"],
    average: { vic: [], vhm: [], vre: [], vpl: [] },
    standardDeviation: { vic: [], vhm: [], vre: [], vpl: [] },
  },
  {
    groupName: "Nhóm công nghệ thông tin và viễn thông",
    companies: ["FPT"],
    average: { fpt: [] },
    standardDeviation: { fpt: [] },
  },
  {
    groupName: "Nhóm tiêu dùng, bán lẻ và sản xuất",
    companies: ["MSN", "MCH", "MWG", "VNM", "SAB"],
    average: { msn: [], mch: [], mwg: [], vnm: [], sab: [] },
    standardDeviation: { msn: [], mch: [], mwg: [], vnm: [], sab: [] },
  },
  {
    groupName: "Nhóm năng lượng, dầu khí",
    companies: ["GAS", "BSR", "POW"],
    average: { gas: [], bsr: [], pow: [] },
    standardDeviation: { gas: [], bsr: [], pow: [] },
  },
  {
    groupName: "Nhóm công nghiệp, nguyên vật liệu và cao su",
    companies: ["HPG", "GVR"],
    average: { hpg: [], gvr: [] },
    standardDeviation: { hpg: [], gvr: [] },
  },
  {
    groupName: "Nhóm tài chính - chứng khoán và bảo hiểm",
    companies: ["SSI", "TCX", "BVH"],
    average: { ssi: [], tcx: [], bvh: [] },
    standardDeviation: { ssi: [], tcx: [], bvh: [] },
  },
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

const getAverage = (nums) => nums.reduce((total, num) => total + num, 0) / nums.length || 0;

const standardDeviation = (nums, average) =>
  Math.sqrt(nums.reduce((total, num) => total + (num - average) ** 2, 0) / nums.length) || 0;

const crawlAll = async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties.fullCalcOnLoad = true;

  if (fs.existsSync(FILE_PATH)) fs.unlinkSync(FILE_PATH);

  // crawl
  for (const group of tickers) {
    for (const ticker of group.companies) {
      const sheet = workbook.addWorksheet(ticker);

      let lineOrder = 1;
      for (let year = Math.min(...years); year <= Math.max(...years); year++) {
        const { dates, prices } = await crawl(ticker, year);

        const average = getAverage([...prices]);
        group.average[ticker.toLowerCase()].push(average);
        group.standardDeviation[ticker.toLowerCase()].push(standardDeviation([...prices], average));

        sheet.getRow(lineOrder).values = [year, ...dates];
        sheet.getRow(lineOrder + 1).values = ["", ...prices];
        lineOrder += 2;
      }

      console.log("crawled", ticker);
    }

    console.log("--CRAWLED", group.groupName.toUpperCase());
  }

  // caculate
  for (const group of tickers) {
    const sheet = workbook.addWorksheet(group.groupName);
    let lineOrder = 1;

    for (const ticker of group.companies) {
      sheet.mergeCells(lineOrder, 1, lineOrder, 3);

      const cell = sheet.getCell(lineOrder, 1);
      cell.value = ticker;

      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };

      sheet.getRow(lineOrder + 1).values = ["", "Trung bình", "Độ lệch chuẩn"];

      for (let i = 0; i < Math.max(...years) - Math.min(...years) + 1; i++) {
        sheet.getRow(lineOrder + 2 + i).values = [
          Math.min(...years) + i,
          group.average[ticker.toLowerCase()][i],
          group.standardDeviation[ticker.toLowerCase()][i],
        ];
      }

      lineOrder += Math.max(...years) - Math.min(...years) + 1 + 2 + 1;
    }
  }

  // adjust
  workbook.eachSheet((sheet) => {
    sheet.columns.forEach((column) => {
      let maxChars = 8;
      column.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.isMerged) return;

        const cellValue = cell.value ? cell.value.toString() : "";
        if (cellValue.length > maxChars) maxChars = cellValue.length;
      });
      column.width = maxChars < 12 ? 12 : maxChars + 2;
    });
  });

  await workbook.xlsx.writeFile(FILE_PATH);
  console.log("Saved to", FILE_PATH);
};

crawlAll();
