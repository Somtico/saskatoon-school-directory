const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

// Helper function to add delay between requests
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const catholicSchools = [
  "Bethlehem Catholic High School",
  "Bishop Filevich Ukrainian Bilingual School",
  "Bishop James Mahoney High School",
  "Bishop Klein Community School",
  "Bishop Murray High School",
  "Bishop Pocock School",
  "Bishop Roborecki Community School",
  "Cyber School",
  "E. D. Feehan Catholic High School",
  "École Cardinal Leger School",
  "École Father Robinson School",
  "École française de Saskatoon",
  "École Holy Mary Catholic School",
  "École Sister O'Brien School",
  "École St. Gerard School",
  "École St. Luke School",
  "École St. Matthew School",
  "École St. Mother Teresa School",
  "École St. Paul School",
  "École St. Peter School",
  "Father Vachon School",
  "Georges Vanier Catholic Fine Arts School",
  "Holy Cross High School",
  "Holy Family Catholic School",
  "Holy Trinity Catholic School",
  "International Student Program",
  "Oskāyak High School",
  "Pope John Paul II School",
  "St. Angela School",
  "St. Anne School",
  "St. Augustine School",
  "St. Augustine School - Humboldt",
  "St. Bernard School",
  "St. Dominic School",
  "St. Dominic School - Humboldt",
  "St. Edward School",
  "St. Frances Cree Bilingual School – Bateman",
  "St. Frances Cree Bilingual School - McPherson",
  "St. Gabriel Biggar",
  "St. George School",
  "St. John Community School",
  "St. Joseph High School",
  "St. Kateri Tekakwitha Catholic School",
  "St. Lorenzo Ruiz Catholic School",
  "St. Marguerite School",
  "St. Maria Goretti Community School",
  "St. Mark Community School",
  "St. Mary's Wellness and Education Centre",
  "St. Michael Community School",
  "St. Nicholas Catholic School",
  "St. Philip School",
  "St. Thérèse of Lisieux Catholic School",
  "St. Volodymyr School",
];

const urlOverrides = {
  "Bethlehem Catholic High School": "BET",
  "Bishop Filevich Ukrainian Bilingual School": "FIL",
  "Bishop James Mahoney High School": "BJM",
  "Bishop Klein Community School": "KLE",
  "Bishop Murray High School": "BMH",
  "Bishop Pocock School": "POC",
  "Bishop Roborecki Community School": "ROB",
  "Cyber School": "cyb",
  "E. D. Feehan Catholic High School": "EDF",
  "École Cardinal Leger School": "LEG",
  "École Father Robinson School": "RBI",
  "École française de Saskatoon": "FRE",
  "École Holy Mary Catholic School": "HMA",
  "École Sister O'Brien School": "OBR",
  "École St. Gerard School": "GER",
  "École St. Luke School": "LUK",
  "École St. Matthew School": "MAT",
  "École St. Mother Teresa School": "TER",
  "École St. Paul School": "PAU",
  "École St. Peter School": "PET",
  "Father Vachon School": "VAC",
  "Georges Vanier Catholic Fine Arts School": "VAN",
  "Holy Cross High School": "HCH",
  "Holy Family Catholic School": "FAM",
  "Holy Trinity Catholic School": "HTR",
  "International Student Program": "ISP",
  "Oskāyak High School": "OSK",
  "Pope John Paul II School": "JP2",
  "St. Angela School": "ANG",
  "St. Anne School": "ANN",
  "St. Augustine School": "AUG",
  "St. Augustine School - Humboldt": "HAU",
  "St. Bernard School": "BER",
  "St. Dominic School": "DOM",
  "St. Dominic School - Humboldt": "HDO",
  "St. Edward School": "EDW",
  "St. Frances Cree Bilingual School – Bateman": "frb",
  "St. Frances Cree Bilingual School - McPherson": "fra",
  "St. Gabriel Biggar": "BGA",
  "St. George School": "GEO",
  "St. John Community School": "JOH",
  "St. Joseph High School": "JOS",
  "St. Kateri Tekakwitha Catholic School": "kat",
  "St. Lorenzo Ruiz Catholic School": "lor",
  "St. Marguerite School": "MAG",
  "St. Maria Goretti Community School": "GOR",
  "St. Mark Community School": "MAK",
  "St. Mary's Wellness and Education Centre": "MRY",
  "St. Michael Community School": "MIC",
  "St. Nicholas Catholic School": "nic",
  "St. Philip School": "PHI",
  "St. Thérèse of Lisieux Catholic School": "the",
  "St. Volodymyr School": "VOL",
};

const removeGeneralWords = (name) =>
  name
    .replace(/^école|^ecole/gi, "") // remove 'École' or 'Ecole' at the start
    .replace(/school|community|collegiate|elementary|high|centre|center/gi, "")
    .replace(/\s+/g, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // remove accents
    .replace(/[^a-zA-Z0-9]/g, "")
    .toLowerCase();

const getSchoolType = (name) => {
  if (name.toLowerCase().includes("high")) return "High School";
  if (name.toLowerCase().includes("cyber")) return "Online";
  if (name.toLowerCase().includes("international")) return "Special Program";
  return "Elementary";
};

const getFrenchStatus = (name) => {
  if (name.toLowerCase().includes("french")) return "French Immersion";
  if (name.toLowerCase().includes("bilingual")) return "Bilingual";
  return "English";
};

async function getContactDetails(page, url) {
  try {
    // Wait for the page to be ready
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
    await delay(2000); // Give extra time for dynamic content to load

    // Extract contact information from the page
    const details = await page.evaluate(() => {
      const getText = (selector) => {
        const el = document.querySelector(selector);
        return el ? el.textContent.trim().replace(/\s+/g, " ") : "";
      };

      // Try different possible selectors for contact information
      const address =
        getText(".address") ||
        getText("[itemprop='address']") ||
        getText(".contact-address") ||
        getText(".school-address") ||
        getText(".school-info .address");
      const phone =
        getText(".phone") ||
        getText("[itemprop='telephone']") ||
        getText(".contact-phone") ||
        getText(".school-phone") ||
        getText(".school-info .phone");
      const email =
        getText(".email") ||
        getText("[itemprop='email']") ||
        getText(".contact-email") ||
        getText(".school-email") ||
        getText(".school-info .email");

      return { address, phone, email };
    });

    return details;
  } catch (e) {
    console.error(`Error getting contact details for ${url}:`, e);
    return { address: "", phone: "", email: "" };
  }
}

async function scrapeCatholicSchools(browser) {
  const schools = [];
  const page = await browser.newPage();

  // Set a longer timeout for navigation
  page.setDefaultNavigationTimeout(60000);

  for (const name of catholicSchools) {
    const urlCode = urlOverrides[name];
    const url = `https://www.gscs.ca/${urlCode}`;
    const schoolType = getSchoolType(name);
    const frenchStatus = getFrenchStatus(name);

    console.log(`Scraping ${name}...`);
    const details = await getContactDetails(page, url);

    schools.push({
      Type: schoolType,
      Category: "Catholic",
      "French Status": frenchStatus,
      Name: name,
      Address: details.address,
      URL: url,
      Phone: details.phone,
      Email: details.email,
    });

    // Add a delay between requests to be respectful to the server
    await delay(3000);
  }

  await page.close();
  return schools;
}

async function main() {
  console.log("Starting to scrape Catholic school data...");
  let browser;

  try {
    browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    const schools = await scrapeCatholicSchools(browser);

    // Save to Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Schools");

    // Add headers
    worksheet.columns = [
      { header: "Type", key: "Type", width: 15 },
      { header: "Category", key: "Category", width: 15 },
      { header: "French Status", key: "French Status", width: 20 },
      { header: "Name", key: "Name", width: 30 },
      { header: "Address", key: "Address", width: 40 },
      { header: "URL", key: "URL", width: 40 },
      { header: "Phone", key: "Phone", width: 20 },
      { header: "Email", key: "Email", width: 30 },
    ];

    // Add rows
    worksheet.addRows(schools);

    // Save the file
    const outputPath = "catholic-schools.xlsx";
    await workbook.xlsx.writeFile(outputPath);

    console.log(`Successfully scraped ${schools.length} schools`);
    console.log(`Data saved to ${process.cwd()}/${outputPath}`);
  } catch (error) {
    console.error("Error:", error);
    process.exit(1); // Exit with error code
  } finally {
    if (browser) {
      await browser.close();
    }
    process.exit(0); // Exit with success code
  }
}

// Handle uncaught exceptions
process.on("uncaughtException", (error) => {
  console.error("Uncaught Exception:", error);
  process.exit(1);
});

// Handle unhandled promise rejections
process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled Rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main();
