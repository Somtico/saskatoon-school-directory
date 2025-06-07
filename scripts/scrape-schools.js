const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");
const { exec } = require("child_process");
const { platform } = require("os");
const path = require("path");
const fs = require("fs");

// Helper function to add delay between requests
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Helper function to play notification sound
function playNotificationSound() {
  console.log("\u0007"); // This will play the system notification sound
}

// Helper function to open file
function openFile(filePath) {
  const isWindows = platform() === "win32";
  if (isWindows) {
    exec(`start "" "${filePath}"`);
  } else {
    exec(`open "${filePath}"`);
  }
}

const catholicSchools = ["Bethlehem Catholic High School"];

const urlOverrides = {
  "Bethlehem Catholic High School": "BET",
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

// Helper to extract phone numbers and emails from text
function extractPhonesAndEmails(text) {
  const phones = Array.from(text.matchAll(/(306[-.\s]?\d{3}[-.\s]?\d{4})/g)).map(m => m[0]);
  const emails = Array.from(text.matchAll(/[\w.-]+@[\w.-]+\.[A-Za-z]{2,}/g)).map(m => m[0]);
  return { phones, emails };
}

async function getContactInfoFromUrl(page, url) {
  try {
    await page.goto(url, { waitUntil: "networkidle0" });
    const phones = await page.evaluate(() => {
      const text = document.body.innerText;
      const phoneRegex = /(?:^|\D)(\d{3}[-.]?\d{3}[-.]?\d{4})(?:\D|$)/g;
      const matches = [...text.matchAll(phoneRegex)];
      return [...new Set(matches.map((m) => m[1]))];
    });
    const emails = await page.evaluate(() => {
      const text = document.body.innerText;
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
      const matches = text.match(emailRegex) || [];
      return [...new Set(matches)];
    });
    return { contactPageUrl: url, phones, emails };
  } catch (err) {
    console.error(`Error getting contact info from ${url}:`, err.message);
    return { contactPageUrl: url, phones: [], emails: [] };
  }
}

async function getSchoolList(page) {
  try {
    await page.goto("https://www.gscs.ca/page/63/find-a-school", {
      waitUntil: "domcontentloaded",
      timeout: 30000,
    });
    await delay(2000);

    const schools = await page.evaluate(() => {
      const rows = document.querySelectorAll(".cifs_listview table tbody tr");
      return Array.from(rows).map((row) => {
        const nameCell = row.querySelector(".rowtitle a");
        const name = nameCell ? nameCell.textContent.trim() : "";
        const url = nameCell ? nameCell.getAttribute("href") : "";
        return { name, url };
      });
    });

    return schools;
  } catch (e) {
    console.error("Error getting school list:", e);
    return [];
  }
}

// Function to check if a school exists in the list
function checkExistingSchools(schoolName, existingSchools) {
  return existingSchools.some(school => 
    school.name.toLowerCase() === schoolName.toLowerCase()
  );
}

// Main function
async function main() {
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });
  await page.setDefaultNavigationTimeout(30000); // 30 second timeout

  // Load existing schools data
  const catholicSchoolsPath = path.join(__dirname, "catholic-schools.json");
  const publicSchoolsPath = path.join(__dirname, "public-schools.json");

  let existingCatholicSchools = [];
  let existingPublicSchools = [];

  try {
    if (fs.existsSync(catholicSchoolsPath)) {
      existingCatholicSchools = JSON.parse(
        fs.readFileSync(catholicSchoolsPath, "utf8")
      );
    }
    if (fs.existsSync(publicSchoolsPath)) {
      existingPublicSchools = JSON.parse(
        fs.readFileSync(publicSchoolsPath, "utf8")
      );
    }
  } catch (err) {
    console.error("Error loading existing schools data:", err);
  }

  // Scrape Catholic schools
  console.log("Scraping Catholic schools...");
  await page.goto("https://www.gscs.ca/schools", { waitUntil: "networkidle0" });
  const catholicSchools = await page.evaluate(() => {
    const schools = [];
    const schoolElements = document.querySelectorAll(".school-list-item");
    schoolElements.forEach((element) => {
      const name = element.querySelector(".school-name")?.textContent.trim();
      const address = element
        .querySelector(".school-address")
        ?.textContent.trim();
      if (name && address) {
        schools.push({ name, address });
      }
    });
    return schools;
  });

  // Process Catholic schools
  const processedCatholicSchools = [];
  for (const school of catholicSchools) {
    if (!checkExistingSchools(school.name, existingCatholicSchools)) {
      const { contactPageUrl, phones, emails } = await getContactInfoFromUrl(
        page,
        school.contactPageUrl
      );
      processedCatholicSchools.push({
        ...school,
        phones: phones.join(", "),
        emails: emails.join(", "),
        contactPageUrl,
      });
    }
  }

  // Save Catholic schools to Excel
  const catholicWorkbook = new ExcelJS.Workbook();
  const catholicSheet = catholicWorkbook.addWorksheet("Catholic Schools");
  catholicSheet.columns = [
    { header: "Name", key: "name" },
    { header: "Address", key: "address" },
    { header: "Phones", key: "phones" },
    { header: "Emails", key: "emails" },
    { header: "Contact Page URL", key: "contactPageUrl" },
  ];
  processedCatholicSchools.forEach((school) => catholicSheet.addRow(school));
  const catholicTimestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const catholicExcelPath = path.join(
    __dirname,
    `catholic-schools-${catholicTimestamp}.xlsx`
  );
  await catholicWorkbook.xlsx.writeFile(catholicExcelPath);
  console.log(`Catholic schools saved to: ${catholicExcelPath}`);

  // Save Catholic schools to JSON
  const catholicJsonPath = path.join(
    __dirname,
    `catholic-schools-${catholicTimestamp}.json`
  );
  fs.writeFileSync(
    catholicJsonPath,
    JSON.stringify(processedCatholicSchools, null, 2)
  );
  console.log(`Catholic schools saved to: ${catholicJsonPath}`);

  // Scrape public schools
  console.log("Scraping public schools...");
  await page.goto("https://www.saskatoonpublicschools.ca/schools", {
    waitUntil: "networkidle0",
  });
  const publicSchools = await page.evaluate(() => {
    const schools = [];
    const schoolElements = document.querySelectorAll(".school-list-item");
    schoolElements.forEach((element) => {
      const name = element.querySelector(".school-name")?.textContent.trim();
      const address = element
        .querySelector(".school-address")
        ?.textContent.trim();
      if (name && address) {
        schools.push({ name, address });
      }
    });
    return schools;
  });

  // Process public schools
  const processedPublicSchools = [];
  for (const school of publicSchools) {
    if (!checkExistingSchools(school.name, existingPublicSchools)) {
      const { contactPageUrl, phones, emails } = await getContactInfoFromUrl(
        page,
        school.contactPageUrl
      );
      processedPublicSchools.push({
        ...school,
        phones: phones.join(", "),
        emails: emails.join(", "),
        contactPageUrl,
      });
    }
  }

  // Save public schools to Excel
  const publicWorkbook = new ExcelJS.Workbook();
  const publicSheet = publicWorkbook.addWorksheet("Public Schools");
  publicSheet.columns = [
    { header: "Name", key: "name" },
    { header: "Address", key: "address" },
    { header: "Phones", key: "phones" },
    { header: "Emails", key: "emails" },
    { header: "Contact Page URL", key: "contactPageUrl" },
  ];
  processedPublicSchools.forEach((school) => publicSheet.addRow(school));
  const publicTimestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const publicExcelPath = path.join(
    __dirname,
    `public-schools-${publicTimestamp}.xlsx`
  );
  await publicWorkbook.xlsx.writeFile(publicExcelPath);
  console.log(`Public schools saved to: ${publicExcelPath}`);

  // Save public schools to JSON
  const publicJsonPath = path.join(
    __dirname,
    `public-schools-${publicTimestamp}.json`
  );
  fs.writeFileSync(
    publicJsonPath,
    JSON.stringify(processedPublicSchools, null, 2)
  );
  console.log(`Public schools saved to: ${publicJsonPath}`);

  await browser.close();
  playNotificationSound();
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

main().catch((e) => {
  console.error("Fatal error in main:", e);
  process.exit(1);
});
