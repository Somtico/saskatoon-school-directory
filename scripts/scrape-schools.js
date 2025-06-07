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
  if (platform() === "win32") {
    exec('powershell -c "[console]::beep(1000,500)"');
  } else if (platform() === "darwin") {
    exec("afplay /System/Library/Sounds/Glass.aiff");
  } else {
    exec(
      "paplay /usr/share/sounds/freedesktop/stereo/complete.oga || aplay /usr/share/sounds/alsa/Front_Center.wav || beep"
    );
  }
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

async function getContactDetails(page, url) {
  try {
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
    await delay(2000);

    // 1. Find the Contact Us link in the main navigation
    const contactHref = await page.evaluate(() => {
      const nav = document.querySelector(".main-nav");
      if (!nav) return null;
      const links = nav.querySelectorAll("a");
      for (const link of links) {
        if (
          link.textContent &&
          link.textContent.toLowerCase().includes("contact")
        ) {
          return link.getAttribute("href");
        }
      }
      return null;
    });

    if (!contactHref) {
      console.warn(`No contact link found on ${url}`);
      return { address: "", phone: "", email: "", contactPageUrl: "" };
    }

    const contactUrl = new URL(contactHref, url).toString();
    await page.goto(contactUrl, { waitUntil: "networkidle2" });

    // 4. Extract contact info from the contact page
    const details = await page.evaluate(() => {
      function getText(selector) {
        const el = document.querySelector(selector);
        return el ? el.innerText.trim() : "";
      }
      function getPhoneNumber() {
        const phoneEl = document.querySelector(".ci-contact-list .number");
        return phoneEl ? phoneEl.innerText.trim() : "";
      }
      function getEmailFromLink() {
        const emailEl = document.querySelector(".contactpg_email a");
        return emailEl ? emailEl.innerText.trim() : "";
      }
      return {
        address: getText("address"),
        phone: getPhoneNumber(),
        email: getEmailFromLink(),
      };
    });

    console.log(`Found for ${contactUrl}:`, details);
    return { ...details, contactPageUrl: contactUrl };
  } catch (e) {
    console.error(`Error getting contact details for ${url}:`, e);
    return { address: "", phone: "", email: "", contactPageUrl: "" };
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

async function main() {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    args: ["--start-maximized"],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 });

    // 1. Get the list of schools
    const schools = await getSchoolList(page);
    console.log(`Found ${schools.length} schools`);

    // 2. For each school, follow its link and extract contact info
    const schoolDetails = [];
    for (const school of schools) {
      console.log(`Scraping ${school.name}...`);
      const details = await getContactDetails(page, school.url);
      console.log(`Found for ${school.url}:`, details);
      schoolDetails.push({
        Type: getSchoolType(school.name),
        Category: "Catholic",
        "French Status": getFrenchStatus(school.name),
        Name: school.name,
        Address: details.address,
        URL: school.url,
        Phone: details.phone,
        Email: details.email,
        ContactPageURL: details.contactPageUrl,
      });
    }

    // 3. Save the data
    // Save to Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Catholic Schools");

    // Add headers
    worksheet.columns = [
      { header: "Type", key: "Type" },
      { header: "Category", key: "Category" },
      { header: "French Status", key: "French Status" },
      { header: "Name", key: "Name" },
      { header: "Address", key: "Address" },
      { header: "URL", key: "URL" },
      { header: "Phone", key: "Phone" },
      { header: "Email", key: "Email" },
      { header: "Contact Page URL", key: "ContactPageURL" },
    ];

    // Add rows
    worksheet.addRows(schoolDetails);

    // Save the file with a unique name to avoid locking issues
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const schoolType = getSchoolType(schools[0].name);
    const excelFilePath = path.join(
      __dirname,
      `${schoolType}-schools-${timestamp}.xlsx`
    );
    await workbook.xlsx.writeFile(excelFilePath);
    console.log(`\nData has been saved to: ${excelFilePath}`);

    // Also save as JSON for backup
    const jsonPath = path.join(
      __dirname,
      `${schoolType}-schools-${timestamp}.json`
    );
    fs.writeFileSync(jsonPath, JSON.stringify(schoolDetails, null, 2));
    console.log(`Data also saved to ${jsonPath}`);

    // Open the file
    openFile(excelFilePath);

    // Play notification sound
    playNotificationSound();
  } catch (e) {
    console.error("Error in main:", e);
  } finally {
    await browser.close();
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
