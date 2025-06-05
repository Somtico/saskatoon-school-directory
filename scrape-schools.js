const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

// Helper function to add delay between requests
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const elementarySchoolsRaw = [
  "École Alvin Buckwold School *",
  "John Lake School",
  "Brevoort Park School",
  "King George Community School",
  "Brownell School",
  "Lakeridge School",
  "Brunskill School",
  "École Lakeview School *",
  "Buena Vista School",
  "Lawson Heights School",
  "Caroline Robins Community School",
  "Lester B. Pearson School",
  "Caswell Community School",
  "Mayfair Community School",
  "Charles Red Hawk Elementary School",
  "Montgomery School",
  "Chief Whitecap School",
  "North Park Wilson School",
  "City Park School",
  "Prince Philip School",
  "Colette Bourgonje School",
  "Queen Elizabeth School",
  "École College Park School *",
  "École River Heights School *",
  "Dr. John G. Egnatoff School",
  "Roland Michener School",
  "École Dundonald School *",
  "École Silverspring School *",
  "Ernest Lindner School",
  "Silverwood Heights School",
  "Fairhaven School",
  "Sutherland Community School",
  "École Forest Grove School *",
  "Sylvia Fedoruk School",
  "Greystone Heights School",
  "École Victoria School **",
  "École Henry Kelsey School **",
  "Vincent Massey Community School",
  "Holliston School",
  "wâhkôhtowin Community School",
  "Howard Coad Community School",
  "Westmount Community School",
  "Hugh Cairns School",
  "Wildwood School",
  "James L. Alexander School",
  "Willowgrove School",
  "John Dolan School",
  "W.P. Bate Community School",
];
const highSchoolsRaw = [
  "Aden Bowman Collegiate",
  "Marion M. Graham Collegiate *",
  "Bedford Road Collegiate",
  "Mount Royal Collegiate",
  "Centennial Collegiate *",
  "Nutana Collegiate",
  "Estey School",
  "Tommy Douglas Collegiate *",
  "Evan Hardy Collegiate",
  "Walter Murray Collegiate *",
];

const removeGeneralWords = (name) =>
  name
    .replace(/school|community|collegiate|elementary/gi, "")
    .replace(/\s+/g, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // remove accents
    .replace(/[^a-zA-Z0-9]/g, "")
    .toLowerCase();

const getFrenchStatus = (name) => {
  if (/\*\*/.test(name)) return "French Immersion only";
  if (/\*/.test(name)) return "English and French";
  return "English only";
};

const cleanName = (name) => name.replace(/\*+/g, "").trim();

async function getContactDetails(page, url) {
  try {
    await page.goto(url, { waitUntil: "networkidle2", timeout: 30000 });
    await delay(1000);
    // Try to extract address, phone, email from the contact page
    const details = await page.evaluate(() => {
      const text = document.body.innerText;
      const addressMatch = text.match(/Address:\s*([\s\S]*?)\n/);
      const phoneMatch = text.match(/Phone:\s*([\d\-() ]+)/);
      const emailMatch = text.match(
        /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/
      );
      return {
        address: addressMatch ? addressMatch[1].trim() : "",
        phone: phoneMatch ? phoneMatch[1].trim() : "",
        email: emailMatch ? emailMatch[0].trim() : "",
      };
    });
    return details;
  } catch (e) {
    return { address: "", phone: "", email: "" };
  }
}

async function scrapeManualSchools(browser) {
  const schools = [];
  const page = await browser.newPage();
  for (const raw of elementarySchoolsRaw) {
    const name = cleanName(raw);
    const urlSegment = removeGeneralWords(name);
    const url = `https://www.spsd.sk.ca/school/${urlSegment}/Contact/Pages/default.aspx#/=`;
    const frenchStatus = getFrenchStatus(raw);
    const details = await getContactDetails(page, url);
    schools.push({
      Name: name,
      Type: "Elementary",
      Category: "Public",
      "French Status": frenchStatus,
      Address: details.address,
      Phone: details.phone,
      Email: details.email,
      URL: url,
      Principal: "",
    });
  }
  for (const raw of highSchoolsRaw) {
    const name = cleanName(raw);
    const urlSegment = removeGeneralWords(name);
    const url = `https://www.spsd.sk.ca/school/${urlSegment}/Contact/Pages/default.aspx#/=`;
    const frenchStatus = getFrenchStatus(raw);
    const details = await getContactDetails(page, url);
    schools.push({
      Name: name,
      Type: "High School",
      Category: "Public",
      "French Status": frenchStatus,
      Address: details.address,
      Phone: details.phone,
      Email: details.email,
      URL: url,
      Principal: "",
    });
  }
  await page.close();
  return schools;
}

async function main() {
  console.log("Starting to scrape school data...");
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  try {
    const schools = await scrapeManualSchools(browser);

    // Save to Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Schools");

    // Add headers
    worksheet.columns = [
      { header: "Name", key: "Name", width: 30 },
      { header: "Type", key: "Type", width: 15 },
      { header: "Category", key: "Category", width: 15 },
      { header: "French Status", key: "French Status", width: 20 },
      { header: "Address", key: "Address", width: 40 },
      { header: "Phone", key: "Phone", width: 20 },
      { header: "Email", key: "Email", width: 30 },
      { header: "URL", key: "URL", width: 40 },
      { header: "Principal", key: "Principal", width: 30 },
    ];

    // Add rows
    worksheet.addRows(schools);

    // Save the file
    const outputPath = "schools.xlsx";
    await workbook.xlsx.writeFile(outputPath);

    console.log(`Successfully scraped ${schools.length} schools`);
    console.log(`Data saved to ${process.cwd()}/${outputPath}`);
  } catch (error) {
    console.error("Error:", error);
  } finally {
    await browser.close();
    process.exit(0);
  }
}

main();
