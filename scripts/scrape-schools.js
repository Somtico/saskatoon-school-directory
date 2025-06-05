const puppeteer = require("puppeteer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// Create data directory if it doesn't exist
const dataDir = path.join(__dirname, "..", "data");
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir);
}

async function scrapePublicSchools(page) {
  console.log("Scraping public schools...");
  const schools = [];

  try {
    await page.goto("https://www.spsd.sk.ca/Schools/Pages/default.aspx#/=", {
      waitUntil: "networkidle0",
    });

    // Wait for the school tables to load
    await page.waitForSelector(".ms-rteTable-1");

    // Get all school tables (elementary and secondary)
    const tables = await page.$$(".ms-rteTable-1");

    for (const table of tables) {
      // Get all school cells from the table
      const schoolCells = await table.$$("td");

      for (const cell of schoolCells) {
        try {
          // Get the school link
          const linkElement = await cell.$("a");
          if (!linkElement) continue;

          const name = await page.evaluate(
            (el) => el.textContent.trim(),
            linkElement
          );
          const url = await page.evaluate((el) => el.href, linkElement);

          // Determine if it's a high school based on the table context
          const isHighSchool = await page.evaluate((table) => {
            const tableText = table.textContent.toLowerCase();
            return (
              tableText.includes("secondary") ||
              tableText.includes("collegiate")
            );
          }, table);

          // Determine if it's a French immersion school
          const isFrenchImmersion =
            name.toLowerCase().includes("école") ||
            name.toLowerCase().includes("french") ||
            (await page.evaluate((cell) => {
              const cellText = cell.textContent;
              return cellText.includes("*") || cellText.includes("French");
            }, cell));

          // Visit the school's page to get additional details
          await page.goto(url, { waitUntil: "networkidle0" });

          // Try to get address and phone from the school's page
          let address = "";
          let phone = "";

          try {
            const addressElement = await page.$(".address");
            if (addressElement) {
              address = await page.evaluate(
                (el) => el.textContent.trim(),
                addressElement
              );
            }

            const phoneElement = await page.$(".phone");
            if (phoneElement) {
              phone = await page.evaluate(
                (el) => el.textContent.trim(),
                phoneElement
              );
            }
          } catch (err) {
            console.log(`Could not get details for ${name}: ${err.message}`);
          }

          schools.push({
            name,
            address,
            phone,
            url,
            type: isHighSchool ? "High School" : "Elementary School",
            isFrenchImmersion,
            board: "Public",
            principal: "", // Will be left empty as requested
            superintendent: "", // Will be left empty as requested
          });

          // Go back to the main page
          await page.goto(
            "https://www.spsd.sk.ca/Schools/Pages/default.aspx#/=",
            { waitUntil: "networkidle0" }
          );
          await page.waitForSelector(".ms-rteTable-1");
        } catch (err) {
          console.error(`Error scraping public school cell: ${err.message}`);
          continue;
        }
      }
    }
  } catch (err) {
    console.error(`Error scraping public schools: ${err.message}`);
  }

  return schools;
}

async function scrapeCatholicSchools(page) {
  console.log("Scraping Catholic schools...");
  const schools = [];

  try {
    await page.goto("https://www.gscs.ca/page/63/find-a-school", {
      waitUntil: "networkidle0",
    });

    // Wait for the school list table to load
    await page.waitForSelector(".cifs_listview table.table");

    // Get all school rows
    const schoolRows = await page.$$(".cifs_listview table.table tbody tr");

    for (const row of schoolRows) {
      try {
        // Extract school name and URL
        const nameElement = await row.$(".rowtitle a");
        const name = await page.evaluate(
          (el) => el.textContent.trim(),
          nameElement
        );
        const url = await page.evaluate((el) => el.href, nameElement);

        // Extract address
        const addressElement = await row.$("td:nth-child(2) a");
        const address = await page.evaluate(
          (el) => el.textContent.trim(),
          addressElement
        );

        // Extract phone
        const phoneElement = await row.$("td:nth-child(3)");
        const phone = await page.evaluate(
          (el) => el.textContent.trim(),
          phoneElement
        );

        // Determine if it's a high school based on the name
        const isHighSchool = name.toLowerCase().includes("high school");

        // Determine if it's a French immersion school
        const isFrenchImmersion =
          name.toLowerCase().includes("école") ||
          name.toLowerCase().includes("french");

        schools.push({
          name,
          address,
          phone,
          url,
          type: isHighSchool ? "High School" : "Elementary School",
          isFrenchImmersion,
          board: "Catholic",
          principal: "", // Will be left empty as requested
          superintendent: "", // Will be left empty as requested
        });
      } catch (err) {
        console.error(`Error scraping Catholic school row: ${err.message}`);
        continue;
      }
    }
  } catch (err) {
    console.error(`Error scraping Catholic schools: ${err.message}`);
  }

  return schools;
}

async function scrapePrivateSchools(page) {
  const schools = [];

  try {
    // Navigate to Google search for private schools in Saskatoon
    await page.goto(
      "https://www.google.com/search?q=private+schools+in+saskatoon"
    );

    // Wait for search results to load
    await page.waitForSelector("#search");

    const searchResults = await page.evaluate(() => {
      const results = document.querySelectorAll("#search .g");
      const schools = [];

      results.forEach((result) => {
        const title = result.querySelector("h3")?.textContent.trim() || "";
        const link = result.querySelector("a")?.href || "";
        const snippet =
          result.querySelector(".VwiC3b")?.textContent.trim() || "";

        // Only include if it looks like a school and isn't already in our list
        if (
          title.toLowerCase().includes("school") &&
          !existingSchoolNames.includes(title.toLowerCase())
        ) {
          schools.push({
            category: "Private",
            level: "Both", // We'll need to determine this from the school's website
            name: title,
            address: "", // We'll need to get this from the school's website
            website: link,
            phone: "", // We'll need to get this from the school's website
            email: "",
            principal: "",
            superintendent: "",
            contactPerson: "",
          });
        }
      });

      return schools;
    });

    // Visit each private school's website to get more details
    for (let school of schools) {
      try {
        await page.goto(school.website);

        // Try to find address and phone
        const details = await page.evaluate(() => {
          const address =
            document.querySelector("address")?.textContent.trim() || "";
          const phone =
            document.querySelector("a[href^='tel:']")?.textContent.trim() || "";
          return { address, phone };
        });

        school.address = details.address;
        school.phone = details.phone;
      } catch (error) {
        console.error(`Error scraping details for ${school.name}:`, error);
      }
    }
  } catch (err) {
    console.error(`Error scraping private schools: ${err.message}`);
  }

  return schools;
}

async function main() {
  let browser;
  try {
    console.log("Starting to scrape school data...");

    // Initialize browser
    browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();

    // Set a longer timeout for navigation
    page.setDefaultNavigationTimeout(60000);

    // Scrape schools from all sources
    const [publicSchools, catholicSchools, privateSchools] = await Promise.all([
      scrapePublicSchools(page),
      scrapeCatholicSchools(page),
      scrapePrivateSchools(page),
    ]);

    // Combine all schools
    const allSchools = [
      ...publicSchools,
      ...catholicSchools,
      ...privateSchools,
    ];

    // Save to Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Schools");

    // Add headers
    worksheet.columns = [
      { header: "Name", key: "name", width: 40 },
      { header: "Type", key: "type", width: 15 },
      { header: "Board", key: "board", width: 15 },
      { header: "Address", key: "address", width: 50 },
      { header: "Phone", key: "phone", width: 20 },
      { header: "URL", key: "url", width: 50 },
      { header: "French Immersion", key: "isFrenchImmersion", width: 15 },
      { header: "Principal", key: "principal", width: 30 },
      { header: "Superintendent", key: "superintendent", width: 30 },
    ];

    // Add rows
    worksheet.addRows(allSchools);

    // Save the file
    const outputPath = path.join(__dirname, "schools.xlsx");
    await workbook.xlsx.writeFile(outputPath);

    console.log(`Successfully scraped ${allSchools.length} schools`);
    console.log(`Data saved to ${outputPath}`);
  } catch (err) {
    console.error("Error scraping schools:", err);
    process.exit(1);
  } finally {
    if (browser) {
      await browser.close();
    }
    process.exit(0);
  }
}

// Run the scraper
main();
