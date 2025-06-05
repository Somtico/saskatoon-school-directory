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

// Helper function to add delay between requests
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function scrapePublicSchools(page) {
  console.log("Scraping public schools...");
  const schools = [];

  try {
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    );
    await page.setExtraHTTPHeaders({
      "Accept-Language": "en-US,en;q=0.9",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    });

    await page.goto("https://www.spsd.sk.ca/Schools/Pages/default.aspx", {
      waitUntil: "networkidle0",
      timeout: 60000,
    });

    await page.waitForSelector(".ms-rteTable-1", { timeout: 30000 });

    // Helper to extract school info from a table
    async function extractSchoolsFromTable(tableSelector, type) {
      return await page.$$eval(tableSelector, (tables, type) => {
        // Each table row contains one or two schools (in <td>s)
        const schools = [];
        tables.forEach(table => {
          const rows = Array.from(table.querySelectorAll('tr'));
          rows.forEach(row => {
            const tds = Array.from(row.querySelectorAll('td'));
            tds.forEach(td => {
              // Find the <a> tag (school link)
              const a = td.querySelector('a');
              if (a) {
                let name = a.textContent.trim();
                let url = a.href;
                let rest = td.textContent.replace(name, '').trim();
                // Check for asterisk(s) after the name
                let status = 'English only';
                if (/\*\*/.test(rest)) {
                  status = 'French Immersion only';
                } else if (/\*/.test(rest)) {
                  status = 'English and French available';
                }
                schools.push({
                  name,
                  url,
                  type,
                  frenchStatus: status
                });
              }
            });
          });
        });
        return schools;
      }, type);
    }

    // The first .ms-rteTable-1 is elementary, the second is high school
    const tables = await page.$$('.ms-rteTable-1');
    let elementarySchools = [];
    let highSchools = [];
    if (tables.length >= 2) {
      elementarySchools = await extractSchoolsFromTable('.ms-rteTable-1:nth-of-type(1)', 'Elementary');
      highSchools = await extractSchoolsFromTable('.ms-rteTable-1:nth-of-type(2)', 'High School');
    }
    const allSchools = [...elementarySchools, ...highSchools];

    // Now visit each school page for address, phone, email
    for (const school of allSchools) {
      try {
        await page.goto(school.url, { waitUntil: "networkidle0", timeout: 30000 });
        await delay(1000);
        let address = "";
        let phone = "";
        let email = "";
        try {
          const paragraphs = await page.$$eval('p', ps => ps.map(p => p.textContent.trim()));
          for (const text of paragraphs) {
            if (!address && /\d+\s+\w+/.test(text) && /Saskatoon/i.test(text)) address = text;
            if (!phone && /\d{3}[-.\s]?\d{3}[-.\s]?\d{4}/.test(text)) phone = text;
            if (!email && /@/.test(text)) email = text;
          }
        } catch (error) {
          console.log(`Error getting details for ${school.name}: ${error.message}`);
        }
        schools.push({
          name: school.name,
          type: school.type,
          category: "Public",
          frenchStatus: school.frenchStatus,
          address,
          phone,
          email,
          url: school.url,
          principal: "",
          superintendent: ""
        });
      } catch (error) {
        console.log(`Error processing school page for ${school.name}: ${error.message}`);
        continue;
      }
    }
  } catch (error) {
    console.log(`Error scraping public schools: ${error.message}`);
  }
  return schools;
}

async function scrapeCatholicSchools(page) {
  console.log("Scraping Catholic schools...");
  const schools = [];

  try {
    // Set user agent and other headers
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    );
    await page.setExtraHTTPHeaders({
      "Accept-Language": "en-US,en;q=0.9",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    });

    // Navigate to the Catholic schools directory
    await page.goto("https://www.gscs.ca/page/63/find-a-school", {
      waitUntil: "networkidle0",
      timeout: 60000,
    });

    // Wait for the school list to load
    await page.waitForSelector(".cifs_listview table.table", {
      timeout: 30000,
    });

    // Get all school links
    const links = await page.$$(".cifs_listview table.table a");

    for (const link of links) {
      try {
        // Get school name and URL
        const name = await page.evaluate((el) => el.textContent.trim(), link);
        const url = await page.evaluate((el) => el.href, link);

        // Skip if not a school link
        if (!name || !url) continue;

        // Determine if it's a high school or French immersion
        const isHighSchool = name.toLowerCase().includes("high school");
        const isFrenchImmersion = name
          .toLowerCase()
          .includes("french immersion");

        // Visit the school's page
        await page.goto(url, { waitUntil: "networkidle0", timeout: 30000 });
        await delay(1000); // Add delay between requests

        // Get address and phone
        let address = "";
        let phone = "";
        let email = "";

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

          const emailElement = await page.$(".email");
          if (emailElement) {
            email = await page.evaluate(
              (el) => el.textContent.trim(),
              emailElement
            );
          }
        } catch (error) {
          console.log(`Error getting details for ${name}: ${error.message}`);
        }

        schools.push({
          name,
          type: isHighSchool ? "High School" : "Elementary",
          isFrenchImmersion,
          address,
          phone,
          email,
          url,
          principal: "", // Leave empty as requested
          superintendent: "", // Leave empty as requested
        });

        // Go back to the main page
        await page.goto("https://www.gscs.ca/page/63/find-a-school", {
          waitUntil: "networkidle0",
          timeout: 30000,
        });
        await delay(1000); // Add delay between requests
      } catch (error) {
        console.log(`Error processing school link: ${error.message}`);
        continue;
      }
    }
  } catch (error) {
    console.log(`Error scraping Catholic schools: ${error.message}`);
  }

  return schools;
}

async function scrapePrivateSchools(page, existingSchools) {
  console.log("Scraping private schools...");
  const schools = [];

  try {
    // Set user agent and other headers
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    );
    await page.setExtraHTTPHeaders({
      "Accept-Language": "en-US,en;q=0.9",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    });

    // Navigate to Google
    await page.goto(
      "https://www.google.com/search?q=private+schools+in+saskatoon",
      {
        waitUntil: "networkidle0",
        timeout: 30000,
      }
    );

    // Wait for search results
    await page.waitForSelector("div.g", { timeout: 30000 });

    // Get all search results
    const results = await page.$$("div.g");

    for (const result of results) {
      try {
        const title = await page.evaluate(
          (el) => el.querySelector("h3")?.textContent.trim(),
          result
        );
        const url = await page.evaluate(
          (el) => el.querySelector("a")?.href,
          result
        );

        if (!title || !url) continue;

        // Check if it's a school and not already in our list
        if (
          title.toLowerCase().includes("school") &&
          !existingSchools.some(
            (school) => school.name.toLowerCase() === title.toLowerCase()
          )
        ) {
          // Visit the school's website
          await page.goto(url, { waitUntil: "networkidle0", timeout: 30000 });
          await delay(1000); // Add delay between requests

          // Try to get address and phone
          let address = "";
          let phone = "";

          try {
            // Look for common patterns in the page content
            const content = await page.content();
            const addressMatch = content.match(
              /\d+\s+[A-Za-z\s,]+(?:Street|Avenue|Road|Drive|Boulevard|Lane|Way|Place|Court|Crescent|Circle|Terrace|Trail|Parkway|Square|Plaza|Heights|Gardens|Manor|Village|Estates|Hills|Valley|Meadows|Woods|Grove|Haven|Point|Bay|Harbour|Cove|Creek|River|Lake|Mountain|Ridge|Summit|Peak|View|Vista|Heights|Gardens|Manor|Village|Estates|Hills|Valley|Meadows|Woods|Grove|Haven|Point|Bay|Harbour|Cove|Creek|River|Lake|Mountain|Ridge|Summit|Peak|View|Vista)/i
            );
            if (addressMatch) {
              address = addressMatch[0].trim();
            }

            const phoneMatch = content.match(
              /\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/
            );
            if (phoneMatch) {
              phone = phoneMatch[0].trim();
            }
          } catch (error) {
            console.log(`Error getting details for ${title}: ${error.message}`);
          }

          schools.push({
            name: title,
            type: "Private",
            isFrenchImmersion: false,
            address,
            phone,
            email: "", // Leave empty as requested
            url,
            principal: "", // Leave empty as requested
            superintendent: "", // Leave empty as requested
          });
        }
      } catch (error) {
        console.log(`Error processing search result: ${error.message}`);
        continue;
      }
    }
  } catch (error) {
    console.log(`Error scraping private schools: ${error.message}`);
  }

  return schools;
}

async function main() {
  console.log("Starting to scrape school data...");
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  try {
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(60000);

    // Only scrape public schools for now
    const publicSchools = await scrapePublicSchools(page);
    // const catholicSchools = await scrapeCatholicSchools(page);
    // const privateSchools = await scrapePrivateSchools(page, publicSchools);

    const allSchools = [...publicSchools /*, ...catholicSchools, ...privateSchools */];

    // Save to Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Schools");

    // Add headers
    worksheet.columns = [
      { header: "Name", key: "name", width: 30 },
      { header: "Type", key: "type", width: 15 },
      { header: "Category", key: "category", width: 15 },
      { header: "French Status", key: "frenchStatus", width: 15 },
      { header: "Address", key: "address", width: 40 },
      { header: "Phone", key: "phone", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "URL", key: "url", width: 40 },
      { header: "Principal", key: "principal", width: 30 },
      { header: "Superintendent", key: "superintendent", width: 30 }
    ];

    // Add rows
    worksheet.addRows(allSchools);

    // Save the file
    const outputPath = "schools.xlsx";
    await workbook.xlsx.writeFile(outputPath);

    console.log(`Successfully scraped ${allSchools.length} schools`);
    console.log(`Data saved to ${process.cwd()}/${outputPath}`);
  } catch (error) {
    console.error("Error:", error);
  } finally {
    await browser.close();
    process.exit(0);
  }
}

// Run the scraper
main();
