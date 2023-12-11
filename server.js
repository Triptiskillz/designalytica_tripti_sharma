const express = require("express");
const bodyParser = require("body-parser");
const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");
const path = require("path");
const XlsxPopulate = require("xlsx-populate");

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use(express.static(path.join(__dirname, "public")));
let result = 0;

// Handle form submission
app.post("/calculate", (req, res) => {
  // Parse input numbers from the request
  const num1 = parseFloat(req.body.num1);
  const num2 = parseFloat(req.body.num2);
  result = num1 + num2;

  // Write data to Excel sheet
  writeToExcel(num1, num2, result);

  // Respond with the result
  res.json({ result });
});

// Handle PDF generation
app.get("/print", async (req, res) => {
  try {
    // Generate PDF from Excel sheet
    await generatePDF();

    // Send a JSON response indicating success
    res.json({ success: true, message: "PDF generated successfully!" });
  } catch (error) {
    console.error("Error while printing to PDF:", error);

    // Send a JSON response indicating failure
    res
      .status(500)
      .json({ success: false, message: "Error while printing to PDF" });
  }
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});

// Function to write data to Excel sheet
async function writeToExcel(num1, num2, result) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sheet1");
  sheet.addRow(["Number 1", "Number 2", "Result"]);
  sheet.addRow([num1, num2, result]);

  // Save the workbook to a file
  await workbook.xlsx.writeFile("result.xlsx");
}

// Function to generate PDF from Excel sheet
async function generatePDF() {
  // Launch Puppeteer
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  // Load the Excel file and convert it to HTML
  const excelFilePath = path.join(__dirname, "result.xlsx");

  try {
    // Read Excel file using XlsxPopulate
    const workbook = await XlsxPopulate.fromFileAsync(excelFilePath);
    const sheetData = workbook.sheet(0).usedRange().value();

    // Generate HTML content from sheet data
    const htmlContent = generateHtmlFromSheetData(sheetData);

    // Set the HTML content on the page
    await page.setContent(htmlContent);

    // Generate PDF
    await page.pdf({ path: "result.pdf", format: "A4" });
  } catch (error) {
    console.error(`Error reading or converting Excel file: ${error.message}`);
  } finally {
    // Close the browser
    await browser.close();
  }
}

// Function to generate HTML from sheet data
function generateHtmlFromSheetData(sheetData) {
  // Implement your logic to generate HTML from sheetData
  // You can use a template engine or manually construct HTML based on your needs.
  // Example: Constructing a simple table
  const tableRows = sheetData
    .map((row) => `<tr>${row.map((cell) => `<td>${cell}</td>`).join("")}</tr>`)
    .join("");
  const htmlContent = `<table>${tableRows}</table>`;
  return htmlContent;
}
