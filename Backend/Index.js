const puppeteer = require("puppeteer");
const path = require('path');
const fs = require('fs');

// Helper function to wait for a download to complete
async function waitForDownload(downloadPath, extension, timeout = 60000) {
    console.log(`Waiting for .${extension} file to download...`);
    const maxAttempts = timeout / 1000;
    for (let attempts = 0; attempts < maxAttempts; attempts++) {
        const files = fs.readdirSync(downloadPath);
        // Find a file with the correct extension that isn't a temporary download file
        const fileName = files.find(file => file.endsWith(`.${extension}`) && !file.endsWith('.crdownload'));

        if (fileName) {
            console.log(`.${extension} download complete: ${fileName}`);
            return fileName;
        }
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    throw new Error(`Download of .${extension} file did not complete within ${timeout / 1000} seconds.`);
}

(async () => {
    // Set up a 'downloads' directory for the files
    const downloadPath = path.resolve('./downloads');
    if (!fs.existsSync(downloadPath)) {
        fs.mkdirSync(downloadPath, { recursive: true });
    }

    console.log('Launching browser...');
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // Tell Puppeteer to allow and save downloads to our specified directory
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath,
    });

    console.log('Navigating to Angular app...');
    await page.goto("http://localhost:4200/", { waitUntil: "networkidle0" });

    try {
        // --- EXPORT EXCEL ---
        console.log('Clicking "Export to Excel"...');
        await page.waitForSelector('#export-excel-button');
        await page.click('#export-excel-button');
        await waitForDownload(downloadPath, 'xlsx');

        // --- EXPORT PDF ---
        console.log('Clicking "Export to PDF"...');
        await page.waitForSelector('#export-pdf-button');
        await page.click('#export-pdf-button');
        await waitForDownload(downloadPath, 'pdf');

        // --- EXPORT PPT ---
        console.log('Clicking "Export to PPT"...');
        await page.waitForSelector('#export-ppt-button');
        await page.click('#export-ppt-button');
        await waitForDownload(downloadPath, 'pptx');

        console.log('All exports completed successfully!');

    } catch (error) {
        console.error('An error occurred during the export process:', error);
    } finally {
        await browser.close();
        console.log('Browser closed.');
    }
})();











// require('dotenv').config(); // Loads variables from .env file into process.env
// const puppeteer = require("puppeteer");
// const XLSX = require("xlsx");
// const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
// const PptxGenJS = require("pptxgenjs");
// const nodemailer = require("nodemailer");

// (async () => {
//     // --- NODEMAILER CONFIGURATION ---
//     if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
//         console.error("Error: EMAIL_USER and EMAIL_PASS must be set in the .env file.");
//         return; // Exit the script if credentials are not found
//     }

//     const transporter = nodemailer.createTransport({
//         service: 'gmail',
//         auth: {
//             user: process.env.EMAIL_USER, // Loaded from .env
//             pass: process.env.EMAIL_PASS     // Loaded from .env
//         }
//     });

//     // Launch headless browser
//     const browser = await puppeteer.launch({ headless: true });
//     const page = await browser.newPage();

//     // 1. Open Angular app
//     await page.goto("http://localhost:4200/", { waitUntil: "networkidle0" });

//     // 2. Wait for AG Grid to load rows
//     await page.waitForSelector(".ag-center-cols-container .ag-row");

//     // Function to extract data from the current page
//     async function extractRows() {
//         return page.evaluate(() => {
//             const rows = [];
//             document.querySelectorAll(".ag-center-cols-container .ag-row").forEach(row => {
//                 const cells = Array.from(row.querySelectorAll(".ag-cell")).map(cell => cell.innerText.trim());
//                 rows.push(cells);
//             });
//             return rows;
//         });
//     }

//     const allData = [];
//     let currentPage = 1;

//     while (true) {
//         console.log(`Extracting page ${currentPage}...`);

//         // Get the unique row-id of the first row to detect a page change.
//         const firstRowId = await page.evaluate(() => {
//             const firstRow = document.querySelector('.ag-center-cols-container .ag-row');
//             return firstRow ? firstRow.getAttribute('row-id') : null;
//         });

//         if (firstRowId === null) {
//             console.log("Could not find any data rows. Ending pagination.");
//             break;
//         }

//         const rows = await extractRows();
//         allData.push(...rows);

//         const hasNextPage = await page.evaluate(() => {
//             const nextButton = Array.from(document.querySelectorAll('button')).find(btn => btn.textContent.trim() === 'Next');
//             if (nextButton && !nextButton.disabled) {
//                 nextButton.click();
//                 return true;
//             }
//             return false;
//         });

//         if (!hasNextPage) {
//             break; // Exit if no "Next" button or it's disabled
//         }

//         // --- FINAL PAGINATION FIX ---
//         // Wait for the row-id of the first row to be different. This is the most reliable method.
//         try {
//             await page.waitForFunction(
//                 (previousRowId) => {
//                     const currentRow = document.querySelector('.ag-center-cols-container .ag-row');
//                     // Check succeeds when the new row exists and its ID is different from the old one.
//                     return currentRow && currentRow.getAttribute('row-id') !== previousRowId;
//                 },
//                 { timeout: 10000 },
//                 firstRowId // Pass the old ID in for comparison
//             );
//         } catch (e) {
//             console.log("The grid's row-id did not change after click. Assuming end of pagination.");
//             break;
//         }

//         currentPage++;
//     }

//     await browser.close();
//     console.log(`Total rows extracted: ${allData.length}`);

//     const headers = [['Name', 'Email', 'Country', 'Phone']];
//     const dataWithHeaders = headers.concat(allData);

//     // 3. Generate Excel buffer
//     const wb = XLSX.utils.book_new();
//     const ws = XLSX.utils.aoa_to_sheet(dataWithHeaders);
//     XLSX.utils.book_append_sheet(wb, ws, "AG Grid Data");
//     const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
//     console.log(`Excel file generated in memory. Size: ${excelBuffer.length} bytes.`);

//     // 4. Generate PDF buffer
//     const pdfDoc = await PDFDocument.create();
//     let pagePDF = pdfDoc.addPage();
//     const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
//     let y = 750;
//     dataWithHeaders.forEach(row => {
//         if (y < 40) {
//             pagePDF = pdfDoc.addPage();
//             y = 750;
//         }
//         pagePDF.drawText(row.join(" | "), { x: 20, y, size: 10, font, color: rgb(0, 0, 0) });
//         y -= 15;
//     });
//     const pdfBuffer = await pdfDoc.save();
//     console.log(`PDF file generated in memory. Size: ${pdfBuffer.length} bytes.`);

//     // 5. Generate PPT buffer
//     let pptx = new PptxGenJS();
//     let slide = pptx.addSlide();
//     let text = dataWithHeaders.map(row => row.join(" | ")).join("\n");
//     slide.addText(text, { x: 0.5, y: 0.5, w: '90%', h: '90%', fontSize: 10, color: "000000" });
//     const pptxBase64 = await pptx.write({ outputType: 'base64' });
//     const pptxBuffer = Buffer.from(pptxBase64, 'base64');
//     console.log(`PPTX file generated in memory. Size: ${pptxBuffer.length} bytes.`);

//     // --- SEND EMAIL WITH ATTACHMENTS ---
//     const mailOptions = {
//         from: process.env.EMAIL_USER, // Sender address from .env file
//         to: 'recipient-email@example.com', // List of receivers
//         subject: 'AG Grid Data Export', // Subject line
//         text: 'Attached are the data exports from the AG Grid.', // Plain text body
//         attachments: [
//             {
//                 filename: 'aggrid-data.xlsx',
//                 content: excelBuffer,
//                 contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//             },
//             {
//                 filename: 'aggrid-data.pdf',
//                 content: pdfBuffer,
//                 contentType: 'application/pdf',
//             },
//             {
//                 filename: 'aggrid-data.pptx',
//                 content: pptxBuffer,
//                 contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
//             }
//         ]
//     };

//     try {
//         let info = await transporter.sendMail(mailOptions);
//         console.log('Email sent successfully! Message ID: ' + info.messageId);
//     } catch (error) { 
//         console.error('Error sending email:', error);
//     }
// })();
