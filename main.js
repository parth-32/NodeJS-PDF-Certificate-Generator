import { PDFDocument, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
// import fs from 'fs'
import { promises as fs } from 'fs';
import XLSX from 'xlsx';

// Read Excel Sheet
var workbook = XLSX.readFile('names.xlsx');
var sheet_name_list = workbook.SheetNames;
var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

console.log('****** Excel Sheet Data ****** ', xlData);

xlData = xlData.map((ele, i) => ({ Name: ele.Name, Number: i + 1 }));

try {
  import('babel-core/register')
    .then(() => console.log('Loaded..'))
    .catch((e) => console.log('***'));
  import('babel-polyfill')
    .then(() => console.log('Loaded..'))
    .catch((e) => console.log('***'));

  const urlFont = 'fonts/Fasthand/Fasthand-Regular.ttf';

  for (const name of xlData) {
    console.log('start -> ', name.Number, name.Name);

    // Create Custom Font Bytes
    const fontBytes = await fs.readFile(urlFont);

    delay(2000);

    // Create template PDF Bytes from sample input PDF
    const existingPdfBytes = await fs.readFile('input.pdf');

    // Load a PDFDocument from the existing PDF bytes
    const pdfDoc = await PDFDocument.load(existingPdfBytes);

    // Register font kit in PDF
    pdfDoc.registerFontkit(fontkit);

    // Embed font
    const customFont = await pdfDoc.embedFont(fontBytes);

    // Get the first page of the document
    const pages = pdfDoc.getPages();

    // Specify Page Number to add text
    const firstPage = pages[0];

    // Get the width and height of the first page
    // const { width, height } = firstPage.getSize()

    // ****** Add Text to PDF ********
    firstPage.drawText(name.Name, {
      x: 200,
      y: 450,
      size: 45,
      font: customFont,
      color: rgb(0.95, 0.2, 0.1),
    });

    // Serialize the PDFDocument to bytes (a Uint8Array)
    const pdfBytes = await pdfDoc.save();

    // Generate and Save Output PDF
    fs.writeFile(`${name.Name}.pdf`, pdfBytes, (err) => {
      if (err) {
        console.log('Error while saving output PDF: ', err);
      }
    });

    console.log('Completed -> ', name.Number, name.Name, '\n');
  }
} catch (error) {
  console.log('ERROR: ', error);
}

function delay(milliseconds) {
  const date = Date.now();
  let currentDate = null;
  do {
    currentDate = Date.now();
  } while (currentDate - date < milliseconds);
}
