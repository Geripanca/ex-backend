const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

// Setup Express app
const app = express();
const PORT = process.env.PORT || 4000; // Atur port yang akan digunakan

// Buat direktori output jika belum ada
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

// Setup multer untuk file upload
const upload = multer({ dest: "uploads/" });

// Fungsi untuk mengekstrak informasi dari teks PDF
// Fungsi untuk mengekstrak informasi dari teks PDF
const extractInfo = (text) => {
  const regexPatterns = {
    invoiceNumber: /(\d{2}[A-Z]+\d{6})/,
    billTo: /Bill\s+To\s+([\s\S]*?)(?=Page|Invoice)/,
    subAccount: /\b\d{4}\b/, // Menangkap setiap angka 4 digit, termasuk sub account atau tahun
    purchaseOrder: /Purchase Order\s*:?\s*(\S+)/,
    itemNumber:
      /(\d+)\n(\d+\.\d+)\n\d+\nYes\n[0-9,.]+\n[0-9,.]+\n(.+?)(?=\n\d+|\n\* \* \* D U P L I C A T E \* \* \*|\Z)/gs, // Item Number regex yang disesuaikan
    invoiceDate: /(\d{1,2}\/\d{1,2}\/\d{4})/,
    total: /Total\s*:?\s*([-]?\d+(?:,\d{3})*)/g,
  };

  const invoiceNumber = text.match(regexPatterns.invoiceNumber)?.[1] || null;
  const finalResult = text.match(regexPatterns.billTo)?.[1] || null;
  const subAccount = text.match(regexPatterns.subAccount)?.[0] || null; // Mengambil angka 4 digit pertama
  const purchaseOrder = text.match(regexPatterns.purchaseOrder)?.[1] || null;

  // Menangkap item number dan deskripsinya
  const itemMatches = [...text.matchAll(regexPatterns.itemNumber)];
  const itemNumber =
    itemMatches
      .map((match) => {
        let description = match[3].trim().replace(/\n/g, " ");
        return `${match[1]}. ${description}`;
      })
      .join("\n") || null;

  const invoiceDate = text.match(regexPatterns.invoiceDate)?.[1] || null;
  const totalMatches = [...text.matchAll(regexPatterns.total)];
  const total = totalMatches[totalMatches.length - 1]?.[1] || null;

  return {
    Invoice: invoiceNumber,
    "Bill To": finalResult,
    "Sub Acct / Year": subAccount, // Bisa sub account atau tahun
    "Purchase Order": purchaseOrder,
    Description: itemNumber,
    "Invoice Date": invoiceDate,
    Total: total,
  };
};

// Route untuk file upload
app.post("/upload", upload.single("pdfFile"), async (req, res) => {
  if (!req.file || req.file.mimetype !== "application/pdf") {
    return res.status(400).json({ message: "Invalid file type" });
  }

  const filePath = req.file.path;
  const pdfData = fs.readFileSync(filePath);

  try {
    const data = await pdfParse(pdfData);
    const text = data.text;

    // Extract information from the text
    const extractedData = extractInfo(text);
    console.log("Extracted Text:", text); // Cetak hasil teks

    // Convert to Excel
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([extractedData]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    const outputFilename = path.join(
      outputDir,
      `${path.parse(req.file.originalname).name}.xlsx`
    );
    XLSX.writeFile(workbook, outputFilename);

    fs.unlinkSync(filePath); // Delete the uploaded file

    res.download(outputFilename, () => {
      fs.unlinkSync(outputFilename); // Delete the generated file after download
    });
  } catch (error) {
    console.error("Error processing PDF:", error);
    res.status(500).json({ message: "Error processing PDF" });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
