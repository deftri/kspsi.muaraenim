const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx'); // Import xlsx
const PDFDocument = require('pdfkit'); // Import pdfkit
const app = express();
const PORT = 3000;

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json()); // Middleware to parse JSON request body

// Route untuk menambahkan informasi perusahaan
app.post('/add-company', (req, res) => {
    const { companyName, employeeName, position, phone, number } = req.body;

    // Log the received data for debugging
    console.log('Received data:', req.body);

    // Tentukan path untuk file Excel
    const filePath = path.join(__dirname, 'public', 'assets', 'companies.xlsx');
    let workbook;

    // Cek apakah file sudah ada
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath); // Baca file jika sudah ada
    } else {
        workbook = xlsx.utils.book_new(); // Buat workbook baru
    }

    // Siapkan data untuk ditulis
    const sheetData = [
        { Company: companyName, Number: number, Name: employeeName, Position: position, Phone: phone }
    ];

    // Cek apakah sheet 'Companies' sudah ada
    const sheetName = 'Companies';
    if (workbook.Sheets[sheetName]) {
        // Ambil data yang sudah ada
        const existingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        const updatedData = existingData.concat(sheetData); // Gabungkan data lama dan baru
        const updatedSheet = xlsx.utils.json_to_sheet(updatedData);
        workbook.Sheets[sheetName] = updatedSheet; // Update sheet yang ada
    } else {
        const newSheet = xlsx.utils.json_to_sheet(sheetData);
        xlsx.utils.book_append_sheet(workbook, newSheet, sheetName); // Tambahkan sheet baru
    }

    // Tulis ke file Excel
    try {
        xlsx.writeFile(workbook, filePath);
        console.log('Data written to file successfully');
        res.sendStatus(200); // Mengirimkan status berhasil
    } catch (error) {
        console.error('Error writing to file:', error);
        res.status(500).send('Error writing to file'); // Mengirimkan status error
    }
});

// Route untuk mengunduh data dalam format Excel
app.get('/download/excel', (req, res) => {
    const filePath = path.join(__dirname, 'public', 'assets', 'companies.xlsx');
    res.download(filePath, 'companies.xlsx', (err) => {
        if (err) {
            console.error('Error downloading the file:', err);
            res.status(500).send('Error downloading file');
        }
    });
});

// Route untuk mengunduh data dalam format PDF
app.get('/download/pdf', (req, res) => {
    const filePath = path.join(__dirname, 'public', 'assets', 'companies.xlsx');
    if (!fs.existsSync(filePath)) {
        return res.status(404).send('Excel file not found');
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = 'Companies';
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Membuat PDF
    const doc = new PDFDocument();
    let pdfPath = path.join(__dirname, 'public', 'assets', 'companies.pdf');

    res.setHeader('Content-disposition', 'attachment; filename=companies.pdf');
    res.setHeader('Content-type', 'application/pdf');

    doc.pipe(fs.createWriteStream(pdfPath));
    doc.fontSize(20).text('Data Perusahaan', { align: 'center' });
    doc.moveDown();

    // Menulis data ke PDF
    data.forEach((item) => {
        doc.fontSize(12).text(`Company: ${item.Company}`);
        doc.text(`Number: ${item.Number}`);
        doc.text(`Name: ${item.Name}`);
        doc.text(`Position: ${item.Position}`);
        doc.text(`Phone: ${item.Phone}`);
        doc.moveDown();
    });

    doc.end();

    // Menyelesaikan pengunduhan PDF
    doc.pipe(res);
});

// Route untuk halaman pengunjung
app.get('/pages', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'pages', 'pages.html'));
});

// Menjalankan server
app.listen(PORT, () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
});
