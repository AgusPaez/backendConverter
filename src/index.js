const express = require("express");
const cors = require("cors"); // Esta es la única declaración que necesitas
const multer = require("multer");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const upload = multer({ dest: "uploads/" });

// Configuración de CORS
app.use(cors({ origin: "https://convert-pdf-to-excel.vercel.app/" }));

app.use(express.json());

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {
    const filePath = req.file.path;
    const pdfBuffer = fs.readFileSync(filePath);
    const pdfData = await pdfParse(pdfBuffer);

    // Crear un nuevo archivo Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Datos PDF");

    // Convertir el texto a filas en Excel
    const rows = pdfData.text.split("\n");
    rows.forEach((row) => {
      worksheet.addRow([row]);
    });

    // Enviar el Excel al cliente
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=converted.xlsx");

    await workbook.xlsx.write(res);
    fs.unlinkSync(filePath); // Eliminar el archivo PDF subido del servidor
  } catch (error) {
    console.error("Error processing the PDF:", error);
    res.status(500).send("Error processing the file");
  }
});

// Asegúrate de declarar la variable `port` antes de utilizarla
const port = process.env.PORT || 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
