const express = require("express");
const cors = require("cors");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const upload = multer({ dest: "uploads/" });

// Configuración de CORS
app.use(cors({ origin: "https://convert-pdf-to-excel.vercel.app" }));

app.use(express.json());

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {
    const filePath = req.file.path;
    const pdfBuffer = fs.readFileSync(filePath);
    const pdfData = await pdfParse(pdfBuffer);

    // Crear un nuevo archivo Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Datos PDF");

    // Encabezados
    worksheet.addRow(["Art", "Descripción", "%Iva", "Precio"]);

    // Separar el texto en líneas
    const rows = pdfData.text.split("\n");
    let currentRubro = "";

    rows.forEach((row) => {
      // Limpiar la línea de texto
      row = row.trim();

      // Ignorar líneas vacías o sin valor
      if (row === "") return;

      // Verificar si la línea es un nuevo Rubro (asumiendo que comienza con "Rubro:" y tiene algún contenido después)
      const rubroMatch = row.match(/^Rubro:\s*(.+)$/);
      if (rubroMatch) {
        currentRubro = rubroMatch[1].trim();
        // Agregar una nueva fila al Excel con el nombre del Rubro
        worksheet.addRow([`Rubro: ${currentRubro}`, "", "", ""]);
        return;
      }

      // Verificar si la línea tiene un producto basado en un patrón típico (número seguido de texto y un precio al final)
      const productMatch = row.match(/^(\d+)\s+([A-Za-z\s]+)\s+([\d,.]+)$/);
      if (productMatch) {
        const codigo = productMatch[1].trim(); // Código de artículo (Art)
        const descripcion = productMatch[2].trim(); // Descripción del artículo (Descripción)
        const precio = productMatch[3]
          .trim()
          .replace(".", "")
          .replace(",", "."); // Precio (Lista 4)

        // El %Iva es fijo en este caso
        const iva = "0";

        // Agregar una nueva fila al Excel con los datos del producto
        worksheet.addRow([codigo, descripcion, iva, precio]);
        return;
      }

      // Si la línea no coincide con un rubro o producto, simplemente la ignoramos.
    });

    // Enviar el Excel al cliente
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=converted.xlsx");

    await workbook.xlsx.write(res);

    // Eliminar el archivo PDF subido del servidor
    fs.unlinkSync(filePath);
  } catch (error) {
    console.error("Error processing the PDF:", error);
    res.status(500).send("Error processing the file");
  }
});

// Asegúrate de declarar la variable `port` antes de utilizarla
const port = process.env.PORT || 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
