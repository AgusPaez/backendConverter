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
    let lastDescription = "";

    rows.forEach((row) => {
      // Limpiar la línea de texto
      row = row.trim();

      // Ignorar líneas vacías o sin valor
      if (row === "") return;

      // Verificar si la línea es un nuevo Rubro
      const rubroMatch = row.match(/^Rubro:\s*(.+)$/);
      if (rubroMatch) {
        currentRubro = rubroMatch[1].trim();
        // Agregar una nueva fila al Excel con el nombre del Rubro
        worksheet.addRow([`Rubro: ${currentRubro}`, "", "", ""]);
        return;
      }

      // Verificar si la línea tiene un producto
      const productMatch = row.match(/^(\d+)\s+(.+?)\s+([\d,.]+)$/);
      if (productMatch) {
        const precio = productMatch[1].trim(); // Código de artículo (Art)
        const descripcionCompleta = productMatch[2].trim();
        const index = descripcionCompleta.search(/ \d/);

        // Si se encuentra un espacio seguido de un número, corta la cadena
        const descripcion =
          index !== -1
            ? descripcionCompleta.slice(0, index)
            : descripcionCompleta;

        // Descripción del artículo (Descripción)
        const codigo = productMatch[3]
          .trim()
          .replace(".", "")
          .replace(",", "."); // Precio (Lista 4)

        // El %Iva es fijo en este caso
        const iva = "0";

        // Agregar una nueva fila al Excel con los datos del producto
        worksheet.addRow([codigo, descripcion, iva, precio]);
        lastDescription = ""; // Reiniciar la última descripción
        return;
      }

      // Si la línea no coincide con un rubro o producto, verificar si es una descripción
      const descriptionMatch = row.match(/^(.+?)\s+([\d,.]+)$/);
      if (descriptionMatch && lastDescription) {
        // Se considera una descripción adicional para el último producto
        const precio = descriptionMatch[2]
          .trim()
          .replace(".", "")
          .replace(",", ".");

        // Agregar una nueva fila con la descripción y precio
        worksheet.addRow([
          lastDescription,
          row.replace(/([\d,.]+)$/, "").trim(),
          "0",
          precio,
        ]);
        lastDescription = ""; // Reiniciar
        return;
      }

      // Guardar la descripción para la próxima coincidencia de precio
      lastDescription = row;
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
