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
    const headers = [
      "Art",
      "Descripción",
      "%Iva",
      "Cantidad",
      "Precio por Unidad",
      "Precio Total",
    ];
    worksheet.addRow(headers);

    // Agregar bordes a los encabezados
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Separar el texto en líneas
    const rows = pdfData.text.split("\n");
    let currentRubro = "";

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
        worksheet.addRow([`Rubro: ${currentRubro}`, "", "", "", "", ""]);
        return;
      }

      // Verificar si la línea tiene un producto y extraer cantidad y precio
      const productMatch = row.match(/^(.+?)\s+([\d,.]+)\s+(\d+)$/);
      if (productMatch) {
        const descripcion = productMatch[1].trim(); // Descripción del artículo (Descripción)
        const totalPriceStr = productMatch[2].trim(); // Precio total como string
        const codigo = productMatch[3].trim(); // Código de artículo (Art)

        // Extraer la cantidad de unidades (número antes de "X")
        const quantityMatch = row.match(/X(\d+)U/);
        const cantidad = quantityMatch ? parseInt(quantityMatch[1]) : 1; // Por defecto es 1 si no se encuentra

        // Convertir el precio total a número (eliminando el signo de pesos)
        const totalPrice = parseFloat(totalPriceStr);

        // Calcular el precio por unidad
        const precioPorUnidad =
          cantidad > 0 ? (totalPrice / cantidad).toFixed(4) : 0;

        // El %Iva es fijo en este caso
        const iva = "0";

        // Agregar una nueva fila al Excel con los datos del producto
        const newRow = worksheet.addRow([
          codigo,
          descripcion,
          iva,
          cantidad,
          ` $${precioPorUnidad}`,
          ` $${totalPrice.toFixed(4)}`, // Formatear el precio total
        ]);

        // Aplicar bordes a la fila del producto
        newRow.eachCell((cell) => {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });

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

// Asegúrate de declarar la variable port antes de utilizarla
const port = process.env.PORT || 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
