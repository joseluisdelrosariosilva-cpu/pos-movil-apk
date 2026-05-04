import { abrirExcel, guardarExcel } from "../utils/excelHelper.js";

export const registrarMermas = async (req, res) => {
  try {
    const { mermas } = req.body;

    if (!Array.isArray(mermas) || mermas.length === 0) {
      return res.status(400).json({ error: "Debe incluir un array de mermas válido" });
    }

    // Abrir datos.xlsx
    const { workbook } = await abrirExcel();
    let mermaSheet = workbook.sheet("Merma");

    // Crear hoja si no existe
    if (!mermaSheet) {
      mermaSheet = workbook.addSheet("Merma");
      // Agregar cabeceras (Tabla3: Codigo, Producto, Cantidad)
      mermaSheet.cell(1, 1).value("Codigo");
      mermaSheet.cell(1, 2).value("Producto");
      mermaSheet.cell(1, 3).value("Cantidad");
    }

    // Obtener última fila
    let lastRow = 1;
    const usedRange = mermaSheet.usedRange();
    if (usedRange) {
      lastRow = usedRange.end().row;
    }

    // Escribir mermas
    for (const merma of mermas) {
      const { codigo, producto, cantidad } = merma;

      if (!codigo || !producto || !cantidad) {
        console.warn("⚠️ Merma inválida, saltando:", merma);
        continue;
      }

      lastRow += 1;
      mermaSheet.cell(lastRow, 1).value(codigo);
      mermaSheet.cell(lastRow, 2).value(producto);
      mermaSheet.cell(lastRow, 3).value(cantidad);
    }

    // Guardar cambios
    await guardarExcel(workbook);

    res.status(201).json({ success: true, count: mermas.length });
  } catch (error) {
    console.error("❌ Error en registrarMermas:", error.message);
    res.status(500).json({
      error: "Error al registrar mermas",
      detalle: error.message,
    });
  }
};
