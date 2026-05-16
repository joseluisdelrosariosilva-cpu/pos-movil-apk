import { abrirExcel, guardarExcel, obtenerHojaPorNombre } from "../utils/excelHelper.js";

// ============================================
// POST /api/gastos
// Recibe un array de gastos y los escribe en la hoja "Gastos"
// Columnas: Fecha, Descripcion, Monto
// ============================================
export const registrarGastos = async (req, res) => {
  try {
    const { gastos } = req.body;

    if (!gastos?.length) {
      return res.status(400).json({ error: "Debe incluir al menos un gasto" });
    }

    const { workbook } = await abrirExcel();

    // Obtener o crear hoja Gastos
    let hojaGastos;
    try {
      hojaGastos = obtenerHojaPorNombre(workbook, "Gastos");
    } catch (e) {
      // Si no existe la hoja, crearla con encabezados
      hojaGastos = workbook.addSheet("Gastos");
      hojaGastos.cell("A1").value("Fecha");
      hojaGastos.cell("B1").value("Descripcion");
      hojaGastos.cell("C1").value("Monto");
    }

    // Encontrar primera fila vacía (después de encabezados)
    let fila = 2;
    while (hojaGastos.cell(`A${fila}`).value()) {
      fila++;
    }

    // Escribir cada gasto
    for (const gasto of gastos) {
      hojaGastos.cell(`A${fila}`).value(gasto.fecha);
      hojaGastos.cell(`B${fila}`).value(gasto.descripcion);
      hojaGastos.cell(`C${fila}`).value(Number(gasto.monto || 0));
      fila++;
    }

    // Guardar cambios
    await guardarExcel(workbook);

    console.log(`✅ ${gastos.length} gastos registrados en Excel (hoja Gastos)`);

    res.json({
      success: true,
      sincronizadas: gastos.length,
    });

  } catch (error) {
    console.error("❌ Error registrando gastos:", error.message);
    res.status(500).json({
      error: "Error al registrar gastos",
      detalle: error.message,
    });
  }
};
