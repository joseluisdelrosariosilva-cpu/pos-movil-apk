import { abrirExcel, guardarExcel, obtenerHojaPorNombre, actualizarStock } from "../utils/excelHelper.js";

// ============================================
// POST /api/mermas
// Recibe un array de mermas y las escribe en la hoja "Merma"
// También actualiza el stock en la hoja "Productos"
// ============================================
export const registrarMermas = async (req, res) => {
  try {
    const { mermas } = req.body;
    
    if (!mermas?.length) {
      return res.status(400).json({ error: "Debe incluir al menos una merma" });
    }
    
    const { workbook } = await abrirExcel();
    
    // 1. Escribir en hoja Merma
    const hojaMerma = obtenerHojaPorNombre(workbook, "Merma");
    
    // Encontrar primera fila vacía en columna A (después de encabezados)
    let fila = 1;
    // Si hay encabezados en A1, empezar desde fila 2
    if (hojaMerma.cell(`A1`).value()) {
      fila = 2;
    }
    while (hojaMerma.cell(`A${fila}`).value()) {
      fila++;
    }
    
    // Escribir cada merma: columna A = código, B = nombre, C = cantidad
    for (const merma of mermas) {
      hojaMerma.cell(`A${fila}`).value(merma.codigo);
      hojaMerma.cell(`B${fila}`).value(merma.nombre);
      hojaMerma.cell(`C${fila}`).value(merma.cantidad);
      fila++;
    }
    
    // 2. Actualizar stock en hoja Productos (restar cantidades)
    for (const merma of mermas) {
      await actualizarStock(workbook, merma.codigo, merma.cantidad);
    }
    
    // 3. Guardar cambios
    await guardarExcel(workbook);
    
    console.log(`✅ ${mermas.length} mermas registradas en Excel`);
    
    res.json({
      success: true,
      sincronizadas: mermas.length,
    });
    
  } catch (error) {
    console.error("❌ Error registrando mermas:", error.message);
    res.status(500).json({
      error: "Error al registrar mermas",
      detalle: error.message,
    });
  }
};
