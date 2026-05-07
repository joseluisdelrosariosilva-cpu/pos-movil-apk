import { abrirExcel, guardarExcel, obtenerHojaPorNombre } from "../utils/excelHelper.js";

const obtenerPrimeraFilaVacia = (hoja, columnaInicio = "A") => {
  let fila = 2;
  while (hoja.cell(`${columnaInicio}${fila}`).value()) {
    fila++;
  }
  return fila;
};

const sumarStockProducto = (workbook, codigoProducto, cantidadSumar) => {
  const hojaProductos = workbook.sheet("Productos");
  if (!hojaProductos) {
    throw new Error('No se encontró la hoja "Productos"');
  }

  let fila = 2;
  while (true) {
    const codigoCelda = hojaProductos.cell(`A${fila}`).value();
    if (!codigoCelda) break;

    if (codigoCelda.toString() === codigoProducto.toString()) {
      const stockActual = Number(hojaProductos.cell(`C${fila}`).value() || 0);
      const nuevoStock = stockActual + Number(cantidadSumar || 0);
      hojaProductos.cell(`C${fila}`).value(nuevoStock);

      return {
        fila,
        stockActual,
        nuevoStock,
      };
    }

    fila++;
  }

  throw new Error(`Producto ${codigoProducto} no encontrado en hoja Productos`);
};

export const sincronizarAbastecimientos = async (req, res) => {
  try {
    const { abastecimientos } = req.body;

    if (!abastecimientos?.length) {
      return res.status(400).json({
        success: false,
        error: "NO_HAY_ABASTECIMIENTOS",
        mensaje: "No se recibieron abastecimientos para sincronizar",
      });
    }

    const { workbook } = await abrirExcel();
    const hojaAbastecimiento = obtenerHojaPorNombre(workbook, "Abastecimiento");

    let sincronizadas = 0;
    const detalles = [];

    for (const abastecimiento of abastecimientos) {
      const codigo = abastecimiento.codigo;
      const cantidad = Number(abastecimiento.cantidad || 0);
      const fecha = abastecimiento.fechaHora || new Date().toISOString();

      if (!codigo || !cantidad) {
        continue;
      }

      const filaAbastecimiento = obtenerPrimeraFilaVacia(hojaAbastecimiento, "A");
      hojaAbastecimiento.cell(`A${filaAbastecimiento}`).value(codigo);
      hojaAbastecimiento.cell(`B${filaAbastecimiento}`).value(cantidad);
      hojaAbastecimiento.cell(`C${filaAbastecimiento}`).value(fecha);

      const stock = sumarStockProducto(workbook, codigo, cantidad);

      detalles.push({
        codigo,
        cantidad,
        filaAbastecimiento,
        filaProductos: stock.fila,
        stockAntes: stock.stockActual,
        stockDespues: stock.nuevoStock,
      });

      sincronizadas++;
    }

    await guardarExcel(workbook);

    return res.json({
      success: true,
      sincronizadas,
      detalles,
    });
  } catch (error) {
    console.error("❌ Error sincronizando abastecimientos:", error.message);
    return res.status(500).json({
      success: false,
      error: "ERROR_SINCRONIZACION_ABASTECIMIENTOS",
      mensaje: error.message,
    });
  }
};
