// ============================================
// CONTROLLER: Resumen de ventas por rango de fechas
// ============================================
// Endpoint: GET /api/resumen?desde=YYYY-MM-DD&hasta=YYYY-MM-DD
// Si no se envía desde, se usa el día actual
// Si no se envía hasta, se usa el mismo día que desde
// Devuelve: total vendido, ventas por producto, efectivo y transferencia
// ============================================

import { abrirExcel } from "../utils/excelHelper.js";

const HOJA_PENDIENTES = "Pendientes";

export const getResumen = async (req, res) => {
  try {
    const ahora = new Date();
    const hoy = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, "0")}-${String(ahora.getDate()).padStart(2, "0")}`;

    // Obtener rango desde query params (backward compatible: si no hay params, solo hoy)
    const desde = req.query.desde || hoy;
    const hasta = req.query.hasta || desde;

    // Abrir Excel
    const { workbook, hoja } = await abrirExcel();

    // Obtener rango de datos
    const rango = hoja.usedRange();
    const datos = rango ? rango.value() : [];

    if (datos.length < 2) {
      return res.json({
        desde: desde,
        hasta: hasta,
        totalIngresado: 0,
        efectivo: 0,
        transferencia: 0,
        productosVendidos: [],
      });
    }

    // Estructuras para calcular
    let totalIngresado = 0;
    let totalEfectivo = 0;
    let totalTransferencia = 0;
    const productosAgrupados = {};
    const facturasVistas = {};

    // Procesar cada fila (empezando desde índice 1,skip headers)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || !fila[0]) continue; // Skip empty rows

      const facturaId = fila[0];           // Columna A
      const fechaHora = fila[1];           // Columna B
      const codigoProducto = fila[2];    // Columna C
      const nombre = fila[3];             // Columna D
      const cantidad = fila[4] || 0;     // Columna E
      const precio = fila[5] || 0;       // Columna F
      const subtotal = fila[6] || 0;      // Columna G
      const efectivo = fila[7] || 0;    // Columna H
      const transferencia = fila[8] || 0; // Columna I
      const procesado = fila[9];         // Columna J

      const facturaKey = facturaId ? facturaId.toString() : `row_${i}`;

      // Solo ventas NO procesadas
      if (procesado === true || procesado === "VERDADERO" || procesado === "TRUE" || procesado === 1) {
        continue;
      }

      // Verificar si está dentro del rango de fechas
      const fechaVenta = fechaHora ? fechaHora.toString().split("T")[0] : "";
      if (fechaVenta < desde || fechaVenta > hasta) {
        continue;
      }

      const cantidadNum = Number(cantidad) || 0;
      const subtotalNum = Number(subtotal) || 0;
      const efectivoNum = Number(efectivo) || 0;
      const transferenciaNum = Number(transferencia) || 0;

      // Sumar totales
      totalIngresado += subtotalNum;

      // Efectivo/transferencia vienen repetidos por cada línea de la misma factura
      if (!facturasVistas[facturaKey]) {
        totalEfectivo += efectivoNum;
        totalTransferencia += transferenciaNum;
        facturasVistas[facturaKey] = true;
      }

      // Agrupar por producto
      const key = codigoProducto;
      if (!productosAgrupados[key]) {
        productosAgrupados[key] = {
          codigo: codigoProducto,
          nombre: nombre,
          cantidad: 0,
          total: 0,
        };
      }
      productosAgrupados[key].cantidad += cantidadNum;
      productosAgrupados[key].total += subtotalNum;
    }

    // Convertir a array y ordenar por cantidad
    const productosVendidos = Object.values(productosAgrupados).sort(
      (a, b) => b.cantidad - a.cantidad
    );

    // Responder
    res.json({
      desde: desde,
      hasta: hasta,
      totalIngresado: Math.round(totalIngresado * 1000) / 1000,
      efectivo: Math.round(totalEfectivo * 1000) / 1000,
      transferencia: Math.round(totalTransferencia * 1000) / 1000,
      cantidadVentas: Object.keys(facturasVistas).length,
      productosVendidos: productosVendidos.map((p) => ({
        codigo: p.codigo,
        nombre: p.nombre,
        cantidad: p.cantidad,
        total: Math.round(p.total * 1000) / 1000,
      })),
    });
  } catch (error) {
    console.error("❌ Error en getResumen:", error.message);
    res.status(500).json({
      error: "Error al obtener resumen",
      detalle: error.message,
    });
  }
};
