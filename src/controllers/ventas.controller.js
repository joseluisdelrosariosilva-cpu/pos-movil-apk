import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import {
  abrirExcel,
  conExcelLock,
  generarFacturaIdExcel,
  escribirLineaVenta,
  actualizarStock,
  guardarExcel,
} from "../utils/excelHelper.js";
import { fechaLocalISO, obtenerFechaActual } from "../utils/date.util.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DATA_DIR = path.join(__dirname, "..", "..", "data");
const RUTA_LOG = path.join(DATA_DIR, "ventas.log");

function escribirEnLog(facturaId, fechaHora, total, cantidadProductos) {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }

    const linea = `${facturaId},${fechaHora},${total},${cantidadProductos}\n`;
    fs.appendFileSync(RUTA_LOG, linea, "utf8");
    console.log(`📝 Log escrito: ${facturaId}`);
  } catch (error) {
    console.error("❌ Error escribiendo log:", error.message);
  }
}

export const crearVenta = async (req, res) => {
  try {
    const { pago, productos } = req.body;

    // Validaciones básicas
    if (!productos?.length) {
      return res
        .status(400)
        .json({ error: "Debe incluir al menos un producto" });
    }

    // Calcular total
    const totalCalculado = productos.reduce(
      (sum, p) => sum + p.cantidad * p.precio,
      0,
    );

    // Validar pago
    const totalPagado = pago.efectivo + pago.transferencia;
    if (Math.abs(totalPagado - totalCalculado) > 0.01) {
      return res.status(400).json({
        error: "El monto pagado no coincide con el total",
        total: totalCalculado,
        pagado: totalPagado,
      });
    }

    // ===== PROCESAR VENTA (TODO O NADA) =====
    const resultado = await conExcelLock(async () => {
      const { workbook, hoja } = await abrirExcel();

      // 1. Verificar stock ANTES de procesar
      const erroresStock = [];
      for (const producto of productos) {
        const stockDisponible = await verificarStock(workbook, producto.codigo);
        if (stockDisponible < producto.cantidad) {
          erroresStock.push({
            codigo: producto.codigo,
            nombre: producto.nombre,
            disponible: stockDisponible,
            solicitado: producto.cantidad,
          });
        }
      }

      // Si hay errores de stock, cancelar TODO
      if (erroresStock.length > 0) {
        return { tipo: "error_stock", erroresStock };
      }

      // 2. Generar FacturaID y FechaHora desde la MISMA referencia de tiempo
      const ahora = obtenerFechaActual();
      const facturaId = generarFacturaIdExcel(hoja, ahora);
      const fechaHora = fechaLocalISO(ahora);

      // 3. Escribir líneas en Pendientes
      for (const producto of productos) {
        const subtotal = producto.cantidad * producto.precio;
        await escribirLineaVenta(hoja, {
          facturaId,
          fechaHora,
          codigoProducto: producto.codigo,
          nombreProducto: producto.nombre,
          cantidad: producto.cantidad,
          precioUnitario: producto.precio,
          subtotal,
          efectivo: pago.efectivo,
          transferencia: pago.transferencia,
        });
      }

      // 4. ACTUALIZAR STOCK en hoja Productos
      for (const producto of productos) {
        await actualizarStock(workbook, producto.codigo, producto.cantidad);
      }

      // 5. Guardar TODO (un solo save)
      await guardarExcel(workbook);

      return {
        tipo: "ok",
        facturaId,
        fechaHora,
        totalCalculado,
        productosResumen: productos.map((p) => ({
          codigo: p.codigo,
          nombre: p.nombre,
          cantidad: p.cantidad,
          subtotal: p.cantidad * p.precio,
        })),
      };
    });

    // Procesar resultado FUERA del mutex
    if (resultado.tipo === "error_stock") {
      return res.status(400).json({
        error: "Stock insuficiente para algunos productos",
        productos: resultado.erroresStock,
      });
    }

    // 6. ESCRIBIR EN LOG para Excel
    escribirEnLog(resultado.facturaId, resultado.fechaHora, resultado.totalCalculado, productos.length);

    // Respuesta exitosa
    res.status(201).json({
      success: true,
      facturaId: resultado.facturaId,
      total: resultado.totalCalculado,
      productos: resultado.productosResumen,
    });
  } catch (error) {
    console.error("❌ Error en crearVenta:", error.message);
    res.status(500).json({
      error: "Error al registrar la venta",
      detalle: error.message,
    });
  }
};

// Función auxiliar para verificar stock
async function verificarStock(workbook, codigoProducto) {
  try {
    const hojaProductos = workbook.sheet("Productos");
    let fila = 2;

    while (true) {
      const codigoCelda = hojaProductos.cell(`A${fila}`).value();
      if (!codigoCelda) break;

      if (codigoCelda.toString() === codigoProducto.toString()) {
        return hojaProductos.cell(`C${fila}`).value() || 0;
      }
      fila++;
    }
    return 0;
  } catch (error) {
    console.error("Error verificando stock:", error);
    return 0;
  }
}

// ============================================
// Función auxiliar: Obtener última factura ID de la hoja
// Lee la columna A y devuelve el valor máximo (última factura)
// ============================================
async function obtenerUltimaFacturaId(hoja) {
  try {
    const rango = hoja.usedRange();
    const datos = rango ? rango.value() : [];
    
    let maxFacturaId = '';
    
    // Empezar desde 1 (skip header)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || !fila[0]) continue;
      
      const facturaId = fila[0].toString();
      // Comparación string funciona porque son todos dígitos
      if (facturaId > maxFacturaId) {
        maxFacturaId = facturaId;
      }
    }
    
    return maxFacturaId;
  } catch (error) {
    console.error("Error obteniendo última factura ID:", error.message);
    return '';
  }
}

// ============================================
// SYN: Sincronizar ventas offline desde APK
// Recibe un array de ventas generadas offline
// ============================================
export const sincronizarVentas = async (req, res) => {
  try {
    const { ventas } = req.body;

    // Validaciones
    if (!ventas?.length) {
      return res.status(400).json({ error: "Debe incluir al menos una venta" });
    }

    console.log(`🔄 Sync: Recibiendo ${ventas.length} ventas para sincronizar`);

    // TODO el ciclo open → modificar → save corre dentro del mutex
    const resultado = await conExcelLock(async () => {
      // Abrir Excel una sola vez
      const { workbook, hoja } = await abrirExcel();

      let sincronizadas = 0;
      let errores = [];
      const resumenPorFactura = {};

      // Procesar cada venta (acepta formato plano o agrupado)
      for (const venta of ventas) {
        try {
          if (Array.isArray(venta.productos) && venta.productos.length > 0) {
            // Formato agrupado por factura
            for (const item of venta.productos) {
              await escribirLineaVenta(hoja, {
                facturaId: venta.facturaId,
                fechaHora: venta.fechaHora,
                codigoProducto: item.codigoProducto,
                nombreProducto: item.nombre,
                cantidad: item.cantidad,
                precioUnitario: item.precio,
                subtotal: item.subtotal,
                efectivo: venta.efectivo,
                transferencia: venta.transferencia,
              });

              await actualizarStock(workbook, item.codigoProducto, item.cantidad);

              const claveFactura = String(venta.facturaId || "sin_factura");
              if (!resumenPorFactura[claveFactura]) {
                resumenPorFactura[claveFactura] = {
                  facturaId: venta.facturaId,
                  fechaHora: venta.fechaHora,
                  total: 0,
                  cantidadProductos: 0,
                };
              }
              resumenPorFactura[claveFactura].total += Number(item.subtotal || 0);
              resumenPorFactura[claveFactura].cantidadProductos += 1;

              sincronizadas++;
            }
          } else {
            // Formato plano por línea
            const {
              facturaId,
              fechaHora,
              codigoProducto,
              nombre,
              cantidad,
              precio,
              subtotal,
              efectivo,
              transferencia,
            } = venta;

            await escribirLineaVenta(hoja, {
              facturaId,
              fechaHora,
              codigoProducto,
              nombreProducto: nombre,
              cantidad,
              precioUnitario: precio,
              subtotal,
              efectivo,
              transferencia,
            });

            await actualizarStock(workbook, codigoProducto, cantidad);

            const claveFactura = String(facturaId || "sin_factura");
            if (!resumenPorFactura[claveFactura]) {
              resumenPorFactura[claveFactura] = {
                facturaId,
                fechaHora,
                total: 0,
                cantidadProductos: 0,
              };
            }
            resumenPorFactura[claveFactura].total += Number(subtotal || 0);
            resumenPorFactura[claveFactura].cantidadProductos += 1;

            sincronizadas++;
          }
        } catch (err) {
          errores.push({
            codigo: venta.codigoProducto || venta.facturaId || "desconocido",
            error: err.message,
          });
        }
      }

      // Guardar Excel (AÚN DENTRO del mutex)
      await guardarExcel(workbook);

      // Obtener la última factura ID de datos.xlsx para referencia
      const ultimaFacturaId = await obtenerUltimaFacturaId(hoja);

      return { sincronizadas, errores, resumenPorFactura, ultimaFacturaId };
    });

    // Escribir ventas sincronizadas también en ventas.log (FUERA del mutex)
    Object.values(resultado.resumenPorFactura).forEach((r) => {
      if (!r.facturaId) return;
      escribirEnLog(r.facturaId, r.fechaHora, r.total, r.cantidadProductos);
    });

    console.log(`✅ Sync completado: ${resultado.sincronizadas} ventas sincronizadas`);
    console.log(`📋 Última factura en datos.xlsx: ${resultado.ultimaFacturaId}`);

    res.json({
      success: true,
      sincronizadas: resultado.sincronizadas,
      ultimaFacturaId: resultado.ultimaFacturaId || undefined,
      errores: resultado.errores.length > 0 ? resultado.errores : undefined,
    });
  } catch (error) {
    console.error("❌ Error en sincronizarVentas:", error.message);
    res.status(500).json({
      error: "Error al sincronizar ventas",
      detalle: error.message,
    });
  }
};
