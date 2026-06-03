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
        escribirLineaVenta(hoja, {
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
        actualizarStock(workbook, producto.codigo, producto.cantidad);
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


