// ============================================
// CONTROLLER: Entrada de Productos
// ============================================
// Sincroniza productos nuevos desde la APK al Excel
// Escribe en hojas "Productos" y "Entrada"
// ============================================

import XlsxPopulate from "xlsx-populate";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const RUTA_EXCEL = path.resolve(__dirname, "..", "..", "data", "datos.xlsx");

// ============================================
// Función para obtener la primera fila vacía en una hoja
// ============================================
const obtenerPrimeraFilaVacia = (hoja, columnaInicio = "A") => {
  let fila = 2; // Empezamos después de encabezados
  while (hoja.cell(`${columnaInicio}${fila}`).value()) {
    fila++;
  }
  return fila;
};

// ============================================
// POST /api/entrada-productos
// ============================================
export const sincronizarEntradaProductos = async (req, res) => {
  try {
    console.log("🔍 [DEBUG SERVER] sincronizarEntradaProductos iniciado");
    const { productos } = req.body;
    console.log("🔍 [DEBUG SERVER] Body recibido:", JSON.stringify(req.body).substring(0, 200));
    console.log("🔍 [DEBUG SERVER] Cantidad de productos:", productos ? productos.length : 0);

    // Validar que viene el array de productos
    if (!productos || !Array.isArray(productos) || productos.length === 0) {
      console.log("🔍 [DEBUG SERVER] No hay productos para sincronizar");
      return res.status(400).json({
        success: false,
        error: "NO_HAY_PRODUCTOS",
        mensaje: "No se recibieron productos para sincronizar",
      });
    }

    console.log(`📦 [entrada-productos] Recibidos ${productos.length} productos`);

    // Abrir Excel
    console.log("🔍 [DEBUG SERVER] Abriendo Excel:", RUTA_EXCEL);
    const workbook = await XlsxPopulate.fromFileAsync(RUTA_EXCEL);
    console.log("🔍 [DEBUG SERVER] Excel abierto correctamente");

    const hojaProductos = workbook.sheet("Productos");
    if (!hojaProductos) {
      throw new Error('No se encontró la hoja "Productos"');
    }
    console.log("🔍 [DEBUG SERVER] Hoja Productos obtenida");

    const hojaEntrada = workbook.sheet("Entrada");
    if (!hojaEntrada) {
      throw new Error('No se encontró la hoja "Entrada"');
    }
    console.log("🔍 [DEBUG SERVER] Hoja Entrada obtenida");

    // Fecha actual para la hoja Entrada
    const fechaActual = new Date().toLocaleDateString("es-ES");
    console.log("🔍 [DEBUG SERVER] Fecha actual:", fechaActual);

    let productosSincronizados = 0;
    const detalles = [];

    // Procesar cada producto
    for (const producto of productos) {
      const {
        codigo,
        nombre,
        cantidad,
        precio_venta,
        precio_costo,
      } = producto;

      console.log("🔍 [DEBUG SERVER] Procesando producto:", codigo, nombre);

      if (!codigo || !nombre) {
        console.warn("⚠️ Producto sin código o nombre, saltando...");
        continue;
      }

      // 1. Escribir en hoja "Productos" (Código, Producto, Disponibilidad, Precio)
      const filaProductos = obtenerPrimeraFilaVacia(hojaProductos, "A");
      console.log("🔍 [DEBUG SERVER] Escribiendo en Productos fila:", filaProductos);
      hojaProductos.cell(`A${filaProductos}`).value(codigo); // Código
      hojaProductos.cell(`B${filaProductos}`).value(nombre); // Producto
      hojaProductos.cell(`C${filaProductos}`).value(Number(cantidad) || 0); // Disponibilidad
      hojaProductos.cell(`D${filaProductos}`).value(Number(precio_venta) || 0); // Precio

      console.log(`  ✅ [Productos] Fila ${filaProductos}: ${codigo} - ${nombre}`);

      // 2. Escribir en hoja "Entrada" (Codigo, Nombre, Fecha, Cant Inicial, Precio Venta, Precio Costo)
      const filaEntrada = obtenerPrimeraFilaVacia(hojaEntrada, "A");
      console.log("🔍 [DEBUG SERVER] Escribiendo en Entrada fila:", filaEntrada);
      hojaEntrada.cell(`A${filaEntrada}`).value(codigo); // Codigo
      hojaEntrada.cell(`B${filaEntrada}`).value(nombre); // Nombre
      hojaEntrada.cell(`C${filaEntrada}`).value(fechaActual); // Fecha
      hojaEntrada.cell(`D${filaEntrada}`).value(Number(cantidad) || 0); // Cant Inicial
      hojaEntrada.cell(`E${filaEntrada}`).value(Number(precio_venta) || 0); // Precio Venta
      hojaEntrada.cell(`F${filaEntrada}`).value(Number(precio_costo) || 0); // Precio Costo

      console.log(`  ✅ [Entrada] Fila ${filaEntrada}: ${codigo} - ${nombre}`);

      detalles.push({
        codigo,
        filaProductos,
        filaEntrada,
      });

      productosSincronizados++;
    }

    // Guardar Excel
    console.log("🔍 [DEBUG SERVER] Guardando Excel...");
    await workbook.toFileAsync(RUTA_EXCEL);
    console.log("🔍 [DEBUG SERVER] Excel guardado correctamente");

    console.log(`✅ [entrada-productos] ${productosSincronizados} productos sincronizados al Excel`);

    return res.json({
      success: true,
      sincronizadas: productosSincronizados,
      sincronizados: productosSincronizados,
      mensaje: `${productosSincronizados} productos sincronizados correctamente`,
      detalles,
    });
  } catch (error) {
    console.error("❌ [entrada-productos] Error:", error.message);
    console.error("❌ [entrada-productos] Stack:", error.stack);
    return res.status(500).json({
      success: false,
      error: "ERROR_SINCRONIZACION",
      mensaje: error.message,
    });
  }
};
