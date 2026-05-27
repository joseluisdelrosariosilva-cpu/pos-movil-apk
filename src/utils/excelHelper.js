// ============================================
// UTILS: Excel Helper
// ============================================
// Funciones reutilizables para manejar el Excel intermediario
// ============================================

import XlsxPopulate from "xlsx-populate";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

// Obtener __dirname equivalente en ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Ruta fija al archivo Excel (desde la raíz del proyecto)
const RUTA_EXCEL = path.resolve(__dirname, "..", "..", "data", "datos.xlsx");
const NOMBRE_HOJA = "Pendientes";

// ============================================
// MUTEX PARA ACCESO EXCLUSIVO AL EXCEL
// ============================================
// Node.js es single-thread, pero async hace que requests concurrentes
// puedan pisar el open/save del otro. Este lock serializa TODO el
// ciclo abrir → modificar → guardar sobre datos.xlsx.
//
// Uso: const resultado = await conExcelLock(() => { ... });
//
// SEGURIDAD:
//   - La cadena de promesas NUNCA puede romperse: si la promesa anterior
//     rechaza, el .catch() mantiene viva la cola.
//   - Si await promesaAnterior falla, se libera la cola igual para no
//     bloquear a los siguientes.
//   - Timeout de advertencia (configurable via EXCEL_LOCK_TIMEOUT_MS,
//     default 30s). LOGEA si una tarea se cuelga, NO libera el lock
//     (liberarlo con la tarea aún ejecutándose = corrupción del Excel).
//   - Estadísticas exportables para monitoreo.
// ============================================
let colaExcel = Promise.resolve();
let tareaActiva = false;
let tareaIdActual = null;
let inicioTareaActual = 0;
let profundidadCola = 0;
let timeoutActual = null;

/**
 * Ejecuta fn() con acceso exclusivo a datos.xlsx.
 * Las llamadas concurrentes se encolan y ejecutan en serie.
 *
 * @param {Function} fn - Función asíncrona a ejecutar con el lock adquirido.
 * @param {number} [timeoutMs] - Timeout de advertencia (default: EXCEL_LOCK_TIMEOUT_MS o 30000).
 * @returns {Promise<any>} Resultado de fn()
 */
export const conExcelLock = async (fn, timeoutMs) => {
  const TIMEOUT =
    timeoutMs ||
    parseInt(process.env.EXCEL_LOCK_TIMEOUT_MS || "30000", 10);

  const idUnico = Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 6);
  let liberar;
  const esperar = new Promise((resolve) => {
    liberar = resolve;
  });

  // Encolar: la promesa anterior debe resolverse (o rechazarse)
  // antes de que esta tarea pueda ejecutarse.
  // .catch(() => esperar) → si la cadena previa rechaza, la cola
  // no se traba: espera igual a que se llame liberar().
  const promesaAnterior = colaExcel;
  colaExcel = colaExcel
    .then(() => esperar)
    .catch(() => esperar);

  profundidadCola++;

  try {
    // Esperar turno
    await promesaAnterior;

    // --- Lock adquirido ---
    tareaActiva = true;
    tareaIdActual = idUnico;
    inicioTareaActual = Date.now();
    profundidadCola--;

    // Timeout de advertencia (solo log, no libera — liberar con la tarea
    // activa causaría corrupción del Excel)
    timeoutActual = setTimeout(() => {
      const duracion = Date.now() - inicioTareaActual;
      console.error(
        `⚠️  [conExcelLock] TIMEOUT — Tarea "${idUnico}" excede ${TIMEOUT}ms ` +
          `(lleva ${duracion}ms). ` +
          `Profundidad de cola: ${profundidadCola}. ` +
          `POSIBLE DEADLOCK. Revisar el servidor.`,
      );
    }, TIMEOUT);

    try {
      return await fn();
    } finally {
      clearTimeout(timeoutActual);
      timeoutActual = null;
      tareaActiva = false;
      tareaIdActual = null;
      inicioTareaActual = 0;
      liberar(); // ← libera al siguiente en la cola
    }
  } catch (err) {
    // Si promesaAnterior rechazó o fn() lanzó, igual liberamos la cola
    // para no dejar trabado al resto.
    clearTimeout(timeoutActual);
    timeoutActual = null;
    tareaActiva = false;
    tareaIdActual = null;
    inicioTareaActual = 0;
    profundidadCola--;
    liberar();
    throw err;
  }
};

// ============================================
// Diagnóstico del mutex
// ============================================
export const diagnosticarMutex = () => ({
  tareaActiva,
  tareaIdActual,
  inicioTareaActual: inicioTareaActual
    ? new Date(inicioTareaActual).toISOString()
    : null,
  duracionMs: inicioTareaActual ? Date.now() - inicioTareaActual : 0,
  profundidadCola: Math.max(0, profundidadCola),
});

/**
 * Fuerza el reinicio de la cola si está trabada.
 * USAR SOLO COMO ÚLTIMO RECURSO — si hay una tarea activa, el Excel
 * puede quedar en estado inconsistente.
 */
export const reiniciarCola = () => {
  console.warn(
    "⚠️  [conExcelLock] REINICIO FORZADO de la cola. " +
      (tareaActiva
        ? `Había una tarea activa: "${tareaIdActual}". Posible corrupción.`
        : "No había tarea activa."),
  );
  colaExcel = Promise.resolve();
  tareaActiva = false;
  tareaIdActual = null;
  inicioTareaActual = 0;
  profundidadCola = 0;
  if (timeoutActual) {
    clearTimeout(timeoutActual);
    timeoutActual = null;
  }
};

// ============================================
// 1. ABRIR EXCEL - Obtiene el workbook y la hoja
// ============================================
export const abrirExcel = async () => {
  try {
    console.log("📂 Abriendo Excel:", RUTA_EXCEL);
    const workbook = await XlsxPopulate.fromFileAsync(RUTA_EXCEL);
    const hoja = workbook.sheet(NOMBRE_HOJA);

    if (!hoja) {
      throw new Error(`No se encontró la hoja "${NOMBRE_HOJA}" en el Excel`);
    }

    return { workbook, hoja, ruta: RUTA_EXCEL };
  } catch (error) {
    console.error("❌ Error abriendo Excel:", error.message);
    throw new Error(`No se pudo abrir el Excel: ${error.message}`);
  }
};

// ============================================
// 2. GUARDAR EXCEL - Persiste los cambios
// ============================================
export const guardarExcel = async (workbook) => {
  try {
    await workbook.toFileAsync(RUTA_EXCEL);
    console.log("✅ Excel guardado correctamente");
    return true;
  } catch (error) {
    console.error("❌ Error guardando Excel:", error.message);
    throw new Error(`No se pudo guardar el Excel: ${error.message}`);
  }
};

// ============================================
// 3. GENERAR FACTURA ID ESTILO EXCEL (fecha + consecutivo)
// ============================================
// Formato: [5 dígitos fecha serial] + [5 dígitos consecutivo]
// Ejemplo: 4537100023 (donde 45371 = fecha, 00023 = factura 23 del día)
// ============================================

/**
 * Convierte una fecha JS a fecha serial de Excel (matching VBA CLng(Date))
 * VBA cuenta días desde 1/1/1900 y considera 1900 como bisiesto (bug del 29/2/1900)
 * IMPORTANTE: Usar fecha LOCAL para coincidir con VBA
 */
const fechaATimestampExcel = (fecha) => {
  // Fecha base de Excel en UTC
  const fechaBase = Date.UTC(1900, 0, 1);

  // Fecha objetivo en UTC (sin hora)
  const fechaObj = Date.UTC(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());

  // Diferencia en días
  const diffDias = Math.floor((fechaObj - fechaBase) / (1000 * 60 * 60 * 24));

  // Ajuste por el bug de Excel (considera 1900 como bisiesto)
  // Si la fecha es posterior al 28/2/1900, sumamos 1
  const ajuste = diffDias > 59 ? 1 : 0;

  return diffDias + 1 + ajuste; // +1 porque la serie de Excel empieza en 1
};

/**
 * Genera un nuevo FacturaID basado en la fecha actual
 * El sufijo se reinicia cada día desde 00001
 */
export const generarFacturaIdExcel = (hoja, fecha = new Date()) => {
  try {
    // 1. Convertir fecha a número Excel (5 dígitos)
    const fechaSerial = fechaATimestampExcel(fecha);
    const prefijo = fechaSerial.toString().padStart(5, "0");

    console.log(
      `📅 Fecha: ${fecha.toLocaleDateString()} → Serial: ${fechaSerial} (prefijo: ${prefijo})`,
    );

    // 2. Buscar IDs con el mismo prefijo (mismo día)
    let maxSufijo = 0;
    let fila = 2;

    while (true) {
      const facturaIdCompleto = hoja.cell(`A${fila}`).value();
      if (!facturaIdCompleto) break;

      const facturaIdStr = facturaIdCompleto.toString();

      // Verificar si tiene el mismo prefijo (mismo día)
      if (facturaIdStr.startsWith(prefijo)) {
        // Extraer el sufijo (últimos 5 dígitos)
        const sufijo = parseInt(facturaIdStr.slice(-5), 10);
        if (sufijo > maxSufijo) {
          maxSufijo = sufijo;
        }
      }

      fila++;
    }

    // 3. Nuevo sufijo (si no hay facturas hoy, empieza desde 1)
    const nuevoSufijo = maxSufijo + 1;

    // 4. Formatear: prefijo + sufijo con 5 dígitos
    const nuevoFacturaId = prefijo + nuevoSufijo.toString().padStart(5, "0");

    console.log(
      `✅ Nuevo FacturaID generado: ${nuevoFacturaId} (sufijo: ${nuevoSufijo})`,
    );

    return nuevoFacturaId;
  } catch (error) {
    console.error("❌ Error generando FacturaID estilo Excel:", error.message);
    throw new Error("No se pudo generar el FacturaID");
  }
};

// ============================================
// 4. ESCRIBIR UNA LÍNEA DE VENTA
// ============================================
export const escribirLineaVenta = (hoja, datosVenta) => {
  try {
    // Encontrar la primera fila vacía
    let fila = 1;
    while (hoja.cell(`A${fila + 1}`).value()) {
      fila++;
    }
    fila++; // Ahora fila es la primera vacía

    const {
      facturaId,
      fechaHora,
      codigoProducto,
      nombreProducto,
      cantidad,
      precioUnitario,
      subtotal,
      efectivo,
      transferencia,
    } = datosVenta;

    // Escribir cada campo en su columna
    hoja.cell(`A${fila}`).value(facturaId); // FacturaID
    hoja.cell(`B${fila}`).value(fechaHora); // FechaHora
    hoja.cell(`C${fila}`).value(codigoProducto); // CodigoProducto
    hoja.cell(`D${fila}`).value(nombreProducto); // Nombre
    hoja.cell(`E${fila}`).value(cantidad); // CantVendida
    hoja.cell(`F${fila}`).value(precioUnitario); // PrecioUnitario
    hoja.cell(`G${fila}`).value(subtotal); // Subtotal
    hoja.cell(`H${fila}`).value(efectivo); // Efectivo
    hoja.cell(`I${fila}`).value(transferencia); // Transferencia
    hoja.cell(`J${fila}`).value(false); // Procesado (FALSE por defecto)

    console.log(`✅ Línea escrita en fila ${fila} para FacturaID ${facturaId}`);
    return fila;
  } catch (error) {
    console.error("❌ Error escribiendo línea de venta:", error.message);
    throw new Error("No se pudo escribir la línea de venta");
  }
};

// ============================================
// 5. VALIDAR QUE EL EXCEL EXISTE (útil al inicio)
// ============================================
export const validarExcelExiste = () => {
  const existe = fs.existsSync(RUTA_EXCEL);

  if (!existe) {
    console.error(`❌ No se encuentra el Excel en: ${RUTA_EXCEL}`);
    console.error("   Asegúrate de crear el archivo en esa ubicación");
    return false;
  }

  console.log(`✅ Excel encontrado en: ${RUTA_EXCEL}`);
  return true;
};

// ============================================
// 6. FUNCIÓN DE PRUEBA (actualizada)
// ============================================
export const probarHelper = async () => {
  console.log("\n🧪 Probando Excel Helper...");

  try {
    // 1. Verificar que el Excel existe
    if (!validarExcelExiste()) {
      console.log("❌ Prueba fallida: Excel no encontrado");
      return;
    }

    // 2. Abrir Excel
    const { workbook, hoja } = await abrirExcel();
    console.log("✅ Excel abierto correctamente");

    // 3. Probar generador de FacturaID estilo Excel
    console.log("\n📋 Probando generador de FacturaID...");
    const facturaId1 = generarFacturaIdExcel(hoja, new Date("2024-01-15")); // Fecha fija para prueba
    console.log(`   FacturaID generado: ${facturaId1}`);

    const facturaId2 = generarFacturaIdExcel(hoja, new Date("2024-01-15")); // Misma fecha
    console.log(`   Segundo ID: ${facturaId2} (debería ser superior)`);

    const facturaId3 = generarFacturaIdExcel(hoja, new Date("2024-01-16")); // Día siguiente
    console.log(`   ID día siguiente: ${facturaId3} (prefijo debe cambiar)`);

    // 4. Escribir una línea de prueba con el nuevo ID
    const datosPrueba = {
      facturaId: facturaId1,
      fechaHora: new Date().toISOString(),
      codigoProducto: 999,
      nombreProducto: "Producto de Prueba",
      cantidad: 1,
      precioUnitario: 100,
      subtotal: 100,
      efectivo: 100,
      transferencia: 0,
    };

    escribirLineaVenta(hoja, datosPrueba);
    console.log("✅ Línea de prueba escrita");

    // 5. Guardar Excel
    await guardarExcel(workbook);
    console.log("✅ Excel guardado");

    console.log("🎉 Prueba completada exitosamente!\n");
  } catch (error) {
    console.error("❌ Prueba fallida:", error.message);
  }
};

// ============================================
// 7. OBTENER HOJA POR NOMBRE (desde un workbook existente)
// ============================================
export const obtenerHojaPorNombre = (workbook, nombreHoja) => {
  try {
    const hoja = workbook.sheet(nombreHoja);

    if (!hoja) {
      throw new Error(`No se encontró la hoja "${nombreHoja}"`);
    }

    console.log(`📄 Hoja "${nombreHoja}" obtenida correctamente`);
    return hoja;
  } catch (error) {
    console.error(`❌ Error obteniendo hoja "${nombreHoja}":`, error.message);
    throw error;
  }
};

// ============================================
// 8. ACTUALIZAR STOCK DE UN PRODUCTO
// ============================================
export const actualizarStock = (workbook, codigoProducto, cantidadVendida) => {
  try {
    // Obtener la hoja de productos
    const hojaProductos = workbook.sheet("Productos");
    if (!hojaProductos) {
      throw new Error('No se encontró la hoja "Productos"');
    }

    // Buscar el producto por código
    let fila = 2; // Empezamos después de encabezados
    let encontrado = false;
    let stockActual = 0;

    while (true) {
      const codigoCelda = hojaProductos.cell(`A${fila}`).value();
      if (!codigoCelda) break; // Fin de los datos

      // Comparar códigos (como string para evitar problemas)
      if (codigoCelda.toString() === codigoProducto.toString()) {
        encontrado = true;

        const columnaStock = "C";
        stockActual = hojaProductos.cell(`${columnaStock}${fila}`).value() || 0;

        // Calcular nuevo stock
        const nuevoStock = Math.max(0, stockActual - cantidadVendida);

        // Actualizar celda
        hojaProductos.cell(`${columnaStock}${fila}`).value(nuevoStock);

        console.log(
          `📦 Producto ${codigoProducto}: stock ${stockActual} → ${nuevoStock}`,
        );
        break;
      }

      fila++;
    }

    if (!encontrado) {
      throw new Error(
        `Producto ${codigoProducto} no encontrado en hoja Productos`,
      );
    }

    return {
      encontrado,
      stockActual,
      nuevoStock: stockActual - cantidadVendida,
    };
  } catch (error) {
    console.error("❌ Error actualizando stock:", error.message);
    throw error;
  }
};
