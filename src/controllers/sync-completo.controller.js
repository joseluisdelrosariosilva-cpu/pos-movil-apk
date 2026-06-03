// ============================================
// CONTROLLER: Sync Completo (Batch)
// ============================================
// Recibe TODO en un solo request y procesa en orden:
//   1. Productos nuevos (con remapeo si hay conflicto)
//   2. Abastecimientos (usa mapa old→new de códigos)
//   3. Ventas (usa mapa old→new de códigos + facturas)
//   4. Mermas (usa mapa old→new de códigos)
//   5. Gastos
//
// Todo dentro de un solo conExcelLock = transacción atómica.
// Si algo falla, guardarExcel() nunca se llama → el Excel no se modifica.
// ============================================

import { conExcelLock, abrirExcel, guardarExcel, obtenerHojaPorNombre, actualizarStock, escribirLineaVenta } from "../utils/excelHelper.js";
import { fechaLocalISO } from "../utils/date.util.js";

// ============================================
// Helpers internos
// ============================================

/** Encuentra la primera fila vacía en una hoja (después de encabezados) */
const primeraFilaVacia = (hoja, col = "A") => {
  let fila = 2;
  while (hoja.cell(`${col}${fila}`).value()) {
    fila++;
  }
  return fila;
};

/**
 * Escanea la hoja Productos y devuelve:
 *   - Set con todos los códigos existentes (para detección rápida)
 *   - maxSufijo: el número más alto después de "Pr_"
 */
const escanearProductosExistentes = (hojaProductos) => {
  const codigos = new Set();
  let maxSufijo = 0;
  let fila = 2;

  while (true) {
    const val = hojaProductos.cell(`A${fila}`).value();
    if (!val) break;

    const str = val.toString().trim();
    codigos.add(str);

    if (str.startsWith("Pr_")) {
      const num = parseInt(str.slice(3), 10);
      if (!isNaN(num) && num > maxSufijo) maxSufijo = num;
    }

    fila++;
  }

  return { codigos, maxSufijo };
};

/**
 * Escanea la hoja Pendientes y devuelve:
 *   - Set con todas las facturas existentes (para detección rápida)
 *   - maxNum: el valor numérico más alto (global, sin importar prefijo)
 */
const escanearFacturasExistentes = (hojaPendientes) => {
  const facturas = new Set();
  let maxNum = 0;
  let fila = 2;

  while (true) {
    const val = hojaPendientes.cell(`A${fila}`).value();
    if (!val) break;

    const str = val.toString().trim();
    facturas.add(str);

    const num = parseInt(str, 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;

    fila++;
  }

  return { facturas, maxNum };
};

/**
 * Suma stock a un producto existente (para abastecimientos).
 * Lanza error si no encuentra el producto.
 */
const sumarStockProducto = (hojaProductos, codigo, cantidad) => {
  let fila = 2;

  while (true) {
    const celda = hojaProductos.cell(`A${fila}`).value();
    if (!celda) break;

    if (celda.toString() === codigo.toString()) {
      const actual = Number(hojaProductos.cell(`C${fila}`).value() || 0);
      const nuevo = actual + Number(cantidad);
      hojaProductos.cell(`C${fila}`).value(nuevo);
      return { fila, actual, nuevo };
    }

    fila++;
  }

  throw new Error(`sumarStockProducto: no se encontró código "${codigo}" en hoja Productos`);
};

/**
 * Extrae la fecha en formato DD/MM/YYYY desde un string datetime local ISO
 * Ejemplo: "2026-05-31T23:00:00.000" → "31/05/2026"
 * Fallback a la fecha actual si no puede procesar
 */
const formatearFechaDesdeISO = (fechaHora) => {
  if (!fechaHora) return new Date().toLocaleDateString("es-ES");
  const parteFecha = fechaHora.split("T")[0]; // "2026-05-31"
  if (!parteFecha) return new Date().toLocaleDateString("es-ES");
  const partes = parteFecha.split("-");
  if (partes.length !== 3) return new Date().toLocaleDateString("es-ES");
  return `${partes[2]}/${partes[1]}/${partes[0]}`; // "31/05/2026"
};

// ============================================
// SYNC COMPLETO
// ============================================
export const sincronizarCompleto = async (req, res) => {
  try {
    const { productos, abastecimientos, ventas, mermas, gastos, elaboraciones } = req.body;

    console.log("🔄 [sync-completo] Recibido batch:");
    console.log(`   ├─ Productos:       ${productos?.length || 0}`);
    console.log(`   ├─ Abastecimientos: ${abastecimientos?.length || 0}`);
    console.log(`   ├─ Ventas:          ${ventas?.length || 0}`);
    console.log(`   ├─ Mermas:          ${mermas?.length || 0}`);
    console.log(`   ├─ Gastos:          ${gastos?.length || 0}`);
    console.log(`   └─ Elaboraciones:   ${elaboraciones?.length || 0}`);

    // Si no viene nada, igual escaneamos Pendientes para devolver ultimaFactura
    if (
      (!productos || !productos.length) &&
      (!abastecimientos || !abastecimientos.length) &&
      (!ventas || !ventas.length) &&
      (!mermas || !mermas.length) &&
      (!gastos || !gastos.length) &&
      (!elaboraciones || !elaboraciones.length)
    ) {
      let ultimaFactura = null;
      try {
        const { workbook } = await abrirExcel();
        const hojaPendientes = workbook.sheet("Pendientes");
        if (hojaPendientes) {
          const { maxNum } = (() => {
            let maxNum = 0;
            let fila = 2;
            while (true) {
              const val = hojaPendientes.cell(`A${fila}`).value();
              if (!val) break;
              const num = parseInt(val.toString().trim(), 10);
              if (!isNaN(num) && num > maxNum) maxNum = num;
              fila++;
            }
            return { maxNum };
          })();
          ultimaFactura = maxNum > 0 ? String(maxNum).padStart(10, "0") : null;
        }
      } catch (_) {
        // Si falla al abrir Excel, devolvemos sin ultimaFactura
      }
      return res.json({
        success: true,
        sincronizados: {},
        remap: { codigos: {}, facturas: {} },
        ultimaFactura,
      });
    }

    const resultado = await conExcelLock(async () => {
      const { workbook } = await abrirExcel();

      // ---- MAPAS DE REMAPEO ----
      const remapCodigos = {};   // old → new
      const remapFacturas = {};  // old → new

      // ---- OBTENER HOJAS ----
      const hojaProductos = workbook.sheet("Productos");
      if (!hojaProductos) throw new Error('No se encontró hoja "Productos"');

      const hojaPendientes = workbook.sheet("Pendientes");
      if (!hojaPendientes) throw new Error('No se encontró hoja "Pendientes"');

      // Contadores
      let countProductos = 0;
      let countAbastecimientos = 0;
      let countVentas = 0;
      let countMermas = 0;
      let countGastos = 0;
      let countElaboraciones = 0;

      // ════════════════════════════════════════
      // PASO 1: PRODUCTOS NUEVOS
      // ════════════════════════════════════════
      if (productos && productos.length) {
        const hojaEntrada = workbook.sheet("Entrada");
        if (!hojaEntrada) throw new Error('No se encontró hoja "Entrada"');

        const { codigos, maxSufijo } = escanearProductosExistentes(hojaProductos);
        let proxSufijo = maxSufijo + 1;

        for (const prod of productos) {
          const codOriginal = prod.codigo?.toString().trim();
          if (!codOriginal || !prod.nombre) {
            console.warn("⚠️ [sync-completo] Producto sin código o nombre, saltando:", codOriginal);
            continue;
          }

          // ¿Conflicto?
          let codFinal = codOriginal;
          if (codigos.has(codOriginal)) {
            // Asignar nuevo código: Pr_ + proxSufijo con 5 dígitos
            codFinal = "Pr_" + String(proxSufijo).padStart(5, "0");
            proxSufijo++;

            remapCodigos[codOriginal] = codFinal;
            console.log(`   ↪ Producto remapeado: ${codOriginal} → ${codFinal}`);
          }

          // Escribir en Productos
          const filaP = primeraFilaVacia(hojaProductos, "A");
          hojaProductos.cell(`A${filaP}`).value(codFinal);
          hojaProductos.cell(`B${filaP}`).value(prod.nombre);
          hojaProductos.cell(`C${filaP}`).value(Number(prod.cantidad) || 0);
          hojaProductos.cell(`D${filaP}`).value(Number(prod.precio_venta) || 0);

          // Escribir en Entrada (con la fecha real de creación del producto)
          const filaE = primeraFilaVacia(hojaEntrada, "A");
          hojaEntrada.cell(`A${filaE}`).value(codFinal);
          hojaEntrada.cell(`B${filaE}`).value(prod.nombre);
          hojaEntrada.cell(`C${filaE}`).value(formatearFechaDesdeISO(prod.fecha_hora));
          hojaEntrada.cell(`D${filaE}`).value(Number(prod.cantidad) || 0);
          hojaEntrada.cell(`E${filaE}`).value(Number(prod.precio_venta) || 0);
          hojaEntrada.cell(`F${filaE}`).value(Number(prod.precio_costo) || 0);

          // Agregar código al Set para que siguientes productos del batch
          // detecten conflicto contra otro producto del mismo batch
          codigos.add(codFinal);

          countProductos++;
        }
        console.log(`   ✅ ${countProductos} productos procesados`);
      }

      // ════════════════════════════════════════
      // PASO 2: ABASTECIMIENTOS
      // ════════════════════════════════════════
      if (abastecimientos && abastecimientos.length) {
        const hojaAbastecimiento = workbook.sheet("Abastecimiento");
        if (!hojaAbastecimiento) throw new Error('No se encontró hoja "Abastecimiento"');

        for (const abast of abastecimientos) {
          const codOriginal = abast.codigo?.toString().trim();
          if (!codOriginal) continue;

          // Traducir código si fue remapeado
          const codFinal = remapCodigos[codOriginal] || codOriginal;
          const cantidad = Number(abast.cantidad || 0);
          if (!cantidad) continue;

          // Escribir en Abastecimiento
          const filaA = primeraFilaVacia(hojaAbastecimiento, "A");
          hojaAbastecimiento.cell(`A${filaA}`).value(codFinal);
          hojaAbastecimiento.cell(`B${filaA}`).value(cantidad);
          hojaAbastecimiento.cell(`C${filaA}`).value(abast.fechaHora || fechaLocalISO());

          // Sumar al stock
          sumarStockProducto(hojaProductos, codFinal, cantidad);

          countAbastecimientos++;
        }
        console.log(`   ✅ ${countAbastecimientos} abastecimientos procesados`);
      }

      // ════════════════════════════════════════
      // PASO 3: VENTAS
      // ════════════════════════════════════════
      if (ventas && ventas.length) {
        const { facturas: facturasExistentes, maxNum: maxFacturaNum } = escanearFacturasExistentes(hojaPendientes);
        let proxFactura = maxFacturaNum + 1;

        for (const venta of ventas) {
          const facturaOriginal = venta.facturaId?.toString().trim();
          if (!facturaOriginal) continue;

          // Traducir producto si fue remapeado
          const codProducto = remapCodigos[venta.codigoProducto?.toString().trim()] || venta.codigoProducto;

          // ¿Conflicto de factura ID?
          let facturaFinal = facturaOriginal;
          if (remapFacturas[facturaOriginal]) {
            // Ya remapeamos esta factura antes en el batch
            facturaFinal = remapFacturas[facturaOriginal];
          } else if (facturasExistentes.has(facturaOriginal)) {
            // Primera vez que vemos esta factura y ya existe en Excel
            facturaFinal = String(proxFactura).padStart(10, "0");
            proxFactura++;
            remapFacturas[facturaOriginal] = facturaFinal;
            // Agregar al Set para detectar conflictos intra-batch
            facturasExistentes.add(facturaFinal);
            console.log(`   ↪ Factura remapeada: ${facturaOriginal} → ${facturaFinal}`);
          }

          // Escribir línea en Pendientes
          escribirLineaVenta(hojaPendientes, {
            facturaId: facturaFinal,
            fechaHora: venta.fechaHora,
            codigoProducto: codProducto,
            nombreProducto: venta.nombre,
            cantidad: venta.cantidad,
            precioUnitario: venta.precio || venta.precioUnitario,
            subtotal: venta.subtotal,
            efectivo: venta.efectivo,
            transferencia: venta.transferencia,
          });

          // Descontar stock
          actualizarStock(workbook, codProducto, venta.cantidad);

          countVentas++;
        }
        console.log(`   ✅ ${countVentas} ventas procesadas`);
      }

      // ════════════════════════════════════════
      // PASO 4: MERMAS
      // ════════════════════════════════════════
      if (mermas && mermas.length) {
        const hojaMerma = workbook.sheet("Merma");
        if (!hojaMerma) throw new Error('No se encontró hoja "Merma"');

        for (const merma of mermas) {
          const codOriginal = merma.codigo?.toString().trim();
          if (!codOriginal) continue;

          const codFinal = remapCodigos[codOriginal] || codOriginal;
          const cantidad = Number(merma.cantidad || 0);
          if (!cantidad) continue;

          // Escribir en Merma
          const filaM = primeraFilaVacia(hojaMerma, "A");
          hojaMerma.cell(`A${filaM}`).value(codFinal);
          hojaMerma.cell(`B${filaM}`).value(merma.nombre || "");
          hojaMerma.cell(`C${filaM}`).value(cantidad);

          // Descontar stock
          actualizarStock(workbook, codFinal, cantidad);

          countMermas++;
        }
        console.log(`   ✅ ${countMermas} mermas procesadas`);
      }

      // ════════════════════════════════════════
      // PASO 5: GASTOS
      // ════════════════════════════════════════
      if (gastos && gastos.length) {
        let hojaGastos;
        try {
          hojaGastos = obtenerHojaPorNombre(workbook, "Gastos");
        } catch (_) {
          hojaGastos = workbook.addSheet("Gastos");
          hojaGastos.cell("A1").value("Fecha");
          hojaGastos.cell("B1").value("Descripcion");
          hojaGastos.cell("C1").value("Monto");
        }

        for (const gasto of gastos) {
          const filaG = primeraFilaVacia(hojaGastos, "A");
          hojaGastos.cell(`A${filaG}`).value(gasto.fecha);
          hojaGastos.cell(`B${filaG}`).value(gasto.descripcion);
          hojaGastos.cell(`C${filaG}`).value(Number(gasto.monto || 0));
          countGastos++;
        }
        console.log(`   ✅ ${countGastos} gastos procesados`);
      }

      // ════════════════════════════════════════
      // PASO 6: ELABORACIONES
      // ════════════════════════════════════════
      if (elaboraciones && elaboraciones.length) {
        let hojaElaboracion;
        try {
          hojaElaboracion = obtenerHojaPorNombre(workbook, "Elaboracion");
        } catch (_) {
          hojaElaboracion = workbook.addSheet("Elaboracion");
          hojaElaboracion.cell("A1").value("NombreReceta");
          hojaElaboracion.cell("B1").value("Lotes");
        }

        for (const elab of elaboraciones) {
          const filaEl = primeraFilaVacia(hojaElaboracion, "A");
          hojaElaboracion.cell(`A${filaEl}`).value(elab.nombre_receta);
          hojaElaboracion.cell(`B${filaEl}`).value(Number(elab.lotes || 0));
          countElaboraciones++;
        }
        console.log(`   ✅ ${countElaboraciones} elaboraciones procesadas`);
      }

      // ════════════════════════════════════════
      // GUARDAR TODO (un solo write)
      // ════════════════════════════════════════
      await guardarExcel(workbook);

      // Obtener la última factura de Pendientes para que el cliente
      // actualice su referencia y no genere IDs duplicados offline
      const { maxNum: maxFacturaNum } = escanearFacturasExistentes(hojaPendientes);
      const ultimaFactura = maxFacturaNum > 0
        ? String(maxFacturaNum).padStart(10, "0")
        : null;

      return {
        sincronizados: {
          productos: countProductos,
          abastecimientos: countAbastecimientos,
          ventas: countVentas,
          mermas: countMermas,
          gastos: countGastos,
          elaboraciones: countElaboraciones,
        },
        remap: {
          codigos: remapCodigos,
          facturas: remapFacturas,
        },
        ultimaFactura,
      };
    });

    console.log("✅ [sync-completo] Batch completado exitosamente");
    console.log("   Remapeo de códigos:", Object.keys(resultado.remap.codigos).length ? resultado.remap.codigos : "ninguno");
    console.log("   Remapeo de facturas:", Object.keys(resultado.remap.facturas).length ? resultado.remap.facturas : "ninguno");

    return res.json({
      success: true,
      sincronizados: resultado.sincronizados,
      remap: resultado.remap,
      ultimaFactura: resultado.ultimaFactura,
    });
  } catch (error) {
    console.error("❌ [sync-completo] Error:", error.message);
    console.error("   Stack:", error.stack);
    return res.status(500).json({
      success: false,
      error: "ERROR_SYNC_COMPLETO",
      mensaje: error.message,
    });
  }
};
