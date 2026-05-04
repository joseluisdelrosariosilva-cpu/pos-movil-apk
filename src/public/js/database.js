// ============================================
// DATABASE MODULE - SQLite para modo offline
// ============================================
// Versión vanilla para ejecutarse en webview del APK

var db = null;
var storageMode = "none"; // sqlite | local
var cancelarEscaneo = false; // Bandera para cancelar escaneo de red

const LS_KEYS = {
  productos: "posmovil_productos",
  ventas: "posmovil_ventas_pending",
  config: "posmovil_config",
};

function leerLocal(key, fallback) {
  try {
    var raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (_) {
    return fallback;
  }
}

function guardarLocal(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.error("❌ Error guardando localStorage:", e);
    return false;
  }
}

function asegurarStoreLocal() {
  if (!localStorage.getItem(LS_KEYS.productos)) guardarLocal(LS_KEYS.productos, []);
  if (!localStorage.getItem(LS_KEYS.ventas)) guardarLocal(LS_KEYS.ventas, []);
  if (!localStorage.getItem(LS_KEYS.config)) guardarLocal(LS_KEYS.config, {});
}

function activarModoLocal(motivo) {
  storageMode = "local";
  db = null;
  asegurarStoreLocal();
  console.warn("📦 [DB] Modo localStorage activo:", motivo);
}

function normalizarProducto(p) {
  return {
    codigo: p.codigo,
    nombre: p.producto || p.nombre,
    precio: Number(p.precio || 0),
    disponibilidad: Number(p.disponibilidad || p.stock || 0),
  };
}

function construirLineasVenta(venta, syncedEstado, facturaIdForzada) {
  var synced = Number.isFinite(Number(syncedEstado)) ? Number(syncedEstado) : 0;
  var facturaId = facturaIdForzada || venta.facturaId;
  var lineas = [];
  for (var i = 0; i < venta.productos.length; i++) {
    var item = venta.productos[i];
    lineas.push({
      id: Date.now() + i + Math.floor(Math.random() * 1000),
      factura_id: facturaId,
      fecha_hora: venta.fechaHora,
      codigo_producto: item.codigo,
      nombre: item.nombre,
      cantidad: Number(item.cantidad || 0),
      precio: Number(item.precio || 0),
      subtotal: Number(item.cantidad || 0) * Number(item.precio || 0),
      efectivo: Number((venta.pago && venta.pago.efectivo) || 0),
      transferencia: Number((venta.pago && venta.pago.transferencia) || 0),
      synced: synced,
    });
  }
  return lineas;
}

function contarFacturasPendientesLocal(ventas) {
  var setFacturas = {};
  for (var i = 0; i < ventas.length; i++) {
    var v = ventas[i];
    if (Number(v.synced || 0) !== 0) continue;
    var key = String(v.factura_id || ("sin_factura_" + i));
    setFacturas[key] = true;
  }
  return Object.keys(setFacturas).length;
}

function fechaLocalHoyISO() {
  var ahora = new Date();
  var y = ahora.getFullYear();
  var m = String(ahora.getMonth() + 1).padStart(2, "0");
  var d = String(ahora.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + d;
}

// ============================================
// LIMPIAR VENTAS ANTIGUAS (más de 30 días y ya sincronizadas)
// ============================================
async function limpiarVentasAntiguas() {
  var ULTIMA_LIMPIEZA_KEY = "posmovil_ultima_limpieza";
  var hoy = fechaLocalHoyISO();
  
  // Ya limpiamos hoy? Skip
  if (localStorage.getItem(ULTIMA_LIMPIEZA_KEY) === hoy) return;
  
  // Calcular fecha límite (30 días atrás)
  var fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - 30);
  var fechaLimiteISO = fechaLimite.getFullYear() + '-' + 
                       String(fechaLimite.getMonth()+1).padStart(2,'0') + '-' + 
                       String(fechaLimite.getDate()).padStart(2,'0');
  
  try {
    if (storageMode === "sqlite" && db) {
      await db.execute(
        "DELETE FROM ventas_pending WHERE synced = 1 AND date(fecha_hora) < ?", 
        [fechaLimiteISO]
      );
      console.log("🗑️ Ventas antiguas eliminadas (antes del " + fechaLimiteISO + ")");
    } else {
      // localStorage: filtrar y guardar solo las recientes
      var ventas = leerLocal(LS_KEYS.ventas, []);
      var ventasFiltradas = ventas.filter(function(v) {
        var fechaVenta = extraerFechaISO(v.fecha_hora || v.fechaHora);
        return fechaVenta >= fechaLimiteISO || Number(v.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.ventas, ventasFiltradas);
    }
    
    localStorage.setItem(ULTIMA_LIMPIEZA_KEY, hoy);
  } catch (error) {
    console.error("❌ Error limpiando ventas antiguas:", error);
  }
}

function extraerFechaISO(valor) {
  if (!valor) return "";

  if (typeof valor === "string") {
    if (valor.indexOf("T") > -1) return valor.split("T")[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(valor)) return valor;
  }

  var fecha = new Date(valor);
  if (Number.isNaN(fecha.getTime())) return "";

  var y = fecha.getFullYear();
  var m = String(fecha.getMonth() + 1).padStart(2, "0");
  var d = String(fecha.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + d;
}

// ============================================
// INICIALIZAR BASE DE DATOS
// ============================================
async function initDatabase() {
  // SQLite en APK, fallback localStorage en cualquier entorno
  if (!window.Capacitor || !window.Capacitor.isNativePlatform()) {
    activarModoLocal("Entorno no nativo");
    return true;
  }

  var sqliteGlobal = window.SQLite;
  if (!sqliteGlobal || typeof sqliteGlobal.createDatabase !== "function") {
    activarModoLocal("Plugin SQLite no expuesto como window.SQLite.createDatabase");
    return true;
  }

  try {
    // Abrir o crear base de datos
    db = await sqliteGlobal.createDatabase({
      database: "posmovil.db",
    });

    // Crear tablas si no existen
    await db.execute(
      "CREATE TABLE IF NOT EXISTS productos (codigo TEXT PRIMARY KEY, nombre TEXT, precio REAL, disponibilidad INTEGER)"
    );

    await db.execute(
      "CREATE TABLE IF NOT EXISTS ventas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, factura_id TEXT, fecha_hora TEXT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, precio REAL, subtotal REAL, efectivo REAL, transferencia REAL, synced INTEGER DEFAULT 0)"
    );

    // Tabla para mermas pendientes
    await db.execute(
      "CREATE TABLE IF NOT EXISTS mermas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT, nombre TEXT, cantidad INTEGER, synced INTEGER DEFAULT 0)"
    );

    // Migración para instalaciones viejas (columna mal escrita: efectividad)
    try {
      await db.execute(
        "ALTER TABLE ventas_pending ADD COLUMN efectivo REAL DEFAULT 0"
      );
    } catch (_) {
      // Ya existe o la tabla no requiere migración
    }

    try {
      await db.execute(
        "UPDATE ventas_pending SET efectivo = COALESCE(efectivo, efectividad, 0)"
      );
    } catch (_) {
      // Si no existe 'efectividad' en esta instalación, ignorar
    }

    // Tabla de configuración (cache del servidor, etc.)
    await db.execute(
      "CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT, timestamp INTEGER)"
    );

    console.log("✅ Base de datos SQLite inicializada");
    storageMode = "sqlite";
    return true;
  } catch (error) {
    console.error("❌ Error inicializando SQLite:", error);
    activarModoLocal("Falló init SQLite");
    return true;
  }
}

// ============================================
// VERIFICAR SI HAY CONEXIÓN
// ============================================
function isOnline() {
  return navigator.onLine;
}

// ============================================
// OBTENER PRODUCTOS DESDE SQLite
// ============================================
async function getProductosLocal() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.productos, []);
  }

  try {
    var result = await db.execute(
      "SELECT codigo, nombre, precio, disponibilidad FROM productos ORDER BY nombre"
    );
    return result.values || [];
  } catch (error) {
    console.error("❌ Error obteniendo productos:", error);
    activarModoLocal("Error consultando productos SQLite");
    return leerLocal(LS_KEYS.productos, []);
  }
}

// ============================================
// GUARDAR PRODUCTOS EN SQLite (para cache offline)
// ============================================
async function syncProductosLocal(productos) {
  var normalizados = (productos || []).map(normalizarProducto);

  if (storageMode !== "sqlite" || !db) {
    return guardarLocal(LS_KEYS.productos, normalizados);
  }

  try {
    // Limpiar tabla existente
    await db.execute("DELETE FROM productos");

    // Insertar cada producto
    for (var i = 0; i < normalizados.length; i++) {
      var p = normalizados[i];
      await db.execute(
        "INSERT OR REPLACE INTO productos (codigo, nombre, precio, disponibilidad) VALUES (?, ?, ?, ?)",
        [p.codigo, p.nombre, p.precio || 0, p.disponibilidad || 0]
      );
    }

    console.log("✅ " + normalizados.length + " productos guardados en SQLite");
    return true;
  } catch (error) {
    console.error("❌ Error guardando productos:", error);
    activarModoLocal("Error guardando productos en SQLite");
    return guardarLocal(LS_KEYS.productos, normalizados);
  }
}

// ============================================
// GUARDAR VENTA EN SQLite (modo offline)
// ============================================
async function guardarVentaOffline(venta) {
  if (!venta || !Array.isArray(venta.productos) || venta.productos.length === 0) {
    return false;
  }

  if (storageMode !== "sqlite" || !db) {
    var actuales = leerLocal(LS_KEYS.ventas, []);
    var lineas = construirLineasVenta(venta, 0);
    return guardarLocal(LS_KEYS.ventas, actuales.concat(lineas));
  }

  try {
    for (var i = 0; i < venta.productos.length; i++) {
      var item = venta.productos[i];

      try {
        await db.execute(
          "INSERT INTO ventas_pending (factura_id, fecha_hora, codigo_producto, nombre, cantidad, precio, subtotal, efectivo, transferencia, synced) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)",
          [
            venta.facturaId,
            venta.fechaHora,
            item.codigo,
            item.nombre,
            item.cantidad,
            item.precio,
            item.cantidad * item.precio,
            venta.pago.efectivo,
            venta.pago.transferencia,
          ]
        );
      } catch (insertError) {
        // Compatibilidad con esquemas viejos que usan 'efectividad'
        if (
          insertError &&
          insertError.message &&
          insertError.message.indexOf("efectivo") !== -1
        ) {
          await db.execute(
            "INSERT INTO ventas_pending (factura_id, fecha_hora, codigo_producto, nombre, cantidad, precio, subtotal, efectividad, transferencia, synced) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)",
            [
              venta.facturaId,
              venta.fechaHora,
              item.codigo,
              item.nombre,
              item.cantidad,
              item.precio,
              item.cantidad * item.precio,
              venta.pago.efectivo,
              venta.pago.transferencia,
            ]
          );
        } else {
          throw insertError;
        }
      }
    }

    console.log("✅ Venta guardada offline");
    return true;
  } catch (error) {
    console.error("❌ Error guardando venta offline:", error);
    activarModoLocal("Error guardando ventas en SQLite");
    var actuales2 = leerLocal(LS_KEYS.ventas, []);
    var lineas2 = construirLineasVenta(venta, 0);
    return guardarLocal(LS_KEYS.ventas, actuales2.concat(lineas2));
  }
}

// ============================================
// GUARDAR MERMA OFFLINE (SQLite o localStorage fallback)
// ============================================
async function guardarMermaOffline(mermas) {
  if (!Array.isArray(mermas) || mermas.length === 0) return false;

  if (storageMode === "sqlite" && db) {
    try {
      for (const merma of mermas) {
        await db.execute(
          "INSERT INTO mermas_pending (codigo, nombre, cantidad, synced) VALUES (?, ?, ?, 0)",
          [merma.codigo, merma.nombre, merma.cantidad]
        );
      }
      console.log("✅ Mermas guardadas offline");
      return true;
    } catch (error) {
      console.error("❌ Error guardando mermas offline (SQLite):", error);
      return false;
    }
  } else {
    // Fallback a localStorage
    const actuales = leerLocal("posmovil_mermas_pending", []);
    const nuevas = mermas.map(m => ({
      ...m,
      id: Date.now() + Math.floor(Math.random() * 1000),
      synced: 0,
    }));
    guardarLocal("posmovil_mermas_pending", actuales.concat(nuevas));
    return true;
  }
}

// ============================================
// OBTENER MERMAS PENDIENTES
// ============================================
async function getMermasPendientes() {
  if (storageMode === "sqlite" && db) {
    try {
      const result = await db.execute("SELECT * FROM mermas_pending WHERE synced = 0");
      return result.values || [];
    } catch (error) {
      console.error("❌ Error obteniendo mermas pendientes:", error);
      return [];
    }
  } else {
    const mermas = leerLocal("posmovil_mermas_pending", []);
    return mermas.filter(m => m.synced === 0);
  }
}

// ============================================
// SINCRONIZAR MERMAS AL SERVIDOR
// ============================================
async function sincronizarMermas() {
  const pendientes = await getMermasPendientes();
  if (pendientes.length === 0) return { success: true, synced: 0 };

  if (!window.SERVER_URL) {
    const serverUrl = await descubrirServidor();
    if (!serverUrl) return { success: false, error: "Servidor no encontrado" };
    window.SERVER_URL = serverUrl;
  }

  try {
    const response = await fetch(window.SERVER_URL + "/api/mermas", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        mermas: pendientes.map(m => ({
          codigo: m.codigo,
          producto: m.nombre,
          cantidad: m.cantidad,
        })),
      }),
    });

    if (!response.ok) throw new Error("HTTP " + response.status);

    // Marcar como sincronizadas
    if (storageMode === "sqlite" && db) {
      for (const m of pendientes) {
        await db.execute("UPDATE mermas_pending SET synced = 1 WHERE id = ?", [m.id]);
      }
    } else {
      const actuales = leerLocal("posmovil_mermas_pending", []);
      const actualizadas = actuales.map(m =>
        pendientes.some(p => p.id === m.id) ? { ...m, synced: 1 } : m
      );
      guardarLocal("posmovil_mermas_pending", actualizadas);
    }

    console.log("✅ " + pendientes.length + " mermas sincronizadas");
    return { success: true, synced: pendientes.length };
  } catch (error) {
    console.error("❌ Error sincronizando mermas:", error);
    return { success: false, error: error.message };
  }
}

// ============================================
// GUARDAR VENTA ONLINE EN HISTÓRICO LOCAL
// (para que el resumen del teléfono incluya online + offline)
// ============================================
async function guardarVentaOnlineLocal(venta, facturaIdServidor) {
  if (!venta || !Array.isArray(venta.productos) || venta.productos.length === 0) {
    return false;
  }

  var lineasVenta = construirLineasVenta(venta, 1, facturaIdServidor);

  if (storageMode !== "sqlite" || !db) {
    var actuales = leerLocal(LS_KEYS.ventas, []);
    return guardarLocal(LS_KEYS.ventas, actuales.concat(lineasVenta));
  }

  try {
    for (var i = 0; i < lineasVenta.length; i++) {
      var l = lineasVenta[i];
      try {
        await db.execute(
          "INSERT INTO ventas_pending (factura_id, fecha_hora, codigo_producto, nombre, cantidad, precio, subtotal, efectivo, transferencia, synced) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)",
          [
            l.factura_id,
            l.fecha_hora,
            l.codigo_producto,
            l.nombre,
            l.cantidad,
            l.precio,
            l.subtotal,
            l.efectivo,
            l.transferencia,
          ]
        );
      } catch (insertError) {
        if (
          insertError &&
          insertError.message &&
          insertError.message.indexOf("efectivo") !== -1
        ) {
          await db.execute(
            "INSERT INTO ventas_pending (factura_id, fecha_hora, codigo_producto, nombre, cantidad, precio, subtotal, efectividad, transferencia, synced) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)",
            [
              l.factura_id,
              l.fecha_hora,
              l.codigo_producto,
              l.nombre,
              l.cantidad,
              l.precio,
              l.subtotal,
              l.efectivo,
              l.transferencia,
            ]
          );
        } else {
          throw insertError;
        }
      }
    }

    return true;
  } catch (error) {
    console.error("❌ Error guardando venta online local:", error);
    activarModoLocal("Error guardando histórico online en SQLite");
    var actuales2 = leerLocal(LS_KEYS.ventas, []);
    return guardarLocal(LS_KEYS.ventas, actuales2.concat(lineasVenta));
  }
}

// ============================================
// OBTENER VENTAS PENDIENTES DE SYNC
// ============================================
async function getVentasPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.ventas, []).filter(function(v) {
      return Number(v.synced || 0) === 0;
    });
  }

  try {
    var result = await db.execute(
      "SELECT * FROM ventas_pending WHERE synced = 0 ORDER BY id"
    );
    return result.values || [];
  } catch (error) {
    console.error("❌ Error obteniendo ventas pendientes:", error);
    activarModoLocal("Error leyendo pendientes SQLite");
    return leerLocal(LS_KEYS.ventas, []).filter(function(v) {
      return Number(v.synced || 0) === 0;
    });
  }
}

// ============================================
// OBTENER RESUMEN OFFLINE DESDE SQLite (ventas del día)
// ============================================
async function getResumenOffline(fechaISO) {
  try {
    var ventas = [];

    if (storageMode === "sqlite" && db) {
      var result = await db.execute(
        "SELECT * FROM ventas_pending ORDER BY id"
      );
      ventas = result.values || [];
    } else {
      ventas = leerLocal(LS_KEYS.ventas, []);
    }

    var fechaFiltro = fechaISO || fechaLocalHoyISO();
    ventas = ventas.filter(function(v) {
      return extraerFechaISO(v.fecha_hora || v.fechaHora) === fechaFiltro;
    });

    if (ventas.length === 0) {
      return {
        totalIngresado: 0,
        efectivo: 0,
        transferencia: 0,
        productosVendidos: []
      };
    }

    var totalIngresado = 0;
    var totalEfectivo = 0;
    var totalTransferencia = 0;
    var productosAgrupados = {};
    var facturasVistas = {};

    for (var i = 0; i < ventas.length; i++) {
      var v = ventas[i];
      var subtotal = v.subtotal || 0;
      var efectivo = v.efectivo || v.efectividad || 0;
      var transferencia = v.transferencia || 0;
      var facturaKey = v.factura_id || ("row_" + i);

      totalIngresado += subtotal;

      // Evitar duplicar el pago por cada línea de la misma factura
      if (!facturasVistas[facturaKey]) {
        totalEfectivo += efectivo;
        totalTransferencia += transferencia;
        facturasVistas[facturaKey] = true;
      }

      // Agrupar por producto
      var key = v.codigo_producto;
      if (!productosAgrupados[key]) {
        productosAgrupados[key] = {
          codigo: v.codigo_producto,
          nombre: v.nombre,
          cantidad: 0,
          total: 0
        };
      }
      productosAgrupados[key].cantidad += v.cantidad;
      productosAgrupados[key].total += subtotal;
    }

    var productosVendidos = [];
    var keys = Object.keys(productosAgrupados);
    for (var j = 0; j < keys.length; j++) {
      var productoKey = keys[j];
      productosVendidos.push(productosAgrupados[productoKey]);
    }

    // Ordenar por cantidad descendente
    productosVendidos.sort(function(a, b) { return b.cantidad - a.cantidad; });

    return {
      totalIngresado: Math.round(totalIngresado * 1000) / 1000,
      efectivo: Math.round(totalEfectivo * 1000) / 1000,
      transferencia: Math.round(totalTransferencia * 1000) / 1000,
      productosVendidos: productosVendidos
    };
  } catch (error) {
    console.error("❌ Error en getResumenOffline:", error);
    return {
      totalIngresado: 0,
      efectivo: 0,
      transferencia: 0,
      productosVendidos: []
    };
  }
}

// ============================================
// CONTAR VENTAS PENDIENTES
// ============================================
async function contarVentasPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return contarFacturasPendientesLocal(leerLocal(LS_KEYS.ventas, []));
  }

  try {
    var result = await db.execute(
      "SELECT COUNT(DISTINCT factura_id) as count FROM ventas_pending WHERE synced = 0"
    );
    return result.values ? result.values[0].count : 0;
  } catch (error) {
    console.error("❌ Error contando pendientes:", error);
    activarModoLocal("Error contando pendientes SQLite");
    return contarFacturasPendientesLocal(leerLocal(LS_KEYS.ventas, []));
  }
}

// ============================================
// SYNCRONIZAR VENTAS AL SERVIDOR
// ============================================
async function sincronizarVentas() {
  var pendientes = await getVentasPendientes();
  if (pendientes.length === 0) {
    return { success: true, sincronizadas: 0 };
  }

  // Intentar auto-descubrir el servidor si no está configurado
  if (!window.SERVER_URL) {
    console.log("🔍 Buscando servidor en la red local...");
    var serverUrlDescubierto = await descubrirServidor();
    if (serverUrlDescubierto) {
      console.log("✅ Servidor encontrado: " + serverUrlDescubierto);
      window.SERVER_URL = serverUrlDescubierto;
    } else {
      return { success: false, error: "No se encontró el servidor. verifica que la PC esté conectada y el servidor activo." };
    }
  }

  // Enviar formato plano por línea (compatible con el backend actual)
  var ventas = [];
  for (var i = 0; i < pendientes.length; i++) {
    var v = pendientes[i];
    ventas.push({
      facturaId: v.factura_id,
      fechaHora: v.fecha_hora,
      codigoProducto: v.codigo_producto,
      nombre: v.nombre,
      cantidad: v.cantidad,
      precio: v.precio,
      subtotal: v.subtotal,
      efectivo: v.efectivo || v.efectividad || 0,
      transferencia: v.transferencia || 0,
    });
  }

  try {
    var response = await fetch(window.SERVER_URL + "/api/sync", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ ventas: ventas }),
    });

    if (!response.ok) {
      throw new Error("HTTP " + response.status);
    }

    var result = await response.json();

    // Guardar referencia de última factura del servidor para uso offline
    if (result.ultimaFacturaId) {
      try {
        localStorage.setItem('posmovil_ultima_factura_referencia', result.ultimaFacturaId);
        // También extraer y guardar el prefijo y sufijo
        var refId = result.ultimaFacturaId.toString();
        if (refId.length >= 10) {
          var prefijoRef = refId.substring(0,5);
          var sufijoRef = parseInt(refId.substring(5,10)) || 0;
          localStorage.setItem('posmovil_ultimo_prefijo', prefijoRef);
          localStorage.setItem('posmovil_ultimo_sufijo', sufijoRef.toString());
          // Guardar también la fecha de hoy para calcular días transcurridos
          var hoyISO = new Date().getFullYear() + '-' +
                     String(new Date().getMonth()+1).padStart(2,'0') + '-' +
                     String(new Date().getDate()).padStart(2,'0');
          localStorage.setItem('posmovil_fecha_referencia', hoyISO);
          console.log('✅ Referencia de factura guardada: ' + refId + ' (prefijo: ' + prefijoRef + ', sufijo: ' + sufijoRef + ', fecha: ' + hoyISO + ')');
        }
      } catch(e) {
        console.error('Error guardando referencia de factura:', e);
      }
    }

    // Marcar como sincronizadas
    if (storageMode === "sqlite" && db) {
      for (var j = 0; j < pendientes.length; j++) {
        await db.execute(
          "UPDATE ventas_pending SET synced = 1 WHERE id = ?",
          [pendientes[j].id]
        );
      }
    } else {
      var todas = leerLocal(LS_KEYS.ventas, []);
      var idsPendientes = {};
      for (var k = 0; k < pendientes.length; k++) {
        idsPendientes[String(pendientes[k].id)] = true;
      }

      for (var x = 0; x < todas.length; x++) {
        if (idsPendientes[String(todas[x].id)]) {
          todas[x].synced = 1;
        }
      }
      guardarLocal(LS_KEYS.ventas, todas);
    }

    console.log("✅ " + ventas.length + " líneas sincronizadas");
    return {
      success: true,
      sincronizadas: result.sincronizadas || ventas.length,
      errores: result.errores || [],
    };
  } catch (error) {
    console.error("❌ Error sincronizando:", error);
    return { success: false, error: error.message };
  }
}

// ============================================
// AUTO-DESCUBRIR SERVIDOR EN RED LOCAL
// ============================================
const SERVIDOR_CACHE_KEY = "servidor_cache";
const SERVIDOR_CACHE_TTL = 1000 * 60 * 30; // 30 minutos

async function descubrirServidor() {
  // 1. Primero verificar si hay servidor cacheado válido
  var servidorCacheado = await obtenerServidorCacheado();
  if (servidorCacheado) {
    console.log("🎯 [descubrirServidor] Usando servidor cacheado:", servidorCacheado);
    return servidorCacheado;
  }

  // 2. Si no hay cache, escanear la red
  var servidorEncontrado = await escanearRed();

  // 3. Si se encontró, guardar en cache
  if (servidorEncontrado) {
    await guardarServidorCacheado(servidorEncontrado);
    console.log("✅ [descubrirServidor] Servidor encontrado y cacheado:", servidorEncontrado);
  }

  return servidorEncontrado;
}

// ============================================
// OBTENER SERVIDOR DESDE CACHE
// ============================================
async function obtenerServidorCacheado() {
  if (storageMode !== "sqlite" || !db) {
    var cfgLocal = leerLocal(LS_KEYS.config, {});
    var filaLocal = cfgLocal[SERVIDOR_CACHE_KEY];
    if (!filaLocal) return null;

    var ahoraLocal = Date.now();
    if (ahoraLocal - (filaLocal.timestamp || 0) < SERVIDOR_CACHE_TTL) {
      return filaLocal.valor;
    }

    delete cfgLocal[SERVIDOR_CACHE_KEY];
    guardarLocal(LS_KEYS.config, cfgLocal);
    return null;
  }
  
  try {
    var result = await db.execute(
      "SELECT valor, timestamp FROM config WHERE clave = ?",
      [SERVIDOR_CACHE_KEY]
    );
    
    if (result.values && result.values.length > 0) {
      var fila = result.values[0];
      var timestamp = fila.timestamp;
      var ahora = Date.now();
      
      // Verificar si el cache aún es válido (menos de 30 min)
      if (ahora - timestamp < SERVIDOR_CACHE_TTL) {
        return fila.valor;
      }
      
      // Cache expirado, eliminar
      await db.execute("DELETE FROM config WHERE clave = ?", [SERVIDOR_CACHE_KEY]);
    }
    
    return null;
  } catch (e) {
    console.error("❌ [obtenerServidorCacheado] Error:", e);
    return null;
  }
}

// ============================================
// GUARDAR SERVIDOR EN CACHE
// ============================================
async function guardarServidorCacheado(url) {
  if (storageMode !== "sqlite" || !db) {
    var cfgLocal = leerLocal(LS_KEYS.config, {});
    cfgLocal[SERVIDOR_CACHE_KEY] = { valor: url, timestamp: Date.now() };
    guardarLocal(LS_KEYS.config, cfgLocal);
    console.log("✅ [guardarServidorCacheado] Cache local guardado:", url);
    return;
  }
  
  try {
    await db.execute(
      "INSERT OR REPLACE INTO config (clave, valor, timestamp) VALUES (?, ?, ?)",
      [SERVIDOR_CACHE_KEY, url, Date.now()]
    );
    console.log("✅ [guardarServidorCacheado] Cache guardado:", url);
  } catch (e) {
    console.error("❌ [guardarServidorCacheado] Error:", e);
  }
}

// ============================================
// LIMPIAR CACHE DEL SERVIDOR
// ============================================
async function limpiarServidorCacheado() {
  if (storageMode !== "sqlite" || !db) {
    var cfgLocal = leerLocal(LS_KEYS.config, {});
    delete cfgLocal[SERVIDOR_CACHE_KEY];
    guardarLocal(LS_KEYS.config, cfgLocal);
    console.log("🗑️ [limpiarServidorCacheado] Cache local eliminado");
    return;
  }
  
  try {
    await db.execute(
      "DELETE FROM config WHERE clave = ?",
      [SERVIDOR_CACHE_KEY]
    );
    console.log("🗑️ [limpiarServidorCacheado] Cache eliminado");
  } catch (e) {
    console.error("❌ [limpiarServidorCacheado] Error:", e);
  }
}

// ============================================
// ESCANEAR RED LOCAL (soporta 192.168.x.x y 10.x.x.x)
// ============================================
// ESCANEAR UN RANGO ESPECÍFICO DE RED
// ============================================
async function escanearRango(baseIP, puerto, ipLocal) {
  console.log("🔍 [escanearRango] Escaneando: " + baseIP + ".x");

  var batchSize = 25;
  var timeoutMs = 800;
  var encontrado = null;

  for (var batchStart = 1; batchStart <= 254; batchStart += batchSize) {
    var promises = [];

    for (var i = 0; i < batchSize && (batchStart + i) <= 254; i++) {
      let ip = baseIP + "." + (batchStart + i);
      if (ipLocal && ip === ipLocal) continue;

      let url = "http://" + ip + ":" + puerto + "/api/estado-publico";

      promises.push(
        fetch(url, { method: "GET", signal: AbortSignal.timeout(timeoutMs) })
          .then(function(res) {
            if (res.ok) return { url: "http://" + ip + ":" + puerto };
            return null;
          })
          .catch(function() { return null; })
      );
    }

    var resultados = await Promise.all(promises);

    for (var r = 0; r < resultados.length; r++) {
      if (resultados[r]) {
        encontrado = resultados[r].url;
        console.log("✅ [escanearRango] Servidor encontrado: " + encontrado);
        return encontrado;
      }
    }
  }

  return null;
}

// ============================================
async function escanearRed() {
  cancelarEscaneo = false; // Reiniciar bandera de cancelación al iniciar
  var puerto = 3000;
  var ipLocal = await obtenerIPLocal();

  console.log("📱 IP del dispositivo: " + (ipLocal || "NO DETECTADA"));

  // Determinar rangos a escanear
  var rangos = [];

  // Si se detecta IP, agregar ese rango primero
  if (ipLocal) {
    var partesIP = ipLocal.split(".");
    if (partesIP.length === 4) {
      if (partesIP[0] === "192" && partesIP[1] === "168") {
        rangos.push({ base: partesIP[0] + "." + partesIP[1] + "." + partesIP[2], origen: "auto-detectado" });
      } else if (partesIP[0] === "10") {
        rangos.push({ base: partesIP[0] + "." + partesIP[1] + "." + partesIP[2], origen: "auto-detectado" });
      } else if (partesIP[0] === "172" && partesIP[1] >= 16 && partesIP[1] <= 31) {
        rangos.push({ base: partesIP[0] + "." + partesIP[1] + "." + partesIP[2], origen: "auto-detectado" });
      }
    }
  }

  // Agregar siempre los rangos más comunes (wifi + tethering Android/iPhone)
  var rangosComunes = [
    { base: "192.168.1", origen: "wifi (común)" },
    { base: "192.168.43", origen: "Android tethering" },
    { base: "192.168.0", origen: "wifi alternativo" },
    { base: "10.0.0", origen: "tethering/otro" },
    { base: "10.0.1", origen: "tethering/otro" },
    { base: "10.1.1", origen: "tethering" },
    { base: "10.10.0", origen: "tethering" },
    { base: "10.100.1", origen: "tethering" },
    { base: "10.200.1", origen: "tethering" },
    { base: "10.225.81", origen: "tethering (caso actual)" },
    { base: "172.20.10", origen: "iPhone tethering" },
  ];

  // Agregar los que no existen ya
  var seen = {};
  for (var i = 0; i < rangos.length; i++) {
    seen[rangos[i].base] = true;
  }

  var maxRangos = 15;
  var agregados = rangos.length;

  for (var j = 0; j < rangosComunes.length && agregados < maxRangos; j++) {
    if (!seen[rangosComunes[j].base]) {
      rangos.push(rangosComunes[j]);
      seen[rangosComunes[j].base] = true;
      agregados++;
    }
  }

  console.log("🔍 [escanearRed] Rangos a escanear:", rangos.map(function(r) { return r.base + " (" + r.origen + ")"; }).join(", "));

  // Escanear cada rango
  for (var k = 0; k < rangos.length; k++) {
    // Verificar si se solicitó cancelar
    if (cancelarEscaneo) {
      console.log("🛑 [escanearRed] Escaneo cancelado por usuario");
      cancelarEscaneo = false; // Reiniciar bandera
      return null;
    }

    var encontrado = await escanearRango(rangos[k].base, puerto, ipLocal);

    if (encontrado) {
      return encontrado;
    }
  }

  console.error("❌ [escanearRed] No se encontró servidor en ningún rango");
  return null;
}

// ============================================
// OBTENER IP LOCAL DEL DISPOSITIVO (MEJORADO)
// ============================================
async function obtenerIPLocal() {
  // MÉTODO 1: RTCPeerConnection (original)
  try {
    var pc = new RTCPeerConnection({ iceServers: [] });
    pc.createDataChannel("");

    var ipRTC = await new Promise(function(resolve) {
      pc.onicecandidate = function(e) {
        if (e.candidate) {
          var match = e.candidate.candidate.match(/([0-9.]+):/);
          if (match && match[1] !== "127.0.0.1" && match[1] !== "0.0.0.0") {
            resolve(match[1]);
            pc.close();
            return;
          }
        }
      };

      setTimeout(function() {
        resolve(null);
        pc.close();
      }, 2000);

      pc.createOffer().catch(function() { resolve(null); });
    });

    if (ipRTC) {
      console.log("📱 [obtenerIPLocal] RTCPeerConnection:", ipRTC);
      return ipRTC;
    }
  } catch (e) {
    console.warn("⚠️ [obtenerIPLocal] RTCPeerConnection falló:", e.message);
  }

  // MÉTODO 2: Network Information API (disponible en algunos browsers)
  if (navigator.connection && navigator.connection.ipAddress) {
    var ipNetInfo = navigator.connection.ipAddress;
    if (ipNetInfo && ipNetInfo !== "127.0.0.1") {
      console.log("📱 [obtenerIPLocal] Network Info API:", ipNetInfo);
      return ipNetInfo;
    }
  }

  // MÉTODO 3: Hacer request a un servicio externo y obtener IP del dispositivo
  // Esto nos da la IP pública, pero también podemos obtener la IP del cliente desde headers del servidor
  // Por ahora, intentamos detectar rangos comunes manual

  console.warn("⚠️ [obtenerIPLocal] Todos los métodos automáticos fallaron");
  console.warn("⚠️ [obtenerIPLocal] Nota: Si estás usando tethering, la IP podría ser 192.168.43.x o 10.x.x.x");

  return null;
}

// ============================================
// EXPORTAR COMO MÓDULO GLOBAL
// ============================================
window.Database = {
  initDatabase: initDatabase,
  getStorageMode: function() { return storageMode; },
  isOnline: isOnline,
  getProductosLocal: getProductosLocal,
  syncProductosLocal: syncProductosLocal,
  guardarVentaOffline: guardarVentaOffline,
  guardarVentaOnlineLocal: guardarVentaOnlineLocal,
  getVentasPendientes: getVentasPendientes,
  contarVentasPendientes: contarVentasPendientes,
  sincronizarVentas: sincronizarVentas,
  getResumenOffline: getResumenOffline,
  limpiarVentasAntiguas: limpiarVentasAntiguas,
  obtenerServidorCacheado: obtenerServidorCacheado,
  guardarServidorCacheado: guardarServidorCacheado,
  limpiarServidorCacheado: limpiarServidorCacheado,
  descubrirServidor: descubrirServidor,
  // Mermas
  guardarMermaOffline: guardarMermaOffline,
  sincronizarMermas: sincronizarMermas,
  getMermasPendientes: getMermasPendientes,
};

console.log("📦 Database module loaded");
