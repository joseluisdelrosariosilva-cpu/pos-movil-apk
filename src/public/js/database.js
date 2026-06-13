// ============================================
// DATABASE MODULE - SQLite para modo offline
// ============================================
// Versión vanilla para ejecutarse en webview del APK

var db = null;
var storageMode = "none"; // sqlite | local
var cancelarEscaneo = false; // Bandera para cancelar escaneo de red
var _servidorEncontrado = null; // Compartido entre escaneos paralelos

const LS_KEYS = {
  productos: "posmovil_productos",
  ventas: "posmovil_ventas_pending",
  mermas: "posmovil_mermas_pending",
  entrada_productos: "posmovil_entrada_productos_pending",
  abastecer: "posmovil_abastecer_pending",
  gastos: "posmovil_gastos_pending",
  config: "posmovil_config",
  recetas: "posmovil_recetas",
  ingredientes: "posmovil_ingredientes",
  elaboraciones: "posmovil_elaboraciones_pending",
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

// Lee un string de localStorage con retrocompatibilidad:
//   - nuevo formato (guardarLocal → JSON.stringify) → JSON.parse
//   - formato legacy (raw string) → devuelve raw
function _leerString(key, fallback) {
  try {
    var raw = localStorage.getItem(key);
    if (raw === null) return fallback;
    try { return JSON.parse(raw); } catch (_) { return raw; }
  } catch (_) {
    return fallback;
  }
}

/**
 * Genera un ID único de factura basado en fecha como "4567800001" (prefijo + sufijo)
 * El prefijo son días desde 1900-01-01, el sufijo es secuencial por día.
 * Persiste en localStorage con guardarLocal() para detectar fallos de cuota.
 */
function generarFacturaId() {
  var now = new Date();
  var hoyISO = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0');

  // Leer referencia guardada del servidor Y su fecha (retrocompatible con datos legacy)
  var ultimaReferencia = _leerString('posmovil_ultima_factura_referencia');
  var fechaReferencia = _leerString('posmovil_fecha_referencia');

  // Si tenemos referencia Y fecha, usarlas
  if (ultimaReferencia && ultimaReferencia.length >= 10 && fechaReferencia) {
    var prefijoRef = ultimaReferencia.substring(0, 5);
    var sufijoRef = parseInt(ultimaReferencia.substring(5, 10)) || 0;
    var refDate = new Date(fechaReferencia);
    var todayDate = new Date(hoyISO);
    var msPerDay = 1000 * 60 * 60 * 24;
    var diasDiff = Math.floor((todayDate - refDate) / msPerDay);
    var prefijoHoy = String(Number(prefijoRef) + diasDiff).padStart(5, '0');

    if (diasDiff === 0) {
      // Mismo día que la referencia, incrementar sufijo
      var nuevoSufijo = sufijoRef + 1;
      var sufijo = String(nuevoSufijo).padStart(5, '0');
      var nuevoId = prefijoHoy + sufijo;

      if (!guardarLocal('posmovil_ultima_factura_referencia', nuevoId)) {
        console.error("🚨 [generarFacturaId] No se pudo guardar referencia — riesgo de ID duplicado");
      }
      guardarLocal('posmovil_ultimo_prefijo', prefijoHoy);
      guardarLocal('posmovil_ultimo_sufijo', String(nuevoSufijo));

      console.log('📋 FacturaID generado: ' + nuevoId + ' (mismo día que referencia)');
      return nuevoId;
    } else {
      // Nuevo día (o días), empezar sufijo en 1
      var nuevoId = prefijoHoy + "00001";

      if (!guardarLocal('posmovil_ultima_factura_referencia', nuevoId)) {
        console.error("🚨 [generarFacturaId] No se pudo guardar referencia — riesgo de ID duplicado");
      }
      guardarLocal('posmovil_ultimo_prefijo', prefijoHoy);
      guardarLocal('posmovil_ultimo_sufijo', '1');
      guardarLocal('posmovil_fecha_referencia', hoyISO);

      console.log('📋 FacturaID generado: ' + nuevoId + ' (nuevo día, díasDiff=' + diasDiff + ')');
      return nuevoId;
    }
  }

  // Fallback: sin referencia, usar la fecha actual (primera vez)
  var fechaBase = new Date(1900, 0, 1);
  var diffMs = now - fechaBase;
  var diffDias = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  var fechaSerial = diffDias + 2;
  var prefijoHoy = String(Math.floor(fechaSerial)).padStart(5, '0');
  var nuevoId = prefijoHoy + "00001";

  if (!guardarLocal('posmovil_ultima_factura_referencia', nuevoId)) {
    console.error("🚨 [generarFacturaId] No se pudo guardar referencia — riesgo de ID duplicado");
  }
  guardarLocal('posmovil_ultimo_prefijo', prefijoHoy);
  guardarLocal('posmovil_ultimo_sufijo', '1');
  guardarLocal('posmovil_fecha_referencia', hoyISO);

  console.log('📋 FacturaID generado (sin referencia): ' + nuevoId);
  return nuevoId;
}

function logSyncDebugAPK(mensaje, tipo) {
  if (tipo === "error") {
    console.error(mensaje);
  } else if (tipo === "warning") {
    console.warn(mensaje);
  } else {
    console.log(mensaje);
  }

  try {
    if (window && typeof window.logSyncAPK === "function") {
      window.logSyncAPK(mensaje, tipo);
    }
  } catch (_) {
    // Evitar romper flujo por logger visual
  }
}

// ═══ LOG VISUAL PARA DEBUG EN APK (temporal) ═══

/**
 * Helper para operaciones SQLite con fallback automático a localStorage.
 * Elimina la duplicación del guard (storageMode !== "sqlite" || !db)
 * y el try/catch con mismo fallback en ~20 funciones.
 *
 * @param {function(db): Promise<any>} operacion - Recibe db, ejecuta SQL
 * @param {function(): any} fallback - Se ejecuta si no hay SQLite o si falla
 */
async function conSQLite(operacion, fallback) {
  if (storageMode !== "sqlite" || !db) {
    return typeof fallback === "function" ? await fallback() : fallback;
  }
  try {
    return await operacion(db);
  } catch (error) {
    console.error("❌ Error en operación SQLite:", error);
    return typeof fallback === "function" ? await fallback() : fallback;
  }
}

function asegurarStoreLocal() {
  if (!localStorage.getItem(LS_KEYS.productos)) guardarLocal(LS_KEYS.productos, []);
  if (!localStorage.getItem(LS_KEYS.ventas)) guardarLocal(LS_KEYS.ventas, []);
  if (!localStorage.getItem(LS_KEYS.mermas)) guardarLocal(LS_KEYS.mermas, []);
  if (!localStorage.getItem(LS_KEYS.entrada_productos)) guardarLocal(LS_KEYS.entrada_productos, []);
  if (!localStorage.getItem(LS_KEYS.abastecer)) guardarLocal(LS_KEYS.abastecer, []);
  if (!localStorage.getItem(LS_KEYS.gastos)) guardarLocal(LS_KEYS.gastos, []);
  if (!localStorage.getItem(LS_KEYS.config)) guardarLocal(LS_KEYS.config, {});
  if (!localStorage.getItem(LS_KEYS.recetas)) guardarLocal(LS_KEYS.recetas, []);
  if (!localStorage.getItem(LS_KEYS.ingredientes)) guardarLocal(LS_KEYS.ingredientes, []);
  if (!localStorage.getItem(LS_KEYS.elaboraciones)) guardarLocal(LS_KEYS.elaboraciones, []);
}

function activarModoLocal(motivo) {
  storageMode = "local";
  db = null;
  asegurarStoreLocal();
  console.warn("📦 [DB] Modo localStorage activo:", motivo);
}

/**
 * Crea un wrapper que adapta la API de @capacitor-community/sqlite v8
 * (Capacitor.Plugins.CapacitorSQLite) a la interfaz que espera el resto del código:
 *   db.execute(sql, params?)  → rutea a run() / query() / execute() según el caso
 *   db.query(sql, params?)    → rutea a query()
 */
function crearWrapperSQLite(plugin, dbName) {
  return {
    execute: async function (sql, params) {
      var trimmed = sql.trim().toUpperCase();

      // 1) SELECT / WITH → query() (el Android plugin exige values[])
      if (trimmed.startsWith("SELECT") || trimmed.startsWith("WITH")) {
        var res = await plugin.query({
          database: dbName,
          statement: sql,
          values: Array.isArray(params) ? params : [],
        });
        return { values: res.values || [] };
      }

      // 2) Con parámetros → run()
      if (params !== undefined && params !== null) {
        var res = await plugin.run({
          database: dbName,
          statement: sql,
          values: Array.isArray(params) ? params : [params],
        });
        return { changes: res.changes || { changes: 0 } };
      }

      // 3) DDL / batch → execute()
      var res = await plugin.execute({ database: dbName, statements: sql });
      return { changes: res.changes || { changes: 0 } };
    },

    query: async function (sql, params) {
      var res = await plugin.query({
        database: dbName,
        statement: sql,
        values: params || [],
      });
      return { values: res.values || [] };
    },
  };
}

/**
 * Crea TODAS las tablas del esquema SQLite.
 * Extraída para no duplicar entre el path de Capacitor 8 y el legacy.
 */
async function crearTablasSQLite() {
  await db.execute(
    "CREATE TABLE IF NOT EXISTS productos (codigo TEXT PRIMARY KEY, nombre TEXT, precio REAL, disponibilidad INTEGER)"
  );
  await db.execute(
    "CREATE TABLE IF NOT EXISTS ventas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, factura_id TEXT, fecha_hora TEXT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, precio REAL, subtotal REAL, efectivo REAL, transferencia REAL, synced INTEGER DEFAULT 0)"
  );
  // Migración para instalaciones viejas (columna mal escrita: efectividad)
  try {
    await db.execute("ALTER TABLE ventas_pending ADD COLUMN efectivo REAL DEFAULT 0");
  } catch (_) {}
  try {
    await db.execute("UPDATE ventas_pending SET efectivo = COALESCE(efectivo, efectividad, 0)");
  } catch (_) {}
  await db.execute("CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT, timestamp INTEGER)");
  await db.execute("CREATE TABLE IF NOT EXISTS mermas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, fecha_hora TEXT, synced INTEGER DEFAULT 0)");
  await db.execute("CREATE TABLE IF NOT EXISTS entrada_productos_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT, nombre TEXT, cantidad INTEGER, precio_venta REAL, precio_costo REAL, fecha_hora TEXT, synced INTEGER DEFAULT 0)");
  await db.execute("CREATE TABLE IF NOT EXISTS abastecer_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, fecha_hora TEXT, synced INTEGER DEFAULT 0)");
  await db.execute("CREATE TABLE IF NOT EXISTS gastos_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, descripcion TEXT, monto REAL, synced INTEGER DEFAULT 0)");
  await db.execute("CREATE TABLE IF NOT EXISTS recetas (nombre TEXT PRIMARY KEY, cant_lote REAL, precio_venta REAL, precio_costo REAL)");
  try {
    await db.execute("ALTER TABLE recetas ADD COLUMN precio_costo REAL DEFAULT 0");
  } catch (_) {}
  await db.execute("CREATE TABLE IF NOT EXISTS ingredientes (ingrediente TEXT, cantidad REAL, unidad TEXT, receta TEXT)");
  await db.execute("CREATE TABLE IF NOT EXISTS elaboraciones_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre_receta TEXT, lotes INTEGER, cantidad_producida REAL, precio_venta REAL, fecha_hora TEXT, synced INTEGER DEFAULT 0)");
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
// LIMPIAR REGISTROS ANTIGUOS (más de 30 días y ya sincronizados)
// Limpia ventas, mermas, entrada_productos, abastecer y gastos
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
      // Ventas
      await db.execute("DELETE FROM ventas_pending WHERE synced = 1 AND date(fecha_hora) < ?", [fechaLimiteISO]);
      // Mermas
      await db.execute("DELETE FROM mermas_pending WHERE synced = 1 AND date(fecha_hora) < ?", [fechaLimiteISO]);
      // Entrada de productos
      await db.execute("DELETE FROM entrada_productos_pending WHERE synced = 1 AND date(fecha_hora) < ?", [fechaLimiteISO]);
      // Abastecimientos
      await db.execute("DELETE FROM abastecer_pending WHERE synced = 1 AND date(fecha_hora) < ?", [fechaLimiteISO]);
      // Elaboraciones
      await db.execute("DELETE FROM elaboraciones_pending WHERE synced = 1 AND date(fecha_hora) < ?", [fechaLimiteISO]);
      // Gastos (usa 'fecha' en vez de 'fecha_hora')
      await db.execute("DELETE FROM gastos_pending WHERE synced = 1 AND date(fecha) < ?", [fechaLimiteISO]);
      
      console.log("🗑️ Registros antiguos eliminados (antes del " + fechaLimiteISO + ")");
    } else {
      // localStorage: filtrar y guardar solo los recientes
      
      // Ventas
      var ventas = leerLocal(LS_KEYS.ventas, []);
      var ventasFiltradas = ventas.filter(function(v) {
        var fechaVenta = extraerFechaISO(v.fecha_hora || v.fechaHora);
        return fechaVenta >= fechaLimiteISO || Number(v.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.ventas, ventasFiltradas);
      
      // Mermas
      var mermas = leerLocal(LS_KEYS.mermas, []);
      var mermasFiltradas = mermas.filter(function(m) {
        return extraerFechaISO(m.fecha_hora) >= fechaLimiteISO || Number(m.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.mermas, mermasFiltradas);
      
      // Entrada de productos
      var entradas = leerLocal(LS_KEYS.entrada_productos, []);
      var entradasFiltradas = entradas.filter(function(e) {
        return extraerFechaISO(e.fecha_hora) >= fechaLimiteISO || Number(e.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.entrada_productos, entradasFiltradas);
      
      // Abastecimientos
      var abastecimientos = leerLocal(LS_KEYS.abastecer, []);
      var abastecimientosFiltrados = abastecimientos.filter(function(a) {
        return extraerFechaISO(a.fecha_hora) >= fechaLimiteISO || Number(a.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.abastecer, abastecimientosFiltrados);
      
      // Elaboraciones
      var elaboraciones = leerLocal(LS_KEYS.elaboraciones, []);
      var elaboracionesFiltradas = elaboraciones.filter(function(e) {
        return extraerFechaISO(e.fecha_hora) >= fechaLimiteISO || Number(e.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.elaboraciones, elaboracionesFiltradas);
      
      // Gastos (usa 'fecha' en vez de 'fecha_hora')
      var gastos = leerLocal(LS_KEYS.gastos, []);
      var gastosFiltrados = gastos.filter(function(g) {
        return extraerFechaISO(g.fecha) >= fechaLimiteISO || Number(g.synced || 0) === 0;
      });
      guardarLocal(LS_KEYS.gastos, gastosFiltrados);
    }
    
    localStorage.setItem(ULTIMA_LIMPIEZA_KEY, hoy);
  } catch (error) {
    console.error("❌ Error limpiando registros antiguos:", error);
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

/**
 * Genera timestamp en formato ISO con HORA LOCAL (sin UTC)
 * Ejemplo: "2026-05-31T23:00:00.000" (sin Z ni offset)
 * POLÍTICA: siempre usar hora local del dispositivo en toda la app
 *
 * ⚠️ DUPLICADA de src/utils/date.util.js (export const fechaLocalISO)
 *    Mantener ambas sincronizadas manualmente — mismo algoritmo, distinto runtime.
 */
function fechaLocalISO() {
  var f = new Date();
  var y = f.getFullYear();
  var m = String(f.getMonth() + 1).padStart(2, "0");
  var d = String(f.getDate()).padStart(2, "0");
  var hh = String(f.getHours()).padStart(2, "0");
  var mm = String(f.getMinutes()).padStart(2, "0");
  var ss = String(f.getSeconds()).padStart(2, "0");
  var ms = String(f.getMilliseconds()).padStart(3, "0");
  return y + "-" + m + "-" + d + "T" + hh + ":" + mm + ":" + ss + "." + ms;
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

  // ============================================
  // 1) Capacitor 8: Capacitor.Plugins.CapacitorSQLite
  // ============================================
  var plugin = null;
  if (window.Capacitor.Plugins && window.Capacitor.Plugins.CapacitorSQLite) {
    plugin = window.Capacitor.Plugins.CapacitorSQLite;
  }

  if (plugin) {
    try {
      await plugin.createConnection({
        database: "posmovil",
        version: 1,
        encrypted: false,
        mode: "no-encryption",
      });
      await plugin.open({ database: "posmovil" });
      db = crearWrapperSQLite(plugin, "posmovil");
      await crearTablasSQLite();
      console.log("✅ Base de datos SQLite inicializada (Capacitor 8)");
      storageMode = "sqlite";
      return true;
    } catch (error) {
      console.error("❌ Error inicializando SQLite (Capacitor 8):", error);
      try { await plugin.closeConnection({ database: "posmovil" }); } catch (_) {}
      activarModoLocal("Falló init SQLite Capacitor 8: " + (error.message || error));
      return true;
    }
  }

  // ============================================
  // 2) Fallback: window.SQLite (legacy Capacitor <8 / Cordova)
  // ============================================
  var sqliteGlobal = window.SQLite;
  if (sqliteGlobal && typeof sqliteGlobal.createDatabase === "function") {
    try {
      db = await sqliteGlobal.createDatabase({ database: "posmovil.db" });
      await crearTablasSQLite();
      console.log("✅ Base de datos SQLite inicializada (legacy)");
      storageMode = "sqlite";
      return true;
    } catch (error) {
      console.error("❌ Error inicializando SQLite (legacy):", error);
      activarModoLocal("Falló init SQLite legacy");
      return true;
    }
  }

  // ============================================
  // 3) No hay plugin SQLite → localStorage
  // ============================================
  activarModoLocal("Plugin SQLite no expuesto");
  return true;
}

// ============================================
// OBTENER ÚLTIMO CÓDIGO DE PRODUCTO
// ============================================
async function getUltimoCodigoProducto() {
  if (storageMode !== "sqlite" || !db) {
    var productosLocal = leerLocal(LS_KEYS.productos, []);
    var ultimoCodigo = 0;
    for (var i = 0; i < productosLocal.length; i++) {
      var codigo = productosLocal[i].codigo || "";
      if (codigo.indexOf("Pr_") === 0) {
        var numero = parseInt(codigo.substring(3)) || 0;
        if (numero > ultimoCodigo) ultimoCodigo = numero;
      }
    }
    return "Pr_" + String(ultimoCodigo + 1).padStart(5, "0");
  }

  try {
    var result = await db.execute(
      "SELECT codigo FROM productos WHERE codigo LIKE 'Pr_%' ORDER BY LENGTH(codigo) DESC, codigo DESC LIMIT 1"
    );
    
    if (result.values && result.values.length > 0) {
      var ultimoCodigo = result.values[0].codigo || "";
      if (ultimoCodigo.indexOf("Pr_") === 0) {
        var numero = parseInt(ultimoCodigo.substring(3)) || 0;
        return "Pr_" + String(numero + 1).padStart(5, "0");
      }
    }
    return "Pr_00001";
  } catch (error) {
    console.error("❌ Error obteniendo último código:", error);
    return "Pr_00001";
  }
}

// ============================================
// GUARDAR NUEVO PRODUCTO EN BD
// ============================================
async function guardarNuevoProducto(producto) {
  if (!producto || !producto.codigo) {
    return false;
  }

  if (storageMode !== "sqlite" || !db) {
    var productosLocal = leerLocal(LS_KEYS.productos, []);
    productosLocal.push({
      codigo: producto.codigo,
      nombre: producto.nombre,
      precio: Number(producto.precioVenta || 0),
      disponibilidad: Number(producto.cantidad || 0),
    });
    return guardarLocal(LS_KEYS.productos, productosLocal);
  }

  try {
    await db.execute(
      "INSERT OR REPLACE INTO productos (codigo, nombre, precio, disponibilidad) VALUES (?, ?, ?, ?)",
      [producto.codigo, producto.nombre, Number(producto.precioVenta || 0), Number(producto.cantidad || 0)]
    );
    console.log("✅ Nuevo producto guardado:", producto.codigo);
    return true;
  } catch (error) {
    console.error("❌ Error guardando nuevo producto:", error);
    return false;
  }
}

// ============================================
// GUARDAR ENTRADA COMPLETA DE PRODUCTO (productos + entrada_productos_pending)
// ============================================
async function guardarEntradaProductoCompleto(producto) {
  if (!producto || !producto.codigo) {
    return false;
  }

  // 1. Guardar en tabla productos (disponible para venta)
  var guardadoEnProductos = await guardarNuevoProducto(producto);
  if (!guardadoEnProductos) {
    console.error("❌ Error guardando en tabla productos");
    return false;
  }

  // 2. Guardar en tabla entrada_productos_pending (para sync al servidor)
  var fechaHora = fechaLocalISO();
  console.log("🔍 [DEBUG] guardarEntradaProductoCompleto - Datos:", JSON.stringify({
    codigo: producto.codigo,
    nombre: producto.nombre,
    cantidad: producto.cantidad,
    precioVenta: producto.precioVenta,
    precioCosto: producto.precioCosto
  }));

  if (storageMode !== "sqlite" || !db) {
    // Modo localStorage
    var entradas = leerLocal(LS_KEYS.entrada_productos, []);
    entradas.push({
      id: Date.now() + Math.floor(Math.random() * 1000),
      codigo: producto.codigo,
      nombre: producto.nombre,
      cantidad: Number(producto.cantidad || 0),
      precio_venta: Number(producto.precioVenta || 0),
      precio_costo: Number(producto.precioCosto || 0),
      fecha_hora: fechaHora,
      synced: 0,
    });
    console.log("🔍 [DEBUG] Guardado en localStorage, total entradas:", entradas.length);
    return guardarLocal(LS_KEYS.entrada_productos, entradas);
  }

  // Modo SQLite
  try {
    console.log("🔍 [DEBUG] Guardando en SQLite entrada_productos_pending...");
    await db.execute(
      "INSERT INTO entrada_productos_pending (codigo, nombre, cantidad, precio_venta, precio_costo, fecha_hora, synced) VALUES (?, ?, ?, ?, ?, ?, ?)",
      [
        producto.codigo,
        producto.nombre,
        Number(producto.cantidad || 0),
        Number(producto.precioVenta || 0),
        Number(producto.precioCosto || 0),
        fechaHora,
        0
      ]
    );
    console.log("✅ Entrada de producto guardada para sync:", producto.codigo);
    return true;
  } catch (error) {
    console.error("❌ Error guardando entrada_producto_pending:", error.message);
    return false;
  }
}

// ============================================
// OBTENER PRODUCTOS DESDE SQLite
// ============================================
async function getProductosLocal() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT codigo, nombre, precio, disponibilidad FROM productos ORDER BY nombre"
      );
      return result.values || [];
    },
    () => leerLocal(LS_KEYS.productos, [])
  );
}

// ============================================
// GUARDAR PRODUCTOS EN SQLite (para cache offline)
// ============================================
async function syncProductosLocal(productos) {
  var normalizados = (productos || []).map(normalizarProducto);

  return conSQLite(
    async (db) => {
      await db.execute("DELETE FROM productos");

      for (var i = 0; i < normalizados.length; i++) {
        var p = normalizados[i];
        await db.execute(
          "INSERT OR REPLACE INTO productos (codigo, nombre, precio, disponibilidad) VALUES (?, ?, ?, ?)",
          [p.codigo, p.nombre, p.precio || 0, p.disponibilidad || 0]
        );
      }

      console.log("✅ " + normalizados.length + " productos guardados en SQLite");
      return true;
    },
    () => guardarLocal(LS_KEYS.productos, normalizados)
  );
}

// ============================================
// RECETAS - Sincronizar y obtener desde SQLite/localStorage
// ============================================
async function syncRecetasLocal(recetas) {
  // Normalizar recetas: los encabezados del Excel pueden venir como
  // "cantlote" o "cant_lote", "precioventa" o "precio_venta", etc.
  var normalizadas = (recetas || []).map(function(r) {
    return {
      nombre: r.nombre || "",
      cant_lote: Number(r.cant_lote ?? r.cantlote ?? 0),
      precio_venta: Number(r.precio_venta ?? r.precioventa ?? 0),
      precio_costo: Number(r.precio_costo ?? r.preciocosto ?? 0),
    };
  });

  return conSQLite(
    async (db) => {
      await db.execute("DELETE FROM recetas");

      for (var i = 0; i < normalizadas.length; i++) {
        var r = normalizadas[i];
        await db.execute(
          "INSERT OR REPLACE INTO recetas (nombre, cant_lote, precio_venta, precio_costo) VALUES (?, ?, ?, ?)",
          [r.nombre, r.cant_lote, r.precio_venta, r.precio_costo]
        );
      }

      console.log("✅ " + normalizadas.length + " recetas guardadas en SQLite");
      return true;
    },
    () => guardarLocal(LS_KEYS.recetas, normalizadas)
  );
}

async function getRecetasLocal() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT nombre, cant_lote, precio_venta, precio_costo FROM recetas ORDER BY nombre"
      );
      return result.values || [];
    },
    () => leerLocal(LS_KEYS.recetas, [])
  );
}

// ============================================
// INGREDIENTES - Sincronizar y obtener desde SQLite/localStorage
// ============================================
async function syncIngredientesLocal(ingredientes) {
  var normalizadas = (ingredientes || []).map(function(i) {
    return {
      ingrediente: i.ingrediente || "",
      cantidad: Number(i.cantidad ?? 0),
      unidad: i.unidad || "",
      receta: i.receta || "",
    };
  });

  return conSQLite(
    async (db) => {
      await db.execute("DELETE FROM ingredientes");

      for (var i = 0; i < normalizadas.length; i++) {
        var ing = normalizadas[i];
        await db.execute(
          "INSERT INTO ingredientes (ingrediente, cantidad, unidad, receta) VALUES (?, ?, ?, ?)",
          [ing.ingrediente, ing.cantidad, ing.unidad, ing.receta]
        );
      }

      console.log("✅ " + normalizadas.length + " ingredientes guardados en SQLite");
      return true;
    },
    () => guardarLocal(LS_KEYS.ingredientes, normalizadas)
  );
}

async function getIngredientesLocal() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT ingrediente, cantidad, unidad, receta FROM ingredientes ORDER BY receta, ingrediente"
      );
      return result.values || [];
    },
    () => leerLocal(LS_KEYS.ingredientes, [])
  );
}

// ============================================
// GUARDAR VENTA EN SQLite (modo offline)
// ============================================
async function guardarVentaOffline(venta) {
  if (!venta || !Array.isArray(venta.productos) || venta.productos.length === 0) {
    return false;
  }

  return conSQLite(
    async (db) => {
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
    },
    function() {
      var actuales = leerLocal(LS_KEYS.ventas, []);
      var lineas = construirLineasVenta(venta, 0);
      return guardarLocal(LS_KEYS.ventas, actuales.concat(lineas));
    }
  );
}

// ============================================
// OBTENER VENTAS PENDIENTES DE SYNC
// ============================================
async function getVentasPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM ventas_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.ventas, []).filter(function(v) {
        return Number(v.synced || 0) === 0;
      });
    }
  );
}

// ============================================
// OBTENER RESUMEN OFFLINE DESDE SQLite (ventas por rango de fechas)
// ============================================
async function getResumenOffline(desde, hasta) {
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

    var fechaDesde = desde || fechaLocalHoyISO();
    var fechaHasta = hasta || fechaDesde;
    ventas = ventas.filter(function(v) {
      var fechaVenta = extraerFechaISO(v.fecha_hora || v.fechaHora);
      return fechaVenta >= fechaDesde && fechaVenta <= fechaHasta;
    });

    if (ventas.length === 0) {
      // Sin ventas en el rango, leer gastos del rango
      var gastos = await leerGastosDelDia(fechaDesde, fechaHasta);
      var totalGastos = gastos.reduce(function(acc, g) { return acc + Number(g.monto || 0); }, 0);
      return {
        totalIngresado: 0,
        efectivo: 0,
        transferencia: 0,
        productosVendidos: [],
        mermas: [],
        totalGastos: Math.round(totalGastos * 1000) / 1000,
        gastos: gastos
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

    // ---- GASTOS DEL RANGO ----
    var gastos = await leerGastosDelDia(fechaDesde, fechaHasta);
    var totalGastos = gastos.reduce(function(acc, g) { return acc + Number(g.monto || 0); }, 0);

    // ---- MERMAS DEL RANGO (agrupadas por producto) ----
    var mermas = [];
    try {
      var mermasRaw = [];
      if (storageMode === "sqlite" && db) {
        var mResult = await db.execute(
          "SELECT * FROM mermas_pending ORDER BY id"
        );
        mermasRaw = mResult.values || [];
      } else {
        mermasRaw = leerLocal(LS_KEYS.mermas, []);
      }
      // Filtrar por rango de fechas
      mermasRaw = mermasRaw.filter(function(m) {
        var fechaMerma = extraerFechaISO(m.fecha_hora);
        return fechaMerma >= fechaDesde && fechaMerma <= fechaHasta;
      });
      // Agrupar por producto
      var mermasAgrupadas = {};
      for (var mi = 0; mi < mermasRaw.length; mi++) {
        var m = mermasRaw[mi];
        var key = m.codigo_producto;
        if (!mermasAgrupadas[key]) {
          mermasAgrupadas[key] = {
            codigo: m.codigo_producto,
            nombre: m.nombre,
            cantidad: 0
          };
        }
        mermasAgrupadas[key].cantidad += Number(m.cantidad || 0);
      }
      // Convertir a array y ordenar por cantidad descendente
      mermas = Object.keys(mermasAgrupadas).map(function(k) { return mermasAgrupadas[k]; });
      mermas.sort(function(a, b) { return b.cantidad - a.cantidad; });
    } catch (mErr) {
      console.error("❌ Error leyendo mermas para resumen:", mErr);
    }

    return {
      totalIngresado: Math.round(totalIngresado * 1000) / 1000,
      efectivo: Math.round(totalEfectivo * 1000) / 1000,
      transferencia: Math.round(totalTransferencia * 1000) / 1000,
      productosVendidos: productosVendidos,
      mermas: mermas,
      totalGastos: Math.round(totalGastos * 1000) / 1000,
      gastos: gastos
    };
  } catch (error) {
    console.error("❌ Error en getResumenOffline:", error);
    return {
      totalIngresado: 0,
      efectivo: 0,
      transferencia: 0,
      productosVendidos: [],
      mermas: []
    };
  }
}

// ============================================
// CONTAR VENTAS PENDIENTES
// ============================================
async function contarVentasPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(DISTINCT factura_id) as count FROM ventas_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    () => contarFacturasPendientesLocal(leerLocal(LS_KEYS.ventas, []))
  );
}

// Marcar ventas como sincronizadas (helper para batch completo)
async function marcarVentasSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE ventas_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todas = leerLocal(LS_KEYS.ventas, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todas.length; x++) {
      if (idsMap[String(todas[x].id)]) {
        todas[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.ventas, todas);
  }
}

// ============================================
// FUNCIONES PARA MERMAS OFFLINE
// ============================================

// Guardar merma en SQLite/localStorage
async function guardarMermaOffline(merma) {
  if (!merma || !Array.isArray(merma.productos) || merma.productos.length === 0) {
    return false;
  }

  function construirLineasMerma() {
    return (merma.productos || []).map(function(item, idx) {
      return {
        id: Date.now() + idx + Math.floor(Math.random() * 1000),
        codigo_producto: item.codigo,
        nombre: item.nombre,
        cantidad: Number(item.cantidad || 0),
        fecha_hora: merma.fechaHora,
        synced: 0,
      };
    });
  }

  return conSQLite(
    async (db) => {
      for (var j = 0; j < merma.productos.length; j++) {
        var prod = merma.productos[j];
        await db.execute(
          "INSERT INTO mermas_pending (codigo_producto, nombre, cantidad, fecha_hora, synced) VALUES (?, ?, ?, ?, 0)",
          [prod.codigo, prod.nombre, prod.cantidad, merma.fechaHora]
        );
      }
      console.log("✅ Merma guardada offline (" + merma.productos.length + " productos)");
      return true;
    },
    function() {
      var actuales = leerLocal(LS_KEYS.mermas, []);
      return guardarLocal(LS_KEYS.mermas, actuales.concat(construirLineasMerma()));
    }
  );
}

// Obtener mermas pendientes de sync
async function getMermasPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM mermas_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.mermas, []).filter(function(m) {
        return Number(m.synced || 0) === 0;
      });
    }
  );
}

// Contar mermas pendientes
async function contarMermasPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(*) as count FROM mermas_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    function() {
      return leerLocal(LS_KEYS.mermas, []).filter(function(m) {
        return Number(m.synced || 0) === 0;
      }).length;
    }
  );
}

// Marcar mermas como sincronizadas
async function marcarMermasSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE mermas_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todas = leerLocal(LS_KEYS.mermas, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todas.length; x++) {
      if (idsMap[String(todas[x].id)]) {
        todas[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.mermas, todas);
  }
}

// ============================================
// FUNCIONES PARA ENTRADA DE PRODUCTOS (sync al servidor)
// ============================================
// FUNCIONES PARA ENTRADA DE PRODUCTOS (sync al servidor)
// ============================================

// Obtener entrada de productos pendientes de sync
async function getEntradaProductosPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM entrada_productos_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.entrada_productos, []).filter(function(p) {
        return Number(p.synced || 0) === 0;
      });
    }
  );
}

// Contar entrada de productos pendientes
async function contarEntradaProductosPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(*) as count FROM entrada_productos_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    function() {
      return leerLocal(LS_KEYS.entrada_productos, []).filter(function(p) {
        return Number(p.synced || 0) === 0;
      }).length;
    }
  );
}

// Contar abastecimientos pendientes
async function contarAbastecerPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(*) as count FROM abastecer_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    function() {
      return leerLocal(LS_KEYS.abastecer, []).filter(function(a) {
        return Number(a.synced || 0) === 0;
      }).length;
    }
  );
}

// Obtener abastecimientos pendientes de sync
async function getAbastecerPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM abastecer_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.abastecer, []).filter(function(a) {
        return Number(a.synced || 0) === 0;
      });
    }
  );
}

// Marcar abastecimientos como sincronizados
async function marcarAbastecerSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE abastecer_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todos = leerLocal(LS_KEYS.abastecer, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todos.length; x++) {
      if (idsMap[String(todos[x].id)]) {
        todos[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.abastecer, todos);
  }
}

// Marcar entrada de productos como sincronizadas
async function marcarEntradaProductosSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE entrada_productos_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todas = leerLocal(LS_KEYS.entrada_productos, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todas.length; x++) {
      if (idsMap[String(todas[x].id)]) {
        todas[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.entrada_productos, todas);
  }
}

// ============================================
// FUNCIONES PARA ABASTECER (reabastecer productos existentes)
// ============================================

// Actualizar stock de un producto (sumar cantidad)
async function actualizarStockProducto(codigo, cantidadSumar) {
  if (storageMode !== "sqlite" || !db) {
    // Modo localStorage
    var productosLocal = leerLocal(LS_KEYS.productos, []);
    var encontrado = false;
    for (var i = 0; i < productosLocal.length; i++) {
      if (productosLocal[i].codigo === codigo) {
        productosLocal[i].disponibilidad = Number(productosLocal[i].disponibilidad || 0) + Number(cantidadSumar);
        encontrado = true;
        break;
      }
    }
    if (encontrado) {
      guardarLocal(LS_KEYS.productos, productosLocal);
      return true;
    }
    return false;
  }

  try {
    await db.execute(
      "UPDATE productos SET disponibilidad = disponibilidad + ? WHERE codigo = ?",
      [Number(cantidadSumar), codigo]
    );
    console.log("✅ Stock actualizado para " + codigo + " (+" + cantidadSumar + ")");
    return true;
  } catch (error) {
    console.error("❌ Error actualizando stock:", error);
    return false;
  }
}

// Guardar abastecimiento en SQLite/localStorage
async function guardarAbastecerOffline(abastecer) {
  if (!abastecer || !abastecer.codigo || !abastecer.cantidad) {
    return false;
  }

  // 1. Sumar cantidad al stock en tabla productos
  var stockActualizado = await actualizarStockProducto(abastecer.codigo, abastecer.cantidad);
  if (!stockActualizado) {
    console.error("❌ Error actualizando stock del producto");
    return false;
  }

  // 2. Guardar registro en abastecer_pending
  var fechaHora = fechaLocalISO();

  if (storageMode !== "sqlite" || !db) {
    // Modo localStorage
    var abastecimientos = leerLocal(LS_KEYS.abastecer, []);
    abastecimientos.push({
      id: Date.now() + Math.floor(Math.random() * 1000),
      codigo_producto: abastecer.codigo,
      nombre: abastecer.nombre,
      cantidad: Number(abastecer.cantidad),
      fecha_hora: fechaHora,
      synced: 0,
    });
    return guardarLocal(LS_KEYS.abastecer, abastecimientos);
  }

  // Modo SQLite
  try {
    await db.execute(
      "INSERT INTO abastecer_pending (codigo_producto, nombre, cantidad, fecha_hora, synced) VALUES (?, ?, ?, ?, ?)",
      [abastecer.codigo, abastecer.nombre, Number(abastecer.cantidad), fechaHora, 0]
    );
    console.log("✅ Abastecimiento guardado:", abastecer.codigo, "+" + abastecer.cantidad);
    return true;
  } catch (error) {
    console.error("❌ Error guardando abastecimiento:", error);
    return false;
  }
}

// ============================================
// FUNCIONES PARA GASTOS OFFLINE
// ============================================

// Guardar gasto en SQLite/localStorage
async function guardarGastoOffline(gasto) {
  if (!gasto || !gasto.descripcion || !gasto.monto) {
    return false;
  }

  var fecha = gasto.fecha || fechaLocalISO().split("T")[0];
  var monto = Number(gasto.monto) || 0;

  function construirGasto() {
    return {
      id: Date.now() + Math.floor(Math.random() * 1000),
      fecha: fecha,
      descripcion: gasto.descripcion,
      monto: monto,
      synced: 0,
    };
  }

  return conSQLite(
    async (db) => {
      await db.execute(
        "INSERT INTO gastos_pending (fecha, descripcion, monto, synced) VALUES (?, ?, ?, 0)",
        [fecha, gasto.descripcion, monto]
      );
      console.log("✅ Gasto guardado offline:", gasto.descripcion, "$" + monto);
      return true;
    },
    function() {
      var actuales = leerLocal(LS_KEYS.gastos, []);
      actuales.push(construirGasto());
      return guardarLocal(LS_KEYS.gastos, actuales);
    }
  );
}

// Obtener gastos pendientes de sync
async function getGastosPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM gastos_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.gastos, []).filter(function(g) {
        return Number(g.synced || 0) === 0;
      });
    }
  );
}

// Contar gastos pendientes
async function contarGastosPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(*) as count FROM gastos_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    function() {
      return leerLocal(LS_KEYS.gastos, []).filter(function(g) {
        return Number(g.synced || 0) === 0;
      }).length;
    }
  );
}

// Marcar gastos como sincronizados
async function marcarGastosSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE gastos_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todos = leerLocal(LS_KEYS.gastos, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todos.length; x++) {
      if (idsMap[String(todos[x].id)]) {
        todos[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.gastos, todos);
  }
}


// ============================================
// FUNCIONES PARA ELABORACIONES (registro de producción local)
// ============================================

// Guardar elaboración en SQLite/localStorage
async function guardarElaboracionOffline(elaboracion) {
  if (!elaboracion || !elaboracion.nombre_receta || !elaboracion.lotes) {
    return false;
  }

  var fechaHora = fechaLocalISO();

  function construirElaboracion() {
    return {
      id: Date.now() + Math.floor(Math.random() * 1000),
      nombre_receta: elaboracion.nombre_receta,
      lotes: Number(elaboracion.lotes || 0),
      cantidad_producida: Number(elaboracion.cantidad_producida || 0),
      precio_venta: Number(elaboracion.precio_venta || 0),
      fecha_hora: fechaHora,
      synced: 0,
    };
  }

  return conSQLite(
    async (db) => {
      await db.execute(
        "INSERT INTO elaboraciones_pending (nombre_receta, lotes, cantidad_producida, precio_venta, fecha_hora, synced) VALUES (?, ?, ?, ?, ?, 0)",
        [
          elaboracion.nombre_receta,
          Number(elaboracion.lotes || 0),
          Number(elaboracion.cantidad_producida || 0),
          Number(elaboracion.precio_venta || 0),
          fechaHora,
        ]
      );
      console.log("✅ Elaboración guardada:", elaboracion.nombre_receta, "x" + elaboracion.lotes + " lote(s)");
      return true;
    },
    function() {
      var actuales = leerLocal(LS_KEYS.elaboraciones, []);
      actuales.push(construirElaboracion());
      return guardarLocal(LS_KEYS.elaboraciones, actuales);
    }
  );
}

// Obtener elaboraciones pendientes de sync
async function getElaboracionesPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT * FROM elaboraciones_pending WHERE synced = 0 ORDER BY id"
      );
      return result.values || [];
    },
    function() {
      return leerLocal(LS_KEYS.elaboraciones, []).filter(function(e) {
        return Number(e.synced || 0) === 0;
      });
    }
  );
}

// Contar elaboraciones pendientes
async function contarElaboracionesPendientes() {
  return conSQLite(
    async (db) => {
      var result = await db.execute(
        "SELECT COUNT(*) as count FROM elaboraciones_pending WHERE synced = 0"
      );
      return result.values ? result.values[0].count : 0;
    },
    function() {
      return leerLocal(LS_KEYS.elaboraciones, []).filter(function(e) {
        return Number(e.synced || 0) === 0;
      }).length;
    }
  );
}

// Marcar elaboraciones como sincronizadas
async function marcarElaboracionesSynced(ids) {
  if (storageMode === "sqlite" && db) {
    for (var i = 0; i < ids.length; i++) {
      await db.execute(
        "UPDATE elaboraciones_pending SET synced = 1 WHERE id = ?",
        [ids[i]]
      );
    }
  } else {
    var todos = leerLocal(LS_KEYS.elaboraciones, []);
    var idsMap = {};
    for (var j = 0; j < ids.length; j++) {
      idsMap[String(ids[j])] = true;
    }
    for (var x = 0; x < todos.length; x++) {
      if (idsMap[String(todos[x].id)]) {
        todos[x].synced = 1;
      }
    }
    guardarLocal(LS_KEYS.elaboraciones, todos);
  }
}

// ============================================
// SINCRONIZAR TODO EN UN SOLO REQUEST BATCH
// ============================================
// Reemplaza 5 requests separados (entrada-productos, abastecimientos,
// ventas, mermas, gastos) por UN solo request.
// El servidor procesa en orden, remapea códigos/facturas si hay
// conflicto, y la APK solo necesita recargar productos al final.
// ============================================
async function sincronizarCompleto(serverUrl) {
  logSyncDebugAPK("📦 [sync-completo] Iniciando batch...");

  var url = serverUrl || window.SERVER_URL;
  if (!url) {
    logSyncDebugAPK("❌ [sync-completo] No hay URL de servidor", "error");
    return { success: false, error: "No hay servidor configurado" };
  }

  // 1. Recolectar datos pendientes de cada categoría
  var pendientesProductos  = await getEntradaProductosPendientes();
  var pendientesAbast      = await getAbastecerPendientes();
  var pendientesVentas     = await getVentasPendientes();
  var pendientesMermas     = await getMermasPendientes();
  var pendientesGastos     = await getGastosPendientes();
  var pendientesElaborac   = await getElaboracionesPendientes();

  var total = pendientesProductos.length + pendientesAbast.length +
              pendientesVentas.length + pendientesMermas.length +
              pendientesGastos.length + pendientesElaborac.length;

  if (total === 0) {
    logSyncDebugAPK("📦 [sync-completo] Sin datos pendientes — consultando solo referencia");
    // Igual llamamos al servidor para obtener ultimaFactura
    try {
      var r = await fetch(url + "/api/sync/completo", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-session-token": window.SESSION_TOKEN || "",
        },
        body: JSON.stringify({}),
      });
      if (r.ok) {
        var resultVacio = await r.json();
        if (resultVacio.ultimaFactura) {
          var uf = resultVacio.ultimaFactura.toString();
          guardarLocal('posmovil_ultima_factura_referencia', uf);
          if (uf.length >= 10) {
            guardarLocal('posmovil_ultimo_prefijo', uf.substring(0, 5));
            guardarLocal('posmovil_ultimo_sufijo', parseInt(uf.substring(5, 10)) || 0);
            var hoy = new Date().getFullYear() + '-' +
                      String(new Date().getMonth() + 1).padStart(2, '0') + '-' +
                      String(new Date().getDate()).padStart(2, '0');
            guardarLocal('posmovil_fecha_referencia', hoy);
          }
        }
        return { success: true, sincronizados: {}, remap: { codigos: {}, facturas: {} }, ultimaFactura: resultVacio.ultimaFactura || null };
      }
    } catch (e) {
      console.error("📦 [sync-completo] Error consultando servidor vacío:", e.message);
    }
    return { success: true, sincronizados: {}, remap: { codigos: {}, facturas: {} }, ultimaFactura: null };
  }

  // Si solo hay elaboraciones, podemos continuar igual (se envían al servidor)

  logSyncDebugAPK("📦 [sync-completo] Enviando " + total + " registros (" +
    pendientesProductos.length + " prod, " +
    pendientesAbast.length + " abast, " +
    pendientesVentas.length + " ventas, " +
    pendientesMermas.length + " mermas, " +
    pendientesGastos.length + " gastos, " +
    pendientesElaborac.length + " elaborac)");

  // 2. Preparar payload (mismos formatos que antes)
  var payload = {};

  if (pendientesProductos.length) {
    payload.productos = pendientesProductos.map(function(p) {
      return {
        codigo: p.codigo,
        nombre: p.nombre,
        cantidad: p.cantidad,
        precio_venta: p.precio_venta,
        precio_costo: p.precio_costo,
        fecha_hora: p.fecha_hora,
      };
    });
  }

  if (pendientesAbast.length) {
    payload.abastecimientos = pendientesAbast.map(function(a) {
      return {
        codigo: a.codigo_producto,
        nombre: a.nombre,
        cantidad: Number(a.cantidad || 0),
        fechaHora: a.fecha_hora,
      };
    });
  }

  if (pendientesVentas.length) {
    payload.ventas = pendientesVentas.map(function(v) {
      return {
        facturaId: v.factura_id,
        fechaHora: v.fecha_hora,
        codigoProducto: v.codigo_producto,
        nombre: v.nombre,
        cantidad: v.cantidad,
        precio: v.precio,
        subtotal: v.subtotal,
        efectivo: v.efectivo || v.efectividad || 0,
        transferencia: v.transferencia || 0,
      };
    });
  }

  if (pendientesMermas.length) {
    payload.mermas = pendientesMermas.map(function(m) {
      return {
        codigo: m.codigo_producto,
        nombre: m.nombre,
        cantidad: m.cantidad,
        fechaHora: m.fecha_hora,
      };
    });
  }

  if (pendientesGastos.length) {
    payload.gastos = pendientesGastos.map(function(g) {
      return {
        fecha: g.fecha,
        descripcion: g.descripcion,
        monto: Number(g.monto || 0),
      };
    });
  }

  if (pendientesElaborac.length) {
    payload.elaboraciones = pendientesElaborac.map(function(e) {
      return {
        nombre_receta: e.nombre_receta,
        lotes: Number(e.lotes || 0),
        cantidad_producida: Number(e.cantidad_producida || 0),
        precio_venta: Number(e.precio_venta || 0),
        fecha_hora: e.fecha_hora,
      };
    });
  }

  // 3. Enviar request
  try {
    var response = await fetch(url + "/api/sync/completo", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-session-token": window.SESSION_TOKEN || "",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      var errorText = await response.text();
      throw new Error("HTTP " + response.status + " - " + errorText);
    }

    var result = await response.json();

    if (!result.success) {
      throw new Error(result.mensaje || result.error || "Error desconocido del servidor");
    }

    // 4. Marcar TODOS los registros como sincronizados
    var idsProductos   = pendientesProductos.map(function(p) { return p.id; });
    var idsAbast       = pendientesAbast.map(function(a) { return a.id; });
    var idsVentas      = pendientesVentas.map(function(v) { return v.id; });
    var idsMermas      = pendientesMermas.map(function(m) { return m.id; });
    var idsGastos      = pendientesGastos.map(function(g) { return g.id; });
    var idsElaborac    = pendientesElaborac.map(function(e) { return e.id; });

    if (idsProductos.length) await marcarEntradaProductosSynced(idsProductos);
    if (idsAbast.length)     await marcarAbastecerSynced(idsAbast);
    if (idsVentas.length)    await marcarVentasSynced(idsVentas);
    if (idsMermas.length)    await marcarMermasSynced(idsMermas);
    if (idsGastos.length)    await marcarGastosSynced(idsGastos);
    if (idsElaborac.length)  await marcarElaboracionesSynced(idsElaborac);

    // 5. Guardar referencia de última factura si el servidor devolvió remap
    if (result.remap && result.remap.facturas) {
      var facturasArray = Object.values(result.remap.facturas);
      if (facturasArray.length > 0) {
        var ultimaFacturaRemap = facturasArray[facturasArray.length - 1];
        if (!guardarLocal('posmovil_ultima_factura_referencia', ultimaFacturaRemap)) {
          console.error("🚨 [sync-completo] No se pudo guardar referencia factura — riesgo de ID duplicado");
        }
        var refStr = ultimaFacturaRemap.toString();
        if (refStr.length >= 10) {
          guardarLocal('posmovil_ultimo_prefijo', refStr.substring(0, 5));
          guardarLocal('posmovil_ultimo_sufijo', parseInt(refStr.substring(5, 10)) || 0);
          var hoyISO = new Date().getFullYear() + '-' +
                      String(new Date().getMonth() + 1).padStart(2, '0') + '-' +
                      String(new Date().getDate()).padStart(2, '0');
          guardarLocal('posmovil_fecha_referencia', hoyISO);
        }
      }
    }

    // 6. Guardar la última factura REAL del servidor (siempre que venga en la respuesta)
    //    Esto asegura que después de cada sync el cliente tenga la referencia correcta
    //    del último ID existente en datos.xlsx, incluso si no hubo conflictos de remap.
    if (result.ultimaFactura) {
      var ultimaFactura = result.ultimaFactura.toString();
      if (!guardarLocal('posmovil_ultima_factura_referencia', ultimaFactura)) {
        console.error("🚨 [sync-completo] No se pudo guardar referencia factura — riesgo de ID duplicado");
      }
      if (ultimaFactura.length >= 10) {
        guardarLocal('posmovil_ultimo_prefijo', ultimaFactura.substring(0, 5));
        guardarLocal('posmovil_ultimo_sufijo', parseInt(ultimaFactura.substring(5, 10)) || 0);
        var hoyISO = new Date().getFullYear() + '-' +
                    String(new Date().getMonth() + 1).padStart(2, '0') + '-' +
                    String(new Date().getDate()).padStart(2, '0');
        guardarLocal('posmovil_fecha_referencia', hoyISO);
      }
    }

    logSyncDebugAPK(
      "✅ [sync-completo] Batch completado: " +
      (result.sincronizados ? (
        Object.keys(result.sincronizados)
          .filter(function(k) { return result.sincronizados[k] > 0; })
          .map(function(k) { return result.sincronizados[k] + " " + k; })
          .join(", ")
      ) : "OK")
    );

    return result;
  } catch (error) {
    logSyncDebugAPK("❌ [sync-completo] Error: " + error.message, "error");
    return { success: false, error: error.message };
  }
}

// ============================================
// LEER GASTOS POR RANGO DE FECHAS DESDE BD LOCAL
// ============================================
async function leerGastosDelDia(desde, hasta) {
  var gastos = [];
  if (storageMode === "sqlite" && db) {
    try {
      var result = await db.execute("SELECT * FROM gastos_pending ORDER BY id");
      gastos = result.values || [];
    } catch (e) {
      gastos = leerLocal(LS_KEYS.gastos, []);
    }
  } else {
    gastos = leerLocal(LS_KEYS.gastos, []);
  }
  var fechaDesde = desde || fechaLocalHoyISO();
  var fechaHasta = hasta || fechaDesde;
  gastos = gastos.filter(function(g) {
    var f = (g.fecha || "").substring(0, 10);
    return f >= fechaDesde && f <= fechaHasta;
  });
  return gastos.map(function(g) {
    return {
      descripcion: g.descripcion || "",
      monto: Number(g.monto || 0),
    };
  });
}

// ============================================
// AUTO-DESCUBRIR SERVIDOR EN RED LOCAL
// ============================================
const SERVIDORES_LISTA_KEY = "servidores_lista";
const MAX_SERVIDORES_GUARDADOS = 10;

// ============================================
// LISTA PERSISTENTE DE SERVIDORES CONOCIDOS
// Orden: auto-detectados primero, manuales después
// Sin TTL — permanentes hasta que el usuario los borre
// ============================================

async function _leerListaServidores() {
  if (storageMode !== "sqlite" || !db) {
    var cfgLocal = leerLocal(LS_KEYS.config, {});
    return cfgLocal[SERVIDORES_LISTA_KEY] || [];
  }
  try {
    var result = await db.execute(
      "SELECT valor FROM config WHERE clave = ?",
      [SERVIDORES_LISTA_KEY]
    );
    if (result.values && result.values.length > 0) {
      try {
        return JSON.parse(result.values[0].valor);
      } catch (e) {
        return [];
      }
    }
    return [];
  } catch (e) {
    console.error("❌ [_leerListaServidores] Error:", e);
    return [];
  }
}

async function _escribirListaServidores(lista) {
  var json = JSON.stringify(lista);
  if (storageMode !== "sqlite" || !db) {
    var cfgLocal = leerLocal(LS_KEYS.config, {});
    cfgLocal[SERVIDORES_LISTA_KEY] = json;
    guardarLocal(LS_KEYS.config, cfgLocal);
    return;
  }
  try {
    await db.execute(
      "INSERT OR REPLACE INTO config (clave, valor, timestamp) VALUES (?, ?, ?)",
      [SERVIDORES_LISTA_KEY, json, Date.now()]
    );
  } catch (e) {
    console.error("❌ [_escribirListaServidores] Error:", e);
  }
}

/**
 * Devuelve la lista completa de servidores conocidos (ordenados por prioridad).
 */
async function obtenerListaServidores() {
  return await _leerListaServidores();
}

/**
 * Agrega un servidor a la lista persistente.
 * @param {string} url - URL del servidor (ej. "http://192.168.1.100:3000")
 * @param {string} tipo - "auto" (va al inicio) | "manual" (va al final)
 */
async function agregarServidorConocido(url, tipo) {
  var lista = await _leerListaServidores();

  // Quitar duplicado si ya existe
  var idx = lista.indexOf(url);
  if (idx !== -1) lista.splice(idx, 1);

  if (tipo === "auto") {
    lista.unshift(url); // al principio
  } else {
    lista.push(url); // al final
  }

  // Limitar tamaño
  if (lista.length > MAX_SERVIDORES_GUARDADOS) {
    lista = lista.slice(0, MAX_SERVIDORES_GUARDADOS);
  }

  await _escribirListaServidores(lista);
  console.log("✅ [agregarServidorConocido] Guardado (" + tipo + "):", url);
}

// ============================================
// PROBAR UN SERVIDOR (health check rápido)
// ============================================
async function probarServidor(url) {
  try {
    var res = await fetch(url + "/api/estado-publico", {
      method: "GET",
      signal: AbortSignal.timeout(2000)
    });
    return res.ok;
  } catch (e) {
    return false;
  }
}

// ============================================
// DESCUBRIR SERVIDOR (probando lista conocida + escaneo)
// ============================================
async function descubrirServidor(onProgreso) {
  // 1. Probar servidores conocidos en orden (auto-detectados primero)
  var lista = await _leerListaServidores();
  for (var i = 0; i < lista.length; i++) {
    if (cancelarEscaneo) break;
    var url = lista[i];
    if (onProgreso) onProgreso("Probando " + url + "...");
    console.log("🔍 [descubrirServidor] Probando servidor conocido:", url);
    var ok = await probarServidor(url);
    if (ok) {
      console.log("🎯 [descubrirServidor] Usando servidor conocido:", url);
      return url;
    }
  }

  // 2. Si ninguno respondió (y no se canceló), escanear la red
  if (!cancelarEscaneo) {
    var servidorEncontrado = await escanearRed(onProgreso);

    // 3. Si se encontró, guardar en la lista
    if (servidorEncontrado) {
      await agregarServidorConocido(servidorEncontrado, "auto");
      console.log("✅ [descubrirServidor] Servidor encontrado y guardado:", servidorEncontrado);
    }

    return servidorEncontrado;
  }

  return null;
}

// ============================================
// ESCANEAR RED LOCAL (soporta 192.168.x.x y 10.x.x.x)
// ============================================
// ESCANEAR UN RANGO ESPECÍFICO DE RED
// ============================================
async function escanearRango(baseIP, puerto, ipLocal) {
  console.log("🔍 [escanearRango] Escaneando: " + baseIP + ".x");

  var batchSize = 25;
  var timeoutMs = 300;

  for (var batchStart = 1; batchStart <= 254; batchStart += batchSize) {
    if (cancelarEscaneo || _servidorEncontrado) break;

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
        _servidorEncontrado = resultados[r].url;
        cancelarEscaneo = true;
        console.log("✅ [escanearRango] Servidor encontrado: " + _servidorEncontrado);
        return _servidorEncontrado;
      }
    }
  }

  return null;
}

// ============================================
async function escanearRed(onProgreso) {
  cancelarEscaneo = false;
  _servidorEncontrado = null;
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

  // Agregar siempre los rangos más comunes (wifi + tethering Android)
  var rangosComunes = [
    { base: "192.168.1", origen: "wifi (común)" },
    { base: "192.168.43", origen: "Android tethering" },
    { base: "192.168.0", origen: "wifi alternativo" },
    { base: "10.0.0", origen: "tethering/otro" },
    { base: "10.0.1", origen: "tethering/otro" },
    { base: "10.225.81", origen: "tethering (caso actual)" },
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

  // Escanear rangos en paralelo (hasta 3 a la vez)
  var concurrency = 3;
  var totalRangos = rangos.length;

  for (var k = 0; k < totalRangos; k += concurrency) {
    if (cancelarEscaneo) {
      console.log("🛑 [escanearRed] Escaneo cancelado por usuario");
      cancelarEscaneo = false;
      return null;
    }

    var batch = rangos.slice(k, k + concurrency);
    var batchNum = Math.floor(k / concurrency) + 1;
    var totalBatches = Math.ceil(totalRangos / concurrency);

    if (onProgreso) {
      var nombres = batch.map(function(r) { return r.base; }).join(", ");
      onProgreso("Escaneando... (grupo " + batchNum + "/" + totalBatches + ": " + nombres + ")");
    }

    await Promise.all(batch.map(function(r) {
      return escanearRango(r.base, puerto, ipLocal);
    }));

    if (_servidorEncontrado) return _servidorEncontrado;
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
// FUNCIONES PARA HISTORIAL DE VENTAS
// ============================================

// Obtener ventas agrupadas por factura_id para una fecha específica
// Devuelve un array de facturas con su estado de sincronización
async function getVentasAgrupadasPorFactura(fechaISO) {
  try {
    var ventas = [];
    var fechaFiltro = fechaISO || fechaLocalHoyISO();
    
    if (storageMode === "sqlite" && db) {
      var result = await db.execute(
        "SELECT * FROM ventas_pending WHERE date(fecha_hora) = ? ORDER BY factura_id, id",
        [fechaFiltro]
      );
      ventas = result.values || [];
    } else {
      ventas = leerLocal(LS_KEYS.ventas, []).filter(function(v) {
        return extraerFechaISO(v.fecha_hora || v.fechaHora) === fechaFiltro;
      });
    }
    
    if (ventas.length === 0) {
      return [];
    }
    
    // Agrupar por factura_id
    var facturasMap = {};
    
    for (var i = 0; i < ventas.length; i++) {
      var v = ventas[i];
      var facturaId = v.factura_id || ("sin_factura_" + i);
      
      if (!facturasMap[facturaId]) {
        facturasMap[facturaId] = {
          facturaId: facturaId,
          fechaHora: v.fecha_hora,
          productos: [],
          total: 0,
          efectivo: v.efectivo || v.efectividad || 0,
          transferencia: v.transferencia || 0,
          synced: Number(v.synced || 0),
          ids: []
        };
      }
      
      var factura = facturasMap[facturaId];
      factura.productos.push({
        codigo: v.codigo_producto,
        nombre: v.nombre,
        cantidad: Number(v.cantidad || 0),
        precio: Number(v.precio || 0),
        subtotal: Number(v.subtotal || 0)
      });
      factura.total += Number(v.subtotal || 0);
      
      // Si alguna línea no está sincronizada, la factura completa no lo está
      if (Number(v.synced || 0) === 0) {
        factura.synced = 0;
      }
      
      factura.ids.push(v.id);
    }
    
    // Convertir a array
    var facturas = [];
    var keys = Object.keys(facturasMap);
    for (var j = 0; j < keys.length; j++) {
      facturas.push(facturasMap[keys[j]]);
    }
    
    // Ordenar por factura_id descendente (más recientes primero)
    facturas.sort(function(a, b) {
      return String(b.facturaId).localeCompare(String(a.facturaId));
    });
    
    return facturas;
  } catch (error) {
    console.error("❌ Error en getVentasAgrupadasPorFactura:", error);
    return [];
  }
}

// Deshacer una venta (solo si no está sincronizada)
// Esto suma las cantidades de vuelta al stock y elimina la venta de ventas_pending
async function deshacerVenta(facturaId) {
  try {
    console.log("🔄 Deshaciendo venta:", facturaId);
    
    // 1. Obtener todas las líneas de la factura que NO estén sincronizadas
    var lineas = [];
    
    if (storageMode === "sqlite" && db) {
      var result = await db.execute(
        "SELECT * FROM ventas_pending WHERE factura_id = ? AND synced = 0",
        [facturaId]
      );
      lineas = result.values || [];
      
      if (lineas.length === 0) {
        console.warn("⚠️ No se encontró la venta o ya está sincronizada");
        return { success: false, error: "La venta no existe o ya fue sincronizada" };
      }
      
      // 2. Sumar cantidades de vuelta al stock
      for (var i = 0; i < lineas.length; i++) {
        var linea = lineas[i];
        var codigo = linea.codigo_producto;
        var cantidad = Number(linea.cantidad || 0);
        
        // Obtener stock actual
        var resultStock = await db.execute(
          "SELECT disponibilidad FROM productos WHERE codigo = ?",
          [codigo]
        );
        
        if (resultStock.values && resultStock.values.length > 0) {
          var stockActual = Number(resultStock.values[0].disponibilidad || 0);
          var nuevoStock = stockActual + cantidad;
          
          await db.execute(
            "UPDATE productos SET disponibilidad = ? WHERE codigo = ?",
            [nuevoStock, codigo]
          );
          
          console.log("📦 Stock restaurado:", codigo, "+" + cantidad, "→", nuevoStock);
        }
      }
      
      // 3. Eliminar las líneas de ventas_pending
      await db.execute(
        "DELETE FROM ventas_pending WHERE factura_id = ? AND synced = 0",
        [facturaId]
      );
      
      console.log("✅ Venta deshecha:", facturaId);
      return { success: true, message: "Venta deshecha correctamente" };
      
    } else {
      // Modo localStorage
      var ventas = leerLocal(LS_KEYS.ventas, []);
      var ventasFiltradas = [];
      var lineasDeshechas = [];
      
      for (var j = 0; j < ventas.length; j++) {
        var v = ventas[j];
        if (String(v.factura_id) === String(facturaId) && Number(v.synced || 0) === 0) {
          lineasDeshechas.push(v);
        } else {
          ventasFiltradas.push(v);
        }
      }
      
      if (lineasDeshechas.length === 0) {
        return { success: false, error: "La venta no existe o ya fue sincronizada" };
      }
      
      // Restaurar stock
      var productosLocal = leerLocal(LS_KEYS.productos, []);
      for (var k = 0; k < lineasDeshechas.length; k++) {
        var ld = lineasDeshechas[k];
        for (var p = 0; p < productosLocal.length; p++) {
          if (String(productosLocal[p].codigo) === String(ld.codigo_producto)) {
            productosLocal[p].disponibilidad = Number(productosLocal[p].disponibilidad || 0) + Number(ld.cantidad || 0);
            break;
          }
        }
      }
      
      guardarLocal(LS_KEYS.productos, productosLocal);
      guardarLocal(LS_KEYS.ventas, ventasFiltradas);
      
      console.log("✅ Venta deshecha (localStorage):", facturaId);
      return { success: true, message: "Venta deshecha correctamente" };
    }
  } catch (error) {
    console.error("❌ Error deshaciendo venta:", error);
    return { success: false, error: error.message };
  }
}

// ============================================
// SISTEMA DE ACTIVACIÓN DE DISPOSITIVOS
// ============================================
// Protege el acceso a la app: el dispositivo debe
// ser autorizado mediante una clave generada desde VBA.
// ============================================

// Constante maestra (DEBE SER IDÉNTICA EN JS Y VBA)
const MASTER_SALT = "GestionPlus2024!";

// Clave de localStorage para el estado de activación
const ACTIVATION_KEY = "gplus_device_activated";
const DEVICE_ID_KEY = "gplus_device_id_cache";

// ============================================
// GENERAR ID DE DISPOSITIVO (ofuscado)
// Combina info del hardware + salt persistente
// Devuelve un string hexadecimal
// ============================================
async function generateDeviceId() {
  // Verificar si ya tenemos uno cacheado (para no regenerar en cada carga)
  var cached = localStorage.getItem(DEVICE_ID_KEY);
  if (cached) return cached;

  var model = "unknown";
  var osVersion = "unknown";
  var platform = "unknown";
  var uuid = "unknown";

  // Si estamos en APK real, obtener info del dispositivo
  if (window.Capacitor && window.Capacitor.isNativePlatform()) {
    var devicePlugin = null;
    // Capacitor 8: plugins accesibles via Capacitor.Plugins
    if (window.Capacitor.Plugins && window.Capacitor.Plugins.Device) {
      devicePlugin = window.Capacitor.Plugins.Device;
    }

    if (devicePlugin && typeof devicePlugin.getInfo === 'function') {
      try {
        var info = await devicePlugin.getInfo();
        model = info.model || model;
        osVersion = info.osVersion || osVersion;
        platform = info.platform || platform;
      } catch (e) {
        console.warn("⚠️ [DeviceID] Error leyendo Device.getInfo:", e.message);
      }
    }

    if (devicePlugin && typeof devicePlugin.getId === 'function') {
      try {
        var idResult = await devicePlugin.getId();
        uuid = idResult.uuid || uuid;
      } catch (e) {
        console.warn("⚠️ [DeviceID] Error leyendo Device.getId:", e.message);
      }
    }
  }

  // Generar salt persistente (se genera UNA VEZ y queda para siempre)
  var salt = localStorage.getItem("gplus_device_salt");
  if (!salt) {
    var arr = [];
    for (var si = 0; si < 8; si++) {
      arr.push(Math.floor(Math.random() * 36).toString(36));
    }
    salt = arr.join("") + Date.now().toString(36);
    try {
      localStorage.setItem("gplus_device_salt", salt);
    } catch (e) {
      salt = "fallback_" + Date.now();
    }
  }

  // Combinar todo en un string crudo
  var raw = model + "|" + osVersion + "|" + platform + "|" + uuid + "|" + salt;

  // Ofuscar: transformación irreversible para el ID mostrado
  var obfuscated = "";
  for (var i = 0; i < raw.length; i++) {
    var code = raw.charCodeAt(i);
    code = ((code * 7) ^ 0x3B) & 0xFF;
    var hex = code.toString(16);
    if (hex.length < 2) hex = "0" + hex;
    obfuscated += hex;
  }

  var deviceId = obfuscated.toUpperCase();

  // Cachear para no regenerar
  try {
    localStorage.setItem(DEVICE_ID_KEY, deviceId);
  } catch (e) {}

  console.log("🔐 [DeviceID] Generado:", deviceId);
  return deviceId;
}

// ============================================
// COMPUTAR CLAVE DE ACTIVACIÓN
// ALGORITMO IDÉNTICO AL DE VBA
//
// ESTRATEGIA: acumular TODOS los caracteres transformados
// en un buffer fijo de 8 bytes con mezcla cruzada.
// Así, un cambio en CUALQUIER parte del input (no solo al inicio)
// modifica TODOS los bytes del buffer final.
// ============================================
function computeActivationKey(deviceId) {
  // 1. Combinar con salt maestra
  var combined = deviceId + MASTER_SALT;

  // 2. Buffer fijo de 8 bytes (→ 16 hex chars → XXXX-XXXX-XXXX-XXXX)
  var buf = [0, 0, 0, 0, 0, 0, 0, 0];

  // 3. Procesar CADA carácter, mezclando en múltiples posiciones del buffer
  for (var i = 0; i < combined.length; i++) {
    var code = combined.charCodeAt(i);
    // XOR con valor dependiente de posición
    code = (code ^ (0x2A + i)) & 0xFF;
    // Rotación derecha 2 bits
    code = ((code >> 2) | ((code & 3) << 6)) & 0xFF;

    // Mezcla cruzada: cada byte afecta 3 posiciones del buffer
    buf[i % 8] = (buf[i % 8] ^ code) & 0xFF;
    buf[(i + 3) % 8] = (buf[(i + 3) % 8] + code) & 0xFF;
    buf[(i + 7) % 8] = (buf[(i + 7) % 8] ^ ((code << 1) & 0xFF)) & 0xFF;
  }

  // 4. Mezcla final: cada byte se mezcla con su vecino
  //    para que cambios mínimos tengan efecto de avalancha
  for (var i = 0; i < 8; i++) {
    buf[i] = (buf[i] ^ buf[(i + 1) % 8]) & 0xFF;
  }

  // 5. Convertir buffer a hex y formatear
  var hexResult = "";
  for (var i = 0; i < buf.length; i++) {
    var h = buf[i].toString(16);
    if (h.length < 2) h = "0" + h;
    hexResult += h;
  }

  var formatted = hexResult.toUpperCase();
  var result = "";
  for (var i = 0; i < formatted.length; i += 4) {
    if (result.length > 0) result += "-";
    result += formatted.substring(i, i + 4);
  }

  return result;
}

// ============================================
// VERIFICAR CLAVE DE ACTIVACIÓN
// ============================================
async function verificarClave(claveIngresada) {
  var deviceId = await generateDeviceId();
  var claveEsperada = computeActivationKey(deviceId);

  console.log("🔐 [Activacion] Clave esperada:", claveEsperada);
  console.log("🔐 [Activacion] Clave ingresada:", claveIngresada);

  return claveIngresada.trim().toUpperCase() === claveEsperada;
}

// ============================================
// GUARDAR ACTIVACIÓN PERMANENTE
// ============================================
async function guardarActivacion() {
  // Guardar en localStorage
  try {
    localStorage.setItem(ACTIVATION_KEY, "true");
  } catch (e) {
    console.error("❌ [Activacion] Error guardando en localStorage:", e);
  }

  // Guardar también en SQLite si está disponible (doble persistencia)
  if (storageMode === "sqlite" && db) {
    try {
      await db.execute(
        "INSERT OR REPLACE INTO config (clave, valor, timestamp) VALUES (?, ?, ?)",
        [ACTIVATION_KEY, "true", Date.now()]
      );
    } catch (e) {
      console.warn("⚠️ [Activacion] No se pudo guardar en SQLite:", e.message);
    }
  }
}

// ============================================
// VERIFICAR SI EL DISPOSITIVO ESTÁ ACTIVADO
// ============================================
async function isActivated() {
  // 1. Revisar localStorage primero (rápido)
  var localState = localStorage.getItem(ACTIVATION_KEY);
  if (localState === "true") return true;

  // 2. Revisar SQLite si está disponible
  if (storageMode === "sqlite" && db) {
    try {
      var result = await db.execute(
        "SELECT valor FROM config WHERE clave = ?",
        [ACTIVATION_KEY]
      );
      if (result.values && result.values.length > 0 && result.values[0].valor === "true") {
        // Sincronizar a localStorage
        try { localStorage.setItem(ACTIVATION_KEY, "true"); } catch (e) {}
        return true;
      }
    } catch (e) {
      console.warn("⚠️ [Activacion] Error leyendo SQLite:", e.message);
    }
  }

  return false;
}

// ============================================
// LIMPIAR ACTIVACIÓN (para pruebas / re-emitir)
// ============================================
async function limpiarActivacion() {
  try {
    localStorage.removeItem(ACTIVATION_KEY);
    localStorage.removeItem(DEVICE_ID_KEY);
    localStorage.removeItem("gplus_device_salt");
  } catch (e) {}

  if (storageMode === "sqlite" && db) {
    try {
      await db.execute("DELETE FROM config WHERE clave = ?", [ACTIVATION_KEY]);
    } catch (e) {}
  }
}

// ============================================
// EXPORTAR COMO MÓDULO GLOBAL
// ============================================
window.Database = {
  initDatabase: initDatabase,
  getStorageMode: function() { return storageMode; },
  getProductosLocal: getProductosLocal,
  syncProductosLocal: syncProductosLocal,
  guardarVentaOffline: guardarVentaOffline,
  getVentasPendientes: getVentasPendientes,
  contarVentasPendientes: contarVentasPendientes,
  // Funciones para mermas
  guardarMermaOffline: guardarMermaOffline,
  getMermasPendientes: getMermasPendientes,
  contarMermasPendientes: contarMermasPendientes,
  getResumenOffline: getResumenOffline,
  limpiarVentasAntiguas: limpiarVentasAntiguas,
  obtenerListaServidores: obtenerListaServidores,
  agregarServidorConocido: agregarServidorConocido,
  descubrirServidor: descubrirServidor,
  // Funciones para nuevo producto
  getUltimoCodigoProducto: getUltimoCodigoProducto,
  guardarNuevoProducto: guardarNuevoProducto,
  guardarEntradaProductoCompleto: guardarEntradaProductoCompleto,
  getEntradaProductosPendientes: getEntradaProductosPendientes,
  contarEntradaProductosPendientes: contarEntradaProductosPendientes,
  // Funciones para abastecer (reabastecer productos existentes)
  actualizarStockProducto: actualizarStockProducto,
  guardarAbastecerOffline: guardarAbastecerOffline,
  contarAbastecerPendientes: contarAbastecerPendientes,
  getAbastecerPendientes: getAbastecerPendientes,
  // Funciones para gastos
  guardarGastoOffline: guardarGastoOffline,
  getGastosPendientes: getGastosPendientes,
  contarGastosPendientes: contarGastosPendientes,
  // Batch completo (reemplaza los 5 sync individuales)
  sincronizarCompleto: sincronizarCompleto,
  // Funciones para historial de ventas
  getVentasAgrupadasPorFactura: getVentasAgrupadasPorFactura,
  deshacerVenta: deshacerVenta,
  // Funciones de activación de dispositivos
  generateDeviceId: generateDeviceId,
  computeActivationKey: computeActivationKey,
  verificarClave: verificarClave,
  guardarActivacion: guardarActivacion,
  isActivated: isActivated,
  limpiarActivacion: limpiarActivacion,
  // Funciones para recetas
  syncRecetasLocal: syncRecetasLocal,
  getRecetasLocal: getRecetasLocal,
  // Funciones para ingredientes
  syncIngredientesLocal: syncIngredientesLocal,
  getIngredientesLocal: getIngredientesLocal,
  // Funciones para elaboraciones
  guardarElaboracionOffline: guardarElaboracionOffline,
  getElaboracionesPendientes: getElaboracionesPendientes,
  contarElaboracionesPendientes: contarElaboracionesPendientes,
};

console.log("📦 Database module loaded");
