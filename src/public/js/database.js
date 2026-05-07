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
  mermas: "posmovil_mermas_pending",
  entrada_productos: "posmovil_entrada_productos_pending",
  abastecer: "posmovil_abastecer_pending",
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

function asegurarStoreLocal() {
  if (!localStorage.getItem(LS_KEYS.productos)) guardarLocal(LS_KEYS.productos, []);
  if (!localStorage.getItem(LS_KEYS.ventas)) guardarLocal(LS_KEYS.ventas, []);
  if (!localStorage.getItem(LS_KEYS.mermas)) guardarLocal(LS_KEYS.mermas, []);
  if (!localStorage.getItem(LS_KEYS.entrada_productos)) guardarLocal(LS_KEYS.entrada_productos, []);
  if (!localStorage.getItem(LS_KEYS.abastecer)) guardarLocal(LS_KEYS.abastecer, []);
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

    // Tabla para mermas offline
    await db.execute(
      "CREATE TABLE IF NOT EXISTS mermas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, fecha_hora TEXT, synced INTEGER DEFAULT 0)"
    );

    // Tabla para entrada de nuevos productos (pendiente de sync al servidor)
    await db.execute(
      "CREATE TABLE IF NOT EXISTS entrada_productos_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT, nombre TEXT, cantidad INTEGER, precio_venta REAL, precio_costo REAL, fecha_hora TEXT, synced INTEGER DEFAULT 0)"
    );

    // Tabla para abastecer (reabastecer productos existentes - offline)
    await db.execute(
      "CREATE TABLE IF NOT EXISTS abastecer_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo_producto TEXT, nombre TEXT, cantidad INTEGER, fecha_hora TEXT, synced INTEGER DEFAULT 0)"
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
  var fechaHora = new Date().toISOString();
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
// FUNCIONES PARA MERMAS OFFLINE
// ============================================

// Guardar merma en SQLite/localStorage
async function guardarMermaOffline(merma) {
  if (!merma || !Array.isArray(merma.productos) || merma.productos.length === 0) {
    return false;
  }

  if (storageMode !== "sqlite" || !db) {
    var actuales = leerLocal(LS_KEYS.mermas, []);
    var lineas = [];
    for (var i = 0; i < merma.productos.length; i++) {
      var item = merma.productos[i];
      lineas.push({
        id: Date.now() + i + Math.floor(Math.random() * 1000),
        codigo_producto: item.codigo,
        nombre: item.nombre,
        cantidad: Number(item.cantidad || 0),
        fecha_hora: merma.fechaHora,
        synced: 0,
      });
    }
    return guardarLocal(LS_KEYS.mermas, actuales.concat(lineas));
  }

  try {
    for (var j = 0; j < merma.productos.length; j++) {
      var prod = merma.productos[j];
      await db.execute(
        "INSERT INTO mermas_pending (codigo_producto, nombre, cantidad, fecha_hora, synced) VALUES (?, ?, ?, ?, 0)",
        [prod.codigo, prod.nombre, prod.cantidad, merma.fechaHora]
      );
    }
    console.log("✅ Merma guardada offline (" + merma.productos.length + " productos)");
    return true;
  } catch (error) {
    console.error("❌ Error guardando merma offline:", error);
    // Fallback a localStorage
    var actuales2 = leerLocal(LS_KEYS.mermas, []);
    var lineas2 = [];
    for (var k = 0; k < merma.productos.length; k++) {
      var p = merma.productos[k];
      lineas2.push({
        id: Date.now() + k + Math.floor(Math.random() * 1000),
        codigo_producto: p.codigo,
        nombre: p.nombre,
        cantidad: Number(p.cantidad || 0),
        fecha_hora: merma.fechaHora,
        synced: 0,
      });
    }
    return guardarLocal(LS_KEYS.mermas, actuales2.concat(lineas2));
  }
}

// Obtener mermas pendientes de sync
async function getMermasPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.mermas, []).filter(function(m) {
      return Number(m.synced || 0) === 0;
    });
  }

  try {
    var result = await db.execute(
      "SELECT * FROM mermas_pending WHERE synced = 0 ORDER BY id"
    );
    return result.values || [];
  } catch (error) {
    console.error("❌ Error obteniendo mermas pendientes:", error);
    return leerLocal(LS_KEYS.mermas, []).filter(function(m) {
      return Number(m.synced || 0) === 0;
    });
  }
}

// Contar mermas pendientes
async function contarMermasPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.mermas, []).filter(function(m) {
      return Number(m.synced || 0) === 0;
    }).length;
  }

  try {
    var result = await db.execute(
      "SELECT COUNT(*) as count FROM mermas_pending WHERE synced = 0"
    );
    return result.values ? result.values[0].count : 0;
  } catch (error) {
    console.error("❌ Error contando mermas pendientes:", error);
    return 0;
  }
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

// Sincronizar mermas al servidor
// Acepta opcionalmente la URL del servidor como parámetro
async function sincronizarMermas(serverUrl) {
  var pendientes = await getMermasPendientes();
  
  if (pendientes.length === 0) {
    return { success: true, sincronizadas: 0 };
  }

  // Usar URL pasada como parámetro, o la del window, o nada
  var url = serverUrl || window.SERVER_URL;
  
  if (!url) {
    return { success: false, error: "No hay servidor configurado para sincronizar mermas" };
  }

  // Preparar datos para enviar
  var mermas = pendientes.map(function(m) {
    return {
      codigo: m.codigo_producto,
      nombre: m.nombre,
      cantidad: m.cantidad,
      fechaHora: m.fecha_hora,
    };
  });

  try {
    var response = await fetch(url + "/api/mermas", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-session-token": window.SESSION_TOKEN || "",
      },
      body: JSON.stringify({ mermas: mermas }),
    });

    if (!response.ok) {
      var errorText = await response.text();
      throw new Error("HTTP " + response.status + " - " + errorText);
    }

    var result = await response.json();

    // Marcar como sincronizadas
    var ids = pendientes.map(function(m) { return m.id; });
    await marcarMermasSynced(ids);

    console.log("✅ " + mermas.length + " mermas sincronizadas");
    return {
      success: true,
      sincronizadas: result.sincronizadas || mermas.length,
    };
  } catch (error) {
    console.error("❌ Error sincronizando mermas:", error.message);
    return { success: false, error: error.message };
  }
}

// ============================================
// FUNCIONES PARA ENTRADA DE PRODUCTOS (sync al servidor)
// ============================================

// Obtener entrada de productos pendientes de sync
async function getEntradaProductosPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.entrada_productos, []).filter(function(p) {
      return Number(p.synced || 0) === 0;
    });
  }

  try {
    var result = await db.execute(
      "SELECT * FROM entrada_productos_pending WHERE synced = 0 ORDER BY id"
    );
    return result.values || [];
  } catch (error) {
    console.error("❌ Error obteniendo entrada productos pendientes:", error);
    return leerLocal(LS_KEYS.entrada_productos, []).filter(function(p) {
      return Number(p.synced || 0) === 0;
    });
  }
}

// Contar entrada de productos pendientes
async function contarEntradaProductosPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.entrada_productos, []).filter(function(p) {
      return Number(p.synced || 0) === 0;
    }).length;
  }

  try {
    var result = await db.execute(
      "SELECT COUNT(*) as count FROM entrada_productos_pending WHERE synced = 0"
    );
    return result.values ? result.values[0].count : 0;
  } catch (error) {
    console.error("❌ Error contando entrada productos pendientes:", error);
    return 0;
  }
}

// Contar abastecimientos pendientes
async function contarAbastecerPendientes() {
  if (storageMode !== "sqlite" || !db) {
    return leerLocal(LS_KEYS.abastecer, []).filter(function(a) {
      return Number(a.synced || 0) === 0;
    }).length;
  }

  try {
    var result = await db.execute(
      "SELECT COUNT(*) as count FROM abastecer_pending WHERE synced = 0"
    );
    return result.values ? result.values[0].count : 0;
  } catch (error) {
    console.error("❌ Error contando abastecimientos pendientes:", error);
    return 0;
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

// Sincronizar entrada de productos al servidor
// Acepta opcionalmente la URL del servidor como parámetro
async function sincronizarEntradaProductos(serverUrl) {
  logSyncDebugAPK("🔍 [DEBUG] sincronizarEntradaProductos() iniciado");
  
  var pendientes = await getEntradaProductosPendientes();
  logSyncDebugAPK("🔍 [DEBUG] Pendientes entrada productos: " + pendientes.length);
  
  if (pendientes.length === 0) {
    logSyncDebugAPK("🔍 [DEBUG] No hay entradas pendientes para sincronizar");
    return { success: true, sincronizadas: 0 };
  }

  // Usar URL pasada como parámetro, o la del window, o nada
  var url = serverUrl || window.SERVER_URL;
  logSyncDebugAPK("🔍 [DEBUG] URL del servidor para entrada productos: " + url);
  
  if (!url) {
    logSyncDebugAPK("🔍 [DEBUG] No hay URL configurada", "error");
    return { success: false, error: "No hay servidor configurado para sincronizar entrada de productos" };
  }

  // Preparar datos para enviar
  var productos = pendientes.map(function(p) {
    return {
      codigo: p.codigo,
      nombre: p.nombre,
      cantidad: p.cantidad,
      precio_venta: p.precio_venta,
      precio_costo: p.precio_costo,
      fecha_hora: p.fecha_hora,
    };
  });
  
  logSyncDebugAPK("🔍 [DEBUG] Enviando " + productos.length + " productos al servidor");

  try {
    logSyncDebugAPK("🔍 [DEBUG] Haciendo fetch a: " + url + "/api/entrada-productos");
    var response = await fetch(url + "/api/entrada-productos", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-session-token": window.SESSION_TOKEN || "",
      },
      body: JSON.stringify({ productos: productos }),
    });

    logSyncDebugAPK("🔍 [DEBUG] Respuesta del servidor: " + response.status + " " + response.statusText);

    if (!response.ok) {
      var errorText = await response.text();
      logSyncDebugAPK("🔍 [DEBUG] Error response: " + errorText, "error");
      throw new Error("HTTP " + response.status + " - " + errorText);
    }

    var result = await response.json();
    logSyncDebugAPK("🔍 [DEBUG] Resultado del servidor: " + (result && result.mensaje ? result.mensaje : "OK"));
    if (result && Array.isArray(result.detalles) && result.detalles.length > 0) {
      var maxDetalles = Math.min(result.detalles.length, 3);
      for (var d = 0; d < maxDetalles; d++) {
        var detalle = result.detalles[d];
        logSyncDebugAPK(
          "🧾 Excel -> " + detalle.codigo + " (Productos F" + detalle.filaProductos + ", Entrada F" + detalle.filaEntrada + ")"
        );
      }
      if (result.detalles.length > maxDetalles) {
        logSyncDebugAPK("🧾 ... y " + (result.detalles.length - maxDetalles) + " más");
      }
    }

    // Marcar como sincronizadas
    var ids = pendientes.map(function(p) { return p.id; });
    await marcarEntradaProductosSynced(ids);
    logSyncDebugAPK("🔍 [DEBUG] Marcadas como sincronizadas: " + ids.length + " productos");

    logSyncDebugAPK("✅ " + productos.length + " entrada de productos sincronizadas");
    return {
      success: true,
      sincronizadas: result.sincronizadas || productos.length,
    };
  } catch (error) {
    logSyncDebugAPK("❌ Error sincronizando entrada de productos: " + error.message, "error");
    return { success: false, error: error.message };
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
  var fechaHora = new Date().toISOString();

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
  // Funciones para mermas
  guardarMermaOffline: guardarMermaOffline,
  getMermasPendientes: getMermasPendientes,
  contarMermasPendientes: contarMermasPendientes,
  sincronizarMermas: sincronizarMermas,
  getResumenOffline: getResumenOffline,
  limpiarVentasAntiguas: limpiarVentasAntiguas,
  obtenerServidorCacheado: obtenerServidorCacheado,
  guardarServidorCacheado: guardarServidorCacheado,
  limpiarServidorCacheado: limpiarServidorCacheado,
  descubrirServidor: descubrirServidor,
  // Funciones para nuevo producto
  getUltimoCodigoProducto: getUltimoCodigoProducto,
  guardarNuevoProducto: guardarNuevoProducto,
  guardarEntradaProductoCompleto: guardarEntradaProductoCompleto,
  // Funciones para sincronizar entrada de productos
  getEntradaProductosPendientes: getEntradaProductosPendientes,
  contarEntradaProductosPendientes: contarEntradaProductosPendientes,
  sincronizarEntradaProductos: sincronizarEntradaProductos,
   // Funciones para abastecer (reabastecer productos existentes)
  actualizarStockProducto: actualizarStockProducto,
  guardarAbastecerOffline: guardarAbastecerOffline,
  contarAbastecerPendientes: contarAbastecerPendientes,
  // Funciones para historial de ventas
  getVentasAgrupadasPorFactura: getVentasAgrupadasPorFactura,
  deshacerVenta: deshacerVenta,
};

console.log("📦 Database module loaded");
