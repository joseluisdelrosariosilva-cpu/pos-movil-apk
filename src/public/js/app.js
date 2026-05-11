// ============================================
// POS MÓVIL - Lógica de la aplicación
// ============================================
// Cargar database module primero
// (se define como global en database.js)

// Configuración
let dbInicializado = false;
let modoOffline = false;
let serverUrlCacheado = null;
let ultimoToggleCarritoTs = 0;

const DECIMALES = 3;
const FACTOR_DECIMALES = 10 ** DECIMALES;
const PASO_CANTIDAD = 1;
const KEY_ULTIMA_SYNC = "posmovil_ultima_sync";
let ultimaSyncISO = localStorage.getItem(KEY_ULTIMA_SYNC) || null;

function redondear3(valor) {
  return Math.round((Number(valor) || 0) * FACTOR_DECIMALES) / FACTOR_DECIMALES;
}

function parsearNumero(valor) {
  if (valor === null || valor === undefined) return 0;
  var normalizado = String(valor).replace(",", ".").trim();
  var num = parseFloat(normalizado);
  return Number.isFinite(num) ? redondear3(num) : 0;
}

function formatearNumero(valor) {
  var n = redondear3(valor);
  var texto = n.toFixed(3).replace(/\.?0+$/, "");
  return texto === "" || texto === "-" ? "0" : texto;
}

function formatearMoneda(valor) {
  return "$" + formatearNumero(valor);
}

function guardarUltimaSync(fecha = new Date()) {
  ultimaSyncISO = fecha.toISOString();
  localStorage.setItem(KEY_ULTIMA_SYNC, ultimaSyncISO);
}

// ============================================
// OBTENER URL DEL SERVIDOR
// Usa localhost si estamos en browser, auto-descubre si estamos en APK
// También prueba rango 10.x.x.x para tethering
// ============================================
function obtenerUrlServidor() {
  // Si ya tenemos URL cacheada, usarla
  if (serverUrlCacheado) {
    console.log("📡 [obtenerUrlServidor] Usando cache:", serverUrlCacheado);
    return serverUrlCacheado;
  }
  
  // En browser, usar window.location.origin
  if (!window.Capacitor || !window.Capacitor.isNativePlatform()) {
    var url = window.location.origin;
    console.log("📡 [obtenerUrlServidor] Navegador:", url);
    return url;
  }
  
  // En APK: probar la URL manual o auto-descubrir
  if (window.SERVER_URL) {
    serverUrlCacheado = window.SERVER_URL;
    console.log("📡 [obtenerUrlServidor] URL manual:", serverUrlCacheado);
    return serverUrlCacheado;
  }
  
  // Debug: mostrar información de Capacitor
  console.log("📡 [obtenerUrlServidor] Capacitor:", JSON.stringify(window.Capacitor));
  console.log("📡 [obtenerUrlServidor] isNativePlatform:", window.Capacitor?.isNativePlatform?.());
  console.log("📡 [obtenerUrlServidor] SERVER_URL:", window.SERVER_URL);

// Fallback: localhost
  console.log("📡 [obtenerUrlServidor] Usando fallback localhost");
  return "http://localhost:3000";
}

// Referencias al database module
const DB = window.Database || {};

// Estado de la aplicación
let productos = [];
let carrito = [];

// Elementos del DOM
const productosContainer = document.getElementById("productosContainer");
const carritoLista = document.getElementById("carritoLista");
const totalItems = document.getElementById("totalItems");
const totalPrecio = document.getElementById("totalPrecio");
const totalPagarSpan = document.getElementById("totalPagar");
const efectivoInput = document.getElementById("efectivo");
const transferenciaInput = document.getElementById("transferencia");
const btnPagar = document.getElementById("btnPagar");
const mensajeDiv = document.getElementById("mensaje");
const searchInput = document.getElementById("searchInput");
const searchClear = document.getElementById("searchClear");
const modalVuelto = document.getElementById("modalVuelto");
const modalVueltoMensaje = document.getElementById("modalVueltoMensaje");
const modalTotal = document.getElementById("modalTotal");
const modalPagado = document.getElementById("modalPagado");
const modalVueltoMonto = document.getElementById("modalVueltoMonto");
const modalCancelar = document.getElementById("modalCancelar");
const modalConfirmar = document.getElementById("modalConfirmar");
const estadoModoEl = document.getElementById("estadoModo");
const estadoPendientesEl = document.getElementById("estadoPendientes");
const estadoUltimaSyncEl = document.getElementById("estadoUltimaSync");
function logSyncAPK(mensaje, tipo) {
  if (tipo === "error") {
    console.error("📱[SYNC APK]", mensaje);
  } else if (tipo === "warning") {
    console.warn("📱[SYNC APK]", mensaje);
  } else {
    console.log("📱[SYNC APK]", mensaje);
  }
}

async function contarPendientesTotales() {
  var pendientesVentas = 0;
  var pendientesMermas = 0;
  var pendientesEntradaProductos = 0;
  var pendientesAbastecer = 0;

  if (DB.contarVentasPendientes) {
    pendientesVentas = await DB.contarVentasPendientes();
  }

  if (DB.contarMermasPendientes) {
    pendientesMermas = await DB.contarMermasPendientes();
  }

  if (DB.contarEntradaProductosPendientes) {
    pendientesEntradaProductos = await DB.contarEntradaProductosPendientes();
  }

  if (DB.contarAbastecerPendientes) {
    pendientesAbastecer = await DB.contarAbastecerPendientes();
  }

  return pendientesVentas + pendientesMermas + pendientesEntradaProductos + pendientesAbastecer;
}

async function actualizarPanelEstado() {
  if (estadoModoEl) {
    var storage = DB.getStorageMode ? DB.getStorageMode() : "-";
    var modoTexto = modoOffline ? "Offline" : "Online";
    estadoModoEl.textContent = modoTexto + " (" + storage + ")";
  }

  if (estadoPendientesEl) {
    try {
      var pendientes = await contarPendientesTotales();
      estadoPendientesEl.textContent = String(pendientes);
    } catch (_) {
      estadoPendientesEl.textContent = "-";
    }
  }

  if (estadoUltimaSyncEl) {
    if (!ultimaSyncISO) {
      estadoUltimaSyncEl.textContent = "Nunca";
    } else {
      var fecha = new Date(ultimaSyncISO);
      if (Number.isNaN(fecha.getTime())) {
        estadoUltimaSyncEl.textContent = "Nunca";
      } else {
        estadoUltimaSyncEl.textContent = fecha.toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
        });
      }
    }
  }
}

// ============================================
// VARIABLES DEL MODAL
// ============================================
let resolverModal = null;

// ============================================
// FUNCIÓN PARA MOSTRAR MODAL DE VUELTO
// ============================================
function mostrarModalVuelto(total, pagado, vuelto) {
  if (!modalVuelto) {
    return Promise.resolve(
      confirm(
        `Vuelto a entregar: ${formatearMoneda(vuelto)}. Confirmas que entregaste el vuelto?`,
      ),
    );
  }

  return new Promise((resolve) => {
    resolverModal = resolve;
    modalVueltoMensaje.textContent =
      "Verifica el monto y confirma la entrega del vuelto.";
    modalTotal.textContent = formatearMoneda(total);
    modalPagado.textContent = formatearMoneda(pagado);
    modalVueltoMonto.textContent = formatearMoneda(vuelto);
    modalVuelto.classList.remove("hidden");
    modalConfirmar.focus();
  });
}

function cerrarModalVuelto(confirmado) {
  if (!modalVuelto || resolverModal === null) return;
  const resolver = resolverModal;
  resolverModal = null;
  modalVuelto.classList.add("hidden");
  resolver(confirmado);
}

// ============================================
// CARGAR PRODUCTOS
// ============================================
async function cargarProductos() {
  console.log("📡 Cargando productos...");
  mostrarMensaje("Cargando...", "info");
  
  var serverUrl = obtenerUrlServidor();
  var fetchUrl = serverUrl + "/api/productos";
  
  console.log("📡 URL:", fetchUrl);

  var timeoutId = setTimeout(function() { 
    console.log("⏱️ Timeout");
  }, 8000);

try {
    console.log("📡 Usando fetch nativo...");
    var controller = new AbortController();
    timeoutId = setTimeout(function() { controller.abort(); }, 8000);
    
    var response = await fetch(fetchUrl, {
      method: "GET",
      headers: { "x-session-token": window.SESSION_TOKEN || "" },
      signal: controller.signal,
    });

    clearTimeout(timeoutId);
    console.log("📡 Status:", response.status);
    
    if (response.ok) {
      modoOffline = false;
      var data = await response.json();
      productos = Array.isArray(data) ? data : (data.productos || []);

      if (productos.length > 0) {
        // Guardar en SQLite para uso offline
        if (dbInicializado && DB.syncProductosLocal) {
          await DB.syncProductosLocal(productos);
        }
        
        renderizarProductos(productos);
        mostrarMensaje(productos.length + " productos", "exito", 2000);
        console.log("✅ " + productos.length + " productos");
      } else {
        productosContainer.innerHTML = '<p class="info">Sin productos</p>';
        mostrarMensaje("Sin productos", "warning");
      }

      await actualizarPanelEstado();
    } else {
      throw new Error("HTTP " + response.status);
    }
} catch(e) {
    clearTimeout(timeoutId);
    var errorMsg = "Error: " + e.message;
    
    if (e.name === "AbortError") {
      errorMsg = "⚠️ Timeout (8s) - no responde";
    } else if (e.message && e.message.includes("Failed to fetch")) {
      errorMsg = "⚠️ Failed to fetch - Android bloquea HTTP?";
    }
    
    console.log("❌ " + errorMsg);
    
    // Intentar cargar desde SQLite si falla el servidor
    if (dbInicializado && DB.getProductosLocal) {
      var productosLocales = await DB.getProductosLocal();
      if (productosLocales.length > 0) {
        modoOffline = true;
        productos = productosLocales;
        renderizarProductos(productos);
        mostrarMensaje("📦 Modo offline: " + productosLocales.length + " productos", "info", 3000);
        console.log("📦 Cargados " + productosLocales.length + " productos desde SQLite");
        await actualizarPanelEstado();
        return;
      }
    }
    
    mostrarMensaje("Sin conexion", "warning");
    productosContainer.innerHTML = '<p class="info">Sin conexion.<br>Activa el servidor.</p>';
    await actualizarPanelEstado();
  }
}

// ============================================
// 2. RENDERIZAR PRODUCTOS
// ============================================
// ============================================
// 2. RENDERIZAR PRODUCTOS (CON BOTÓN -)
// ============================================
function renderizarProductos(lista) {
  if (!lista || lista.length === 0) {
    productosContainer.innerHTML =
      '<p class="info">No hay productos disponibles</p>';
    return;
  }

  console.log("🎨 Renderizando productos:", lista.length);

  let html = "";

  for (let i = 0; i < lista.length; i++) {
    const producto = lista[i];

    const codigo = producto.codigo || "";
    const nombre = producto.producto || producto.nombre || "Producto";
    const precio = parsearNumero(producto.precio || 0);
    const stock = parsearNumero(producto.disponibilidad || producto.stock || 0);
    const sinStock = stock <= 0;

    // Verificar si el producto está en el carrito
    const itemEnCarrito = carrito.find((item) => item.codigo === codigo);
    const cantidadEnCarrito = itemEnCarrito ? itemEnCarrito.cantidad : 0;

    const codigoEscapado = codigo.replace(/"/g, "&quot;");

    html += '<div class="producto-card" data-codigo="' + codigo + '">';
    html += '<div class="producto-info">';
    html += "<h3>" + nombre + "</h3>";
    html += '<div class="precio">' + formatearMoneda(precio) + "</div>";
    html +=
      '<div class="stock">Stock: ' +
      formatearNumero(stock) +
      (sinStock ? " (Sin stock)" : "") +
      "</div>";
    html += "</div>";

    html += '<div class="producto-actions">';

    // En la parte de los botones, debe quedar así:

    // Botón DISMINUIR (-)
    if (cantidadEnCarrito > 0) {
      html +=
        '<button class="btn-cantidad btn-disminuir" onclick="disminuirCantidad(\'' +
        codigoEscapado +
        "')\">−</button>";
    } else {
      html += '<div style="width: 48px; height: 48px;"></div>';
    }

    // Contador
    html +=
      '<span class="cantidad-seleccionada" id="cant-' +
      codigo +
      '" onclick="editarCantidad(\'' +
      codigoEscapado +
      '\', this)" title="Editar cantidad">' +
      formatearNumero(cantidadEnCarrito) +
      "</span>";

    // Botón AUMENTAR (+)
    const puedeSumar = stock > 0 && redondear3(cantidadEnCarrito + PASO_CANTIDAD) <= redondear3(stock);
    html +=
      '<button class="btn-cantidad" ' +
      (puedeSumar ? "" : "disabled ") +
      'onclick="agregarAlCarrito(\'' +
      codigoEscapado +
      "')\">+</button>";

    html += "</div>"; // Cierra producto-actions
    html += "</div>"; // Cierra producto-card
  }

  productosContainer.innerHTML = html;
  console.log("✅ Productos renderizados con botones -");
}

// ============================================
// 3. AGREGAR AL CARRITO
// ============================================
window.agregarAlCarrito = (codigo) => {
  console.log("➕ Agregando producto con código:", codigo);

  const producto = productos.find((p) => p.codigo == codigo);

  if (!producto) {
    console.error("❌ Producto no encontrado:", codigo);
    mostrarMensaje("Error: producto no encontrado", "error");
    return;
  }

  const itemExistente = carrito.find((item) => item.codigo === codigo);
  const stock = parsearNumero(producto.disponibilidad || producto.stock || 0);
  if (stock <= 0) {
    mostrarMensaje("Sin stock", "error", 2000);
    return;
  }

  if (itemExistente) {
    const nuevaCantidad = redondear3(itemExistente.cantidad + PASO_CANTIDAD);
    if (nuevaCantidad <= stock) {
      itemExistente.cantidad = nuevaCantidad;
    } else {
      mostrarMensaje("Stock insuficiente", "error", 2000);
      return;
    }
  } else {
    carrito.push({
      codigo: codigo,
      nombre: producto.producto || producto.nombre,
      precio: parsearNumero(producto.precio || 0),
      cantidad: redondear3(PASO_CANTIDAD),
      maxStock: stock,
    });
  }

  actualizarVistaCarrito();
  renderizarProductos(productos);
  actualizarContadorProducto(codigo);
  reapplySearchFilter();
};

window.editarCantidad = (codigo, spanRef) => {
  const producto = productos.find((p) => String(p.codigo) === String(codigo));
  if (!producto || !spanRef) return;

  // Evitar doble editor en el mismo elemento
  if (spanRef.querySelector("input")) return;

  const stock = parsearNumero(producto.disponibilidad || producto.stock || 0);
  const itemExistente = carrito.find((item) => String(item.codigo) === String(codigo));
  const actual = itemExistente ? itemExistente.cantidad : PASO_CANTIDAD;
  const valorOriginal = itemExistente ? itemExistente.cantidad : 0;

  const input = document.createElement("input");
  input.type = "number";
  input.inputMode = "decimal";
  input.min = "0";
  input.max = String(stock);
  input.step = "0.001";
  input.value = redondear3(actual).toFixed(3);
  input.className = "cantidad-input-inline";

  spanRef.dataset.valorPrevio = spanRef.textContent;
  spanRef.textContent = "";
  spanRef.appendChild(input);
  input.focus();
  input.select();

  const aplicarCantidad = () => {
    const cantidad = redondear3(parsearNumero(input.value));

    if (cantidad > stock) {
      mostrarMensaje("Cantidad supera el stock", "error", 2200);
      spanRef.textContent = formatearNumero(valorOriginal);
      return;
    }

    if (cantidad <= 0) {
      carrito = carrito.filter((item) => String(item.codigo) !== String(codigo));
    } else if (itemExistente) {
      itemExistente.cantidad = cantidad;
    } else {
      carrito.push({
        codigo: codigo,
        nombre: producto.producto || producto.nombre,
        precio: parsearNumero(producto.precio || 0),
        cantidad: cantidad,
        maxStock: stock,
      });
    }

    actualizarVistaCarrito();
    renderizarProductos(productos);
    reapplySearchFilter();
  };

  const cancelar = () => {
    spanRef.textContent = formatearNumero(valorOriginal);
  };

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.blur();
    }
    if (e.key === "Escape") {
      e.preventDefault();
      cancelar();
    }
  });

  input.addEventListener("blur", aplicarCantidad, { once: true });
};

// ============================================
// 3B. DISMINUIR CANTIDAD DEL CARRITO
// ============================================
window.disminuirCantidad = (codigo) => {
  console.log("➖ Disminuyendo producto con código:", codigo);

  const index = carrito.findIndex((item) => item.codigo === codigo);

  if (index !== -1) {
    if (carrito[index].cantidad > PASO_CANTIDAD) {
      const nuevaCantidad = redondear3(carrito[index].cantidad - PASO_CANTIDAD);
      if (nuevaCantidad > 0) {
        carrito[index].cantidad = nuevaCantidad;
        console.log(`🔽 Nueva cantidad: ${carrito[index].cantidad}`);
      } else {
        carrito.splice(index, 1);
      }
    } else {
      // Si es 1, eliminar del carrito
      console.log("🗑️ Eliminando producto (cantidad llegó a 0)");
      carrito.splice(index, 1);
    }
  }

  actualizarVistaCarrito();
  renderizarProductos(productos);
  actualizarContadorProducto(codigo);
  reapplySearchFilter();
};

// ============================================
// 4. ELIMINAR DEL CARRITO
// ============================================
window.eliminarItemCarrito = (codigo) => {
  console.log("🗑️ Eliminando producto con código:", codigo);

  carrito = carrito.filter((item) => item.codigo !== codigo);

  actualizarVistaCarrito();
  actualizarContadorProducto(codigo);
};

// ============================================
// 5. ACTUALIZAR CONTADOR DE PRODUCTO
// ============================================
function actualizarContadorProducto(codigo) {
  const item = carrito.find((i) => i.codigo === codigo);
  const span = document.getElementById(`cant-${codigo}`);
  if (span) {
    span.textContent = formatearNumero(item ? item.cantidad : 0);
  }
}

// ============================================
// 6. ACTUALIZAR VISTA DEL CARRITO
// ============================================

function actualizarVistaCarrito() {
  const totalProductos = redondear3(carrito.reduce((sum, item) => sum + item.cantidad, 0));
  const totalPagar = carrito.reduce(
    (sum, item) => sum + item.cantidad * item.precio,
    0,
  );

  // Actualizar resumen
  totalItems.textContent = `${formatearNumero(totalProductos)} ${totalProductos === 1 ? "producto" : "productos"}`;
  totalPrecio.textContent = formatearMoneda(totalPagar);
  totalPagarSpan.textContent = formatearMoneda(totalPagar);

  // Si el carrito está vacío
  if (carrito.length === 0) {
    carritoLista.innerHTML =
      '<li style="text-align: center; padding: 24px; color: var(--text-secondary);">🛒 Carrito vacío</li>';
    btnPagar.disabled = true;
    cerrarCarrito();
    return;
  }

  // Generar HTML para cada producto
  let html = "";

  for (let i = 0; i < carrito.length; i++) {
    const item = carrito[i];

    // Escapar comillas en el código
    const codigoEscapado = item.codigo.replace(/"/g, "&quot;");
    const subtotal = redondear3(item.cantidad * item.precio);

    // Construir cada item del carrito
    html += '<li class="carrito-item">';

    // Información del producto (lado izquierdo)
    html += '<div class="carrito-item-info">';
    html += '<div class="carrito-item-nombre">' + item.nombre + "</div>";
    html +=
      '<div class="carrito-item-detalle">' +
      formatearMoneda(item.precio) +
      " c/u</div>";
    html += "</div>";

    // Controles y subtotal (lado derecho)
    html += '<div style="display: flex; align-items: center; gap: 8px;">';

    // Subtotal
    html += '<span class="carrito-item-subtotal">' + formatearMoneda(subtotal) + "</span>";

    // En la parte de los controles del carrito:

    // Botón DISMINUIR
    html +=
      '<button class="btn-cantidad-carrito" onclick="disminuirCantidad(\'' +
      codigoEscapado +
      "')\">−</button>";

    // Cantidad actual
    html +=
      '<span class="cantidad-seleccionada" style="min-width: 36px;" onclick="editarCantidad(\'' +
      codigoEscapado +
      '\', this)" title="Editar cantidad">' +
      formatearNumero(item.cantidad) +
      "</span>";

    // Botón AUMENTAR
    html +=
      '<button class="btn-cantidad-carrito" onclick="agregarAlCarrito(\'' +
      codigoEscapado +
      "')\">+</button>";

    // Botón ELIMINAR
    html +=
      '<button class="btn-eliminar-item" onclick="eliminarItemCarrito(\'' +
      codigoEscapado +
      "')\">✕</button>";

    html += "</div>"; // Cierra el div de controles
    html += "</li>"; // Cierra el item
  }

  // Insertar el HTML generado
  carritoLista.innerHTML = html;

  // Habilitar botón de pago
  btnPagar.disabled = false;

  // Actualizar cálculo de pago
  calcularPago();
}

// ============================================
// 7. CÁLCULO DE PAGO
// ============================================
function calcularPago() {
  const total = redondear3(carrito.reduce(
    (sum, item) => sum + item.cantidad * item.precio,
    0,
  ));
  const efectivo = parsearNumero(efectivoInput.value);
  const transferencia = parsearNumero(transferenciaInput.value);

  const pagado = redondear3(efectivo + transferencia);
  const diferencia = redondear3(pagado - total);

  btnPagar.classList.remove("exacto", "falta", "vuelto");

  if (Math.abs(diferencia) < 1 / FACTOR_DECIMALES) {
    btnPagar.textContent = "Pagar ✓";
    btnPagar.classList.add("exacto");
    btnPagar.disabled = false;
  } else if (diferencia > 0) {
    btnPagar.textContent = `Vuelto ${formatearMoneda(diferencia)}`;
    btnPagar.classList.add("vuelto");
    btnPagar.disabled = false;
  } else {
    btnPagar.textContent = `Faltan ${formatearMoneda(Math.abs(diferencia))}`;
    btnPagar.classList.add("falta");
    btnPagar.disabled = true;
  }
}

// ============================================
// REFLEJAR VENTA EN STOCK LOCAL (UI + SQLite)
// ============================================
async function reflejarVentaEnStockLocal(productosVendidos) {
  if (!Array.isArray(productosVendidos) || productosVendidos.length === 0) return;

  for (var i = 0; i < productosVendidos.length; i++) {
    var vendido = productosVendidos[i];
    var prod = productos.find(function(p) {
      return String(p.codigo) === String(vendido.codigo);
    });

    if (!prod) continue;

    var stockActual = parsearNumero(prod.disponibilidad || prod.stock || 0);
    var cantidadVendida = parsearNumero(vendido.cantidad || 0);
    var nuevoStock = redondear3(Math.max(0, stockActual - cantidadVendida));

    if (Object.prototype.hasOwnProperty.call(prod, "disponibilidad")) {
      prod.disponibilidad = nuevoStock;
    } else {
      prod.stock = nuevoStock;
    }
  }

  renderizarProductos(productos);

  // Persistir stock local para modo offline
  if (dbInicializado && DB.syncProductosLocal) {
    try {
      await DB.syncProductosLocal(productos);
    } catch (e) {
      console.log("⚠️ No se pudo persistir stock local:", e.message);
    }
  }
}

// ============================================
// 8. FINALIZAR VENTA (soporta online y offline)
// ============================================

async function finalizarVenta() {
  if (carrito.length === 0) {
    mostrarMensaje("Agrega productos al carrito", "error");
    return;
  }

  let efectivo = parsearNumero(efectivoInput.value);
  let transferencia = parsearNumero(transferenciaInput.value);
  const total = redondear3(carrito.reduce(
    (sum, item) => sum + item.cantidad * item.precio,
    0,
  ));

  const pagado = redondear3(efectivo + transferencia);
  let vuelto = redondear3(pagado - total);

  // 1. Si falta dinero, NO permitir
  if (vuelto < 0) {
    mostrarMensaje(`❌ Faltan ${formatearMoneda(Math.abs(vuelto))}`, "error");
    return;
  }

  // 2. Si hay vuelto, preguntar con modal
  if (vuelto > 0) {
    const confirmar = await mostrarModalVuelto(total, pagado, vuelto);

    if (!confirmar) {
      mostrarMensaje("Venta cancelada", "info", 2000);
      return;
    }

    // 3. AJUSTAR EFECTIVO AL MONTO EXACTO
    if (efectivo >= vuelto) {
      efectivo = redondear3(efectivo - vuelto);
    } else {
      const restante = redondear3(vuelto - efectivo);
      efectivo = 0;
      transferencia = redondear3(Math.max(0, transferencia - restante));
    }

    efectivoInput.value = formatearNumero(efectivo);
    transferenciaInput.value = formatearNumero(transferencia);

    console.log(`💰 Vuelto entregado. Nuevo efectivo: ${formatearMoneda(efectivo)}`);
  }

  // 4. Preparar datos de la venta
  const ventaData = {
    facturaId: generarFacturaId(),
    fechaHora: (function() {
      const f = new Date();
      return f.getFullYear() + '-' + 
             String(f.getMonth()+1).padStart(2,'0') + '-' + 
             String(f.getDate()).padStart(2,'0') + 'T' +
             String(f.getHours()).padStart(2,'0') + ':' + 
             String(f.getMinutes()).padStart(2,'0') + ':' + 
             String(f.getSeconds()).padStart(2,'0') + '.' + 
             String(f.getMilliseconds()).padStart(3,'0');
    })(),
    pago: { efectivo, transferencia },
    productos: carrito.map((item) => ({
      codigo: item.codigo,
      nombre: item.nombre,
      cantidad: item.cantidad,
      precio: item.precio,
    })),
  };

  // 5. Procesar - SIEMPRE offline
  try {
    btnPagar.disabled = true;
    btnPagar.textContent = "Procesando...";

    // Guardar SIEMPRE en SQLite (offline)
    var guardado = await DB.guardarVentaOffline(ventaData);

    if (guardado) {
      modoOffline = true;
      await reflejarVentaEnStockLocal(ventaData.productos);
      
      var pendientes = await DB.contarVentasPendientes();
      mostrarMensaje(
        "✅ Venta guardada offline (" + pendientes + " pendientes)",
        "exito",
        3000,
      );
      await actualizarPanelEstado();
    } else {
      throw new Error("No se pudo guardar offline");
    }

    // Limpiar carrito
    carrito = [];
    efectivoInput.value = "0";
    transferenciaInput.value = "0";
    actualizarVistaCarrito();
    renderizarProductos(productos);
    reapplySearchFilter();
    cerrarCarrito();
    await actualizarPanelEstado();
  } catch (error) {
    console.error("❌ Error:", error);
    mostrarMensaje(error.message, "error");
  } finally {
    btnPagar.disabled = false;
    calcularPago();
  }
}

// ============================================
// PROCESAR MERMA DESDE CARRITO ACTUAL
// ============================================
window.procesarMerma = async function() {
  // Filtrar productos con cantidad > 0
  const productosConCantidad = carrito.filter(item => item.cantidad > 0);
  
  if (productosConCantidad.length === 0) {
    mostrarMensaje("No hay productos con cantidad > 0 en el carrito", "warning");
    return;
  }
  
  // Construir lista para confirmación
  let listaProductos = productosConCantidad.map(p => 
    `${p.nombre}: ${formatearNumero(p.cantidad)} unidades`
  ).join('\n');
  
  // Confirmar con el usuario
  const confirmar = confirm(
    `¿Desea sacar estas cantidades como merma?\n\n${listaProductos}\n\nSe descontarán del stock.`
  );
  
  if (!confirmar) {
    mostrarMensaje("Merma cancelada", "info", 2000);
    return;
  }
  
  // Preparar datos de merma
  const mermaData = {
    fechaHora: (function() {
      const f = new Date();
      return f.getFullYear() + '-' + 
             String(f.getMonth()+1).padStart(2,'0') + '-' + 
             String(f.getDate()).padStart(2,'0') + 'T' +
             String(f.getHours()).padStart(2,'0') + ':' + 
             String(f.getMinutes()).padStart(2,'0') + ':' + 
             String(f.getSeconds()).padStart(2,'0') + '.' + 
             String(f.getMilliseconds()).padStart(3,'0');
    })(),
    productos: productosConCantidad.map(item => ({
      codigo: item.codigo,
      nombre: item.nombre,
      cantidad: item.cantidad,
    }))
  };
  
  try {
    mostrarMensaje("Guardando merma...", "info");
    
    // Guardar merma offline
    var guardado = await DB.guardarMermaOffline(mermaData);
    
    if (!guardado) {
      throw new Error("No se pudo guardar la merma offline");
    }
    
    // Actualizar stock local (descontar cantidades)
    await reflejarMermaEnStockLocal(mermaData.productos);
    
    // Resetear cantidades del carrito a 0
    carrito = [];
    efectivoInput.value = "0";
    transferenciaInput.value = "0";
    actualizarVistaCarrito();
    renderizarProductos(productos);
    reapplySearchFilter();
    cerrarCarrito();
    
    // Cambiar a modo offline (todo se guarda offline hasta sincronizar)
    modoOffline = true;
    
    // Actualizar indicadores (modo y pendientes)
    await actualizarPanelEstado();
    await actualizarIndicadorSync();
    
    mostrarMensaje(
      "✅ Merma guardada offline (" + productosConCantidad.length + " productos)",
      "exito",
      3000
    );
    
  } catch (error) {
    console.error("❌ Error en procesarMerma:", error);
    mostrarMensaje(error.message, "error");
  }
}

// ============================================
// REFLEJAR MERMA EN STOCK LOCAL (UI + SQLite)
// ============================================
async function reflejarMermaEnStockLocal(productosMerma) {
  if (!Array.isArray(productosMerma) || productosMerma.length === 0) return;
  
  for (var i = 0; i < productosMerma.length; i++) {
    var merma = productosMerma[i];
    var prod = productos.find(function(p) {
      return String(p.codigo) === String(merma.codigo);
    });
    
    if (!prod) continue;
    
    var stockActual = parsearNumero(prod.disponibilidad || prod.stock || 0);
    var cantidadMerma = parsearNumero(merma.cantidad || 0);
    var nuevoStock = redondear3(Math.max(0, stockActual - cantidadMerma));
    
    if (Object.prototype.hasOwnProperty.call(prod, "disponibilidad")) {
      prod.disponibilidad = nuevoStock;
    } else {
      prod.stock = nuevoStock;
    }
  }
  
  renderizarProductos(productos);
  
  // Persistir stock local para modo offline
  if (dbInicializado && DB.syncProductosLocal) {
    try {
      await DB.syncProductosLocal(productos);
    } catch (e) {
      console.log("⚠️ No se pudo persistir stock local:", e.message);
    }
  }
}

// ============================================
// ABASTECER (reabastecer productos existentes - offline)
// ============================================

window.mostrarAbastecer = function() {
  var modal = document.getElementById("modalAbastecer");
  if (!modal) return;
  
  // Limpiar búsqueda anterior
  var searchInput = document.getElementById("abProductoSearch");
  var dropdown = document.getElementById("abProductoDropdown");
  var codigoInput = document.getElementById("abProductoCodigo");
  var infoDiv = document.getElementById("abProductoInfo");
  
  if (searchInput) searchInput.value = "";
  if (dropdown) {
    dropdown.innerHTML = "";
    dropdown.classList.add("hidden");
  }
  if (codigoInput) codigoInput.value = "";
  if (infoDiv) infoDiv.innerHTML = "";
  
  // Resetear cantidad
  var cantidadInput = document.getElementById("abCantidad");
  if (cantidadInput) {
    cantidadInput.value = "1";
  }
  
  modal.classList.remove("hidden");
  if (searchInput) searchInput.focus();
};

function cerrarModalAbastecer() {
  var modal = document.getElementById("modalAbastecer");
  if (modal) {
    modal.classList.add("hidden");
  }
  // Ocultar dropdown
  var dropdown = document.getElementById("abProductoDropdown");
  if (dropdown) dropdown.classList.add("hidden");
}

// Filtrar productos y mostrar dropdown
function filtrarProductosAbastecer(termino) {
  var dropdown = document.getElementById("abProductoDropdown");
  if (!dropdown) return;
  
  termino = termino.toLowerCase().trim();
  
  // Filtrar productos
  var resultados = [];
  for (var i = 0; i < productos.length; i++) {
    var p = productos[i];
    var nombre = (p.producto || p.nombre || "").toLowerCase();
    var codigo = (p.codigo || "").toLowerCase();
    if (nombre.includes(termino) || codigo.includes(termino)) {
      resultados.push(p);
    }
  }
  
  // Mostrar resultados
  dropdown.innerHTML = "";
  if (resultados.length === 0) {
    dropdown.innerHTML = '<div class="dropdown-item" style="color: var(--text-secondary);">No se encontraron productos</div>';
  } else {
    for (let j = 0; j < resultados.length; j++) {
      const p = resultados[j];
      const codigo = p.codigo || "";
      const nombre = p.producto || p.nombre || "Producto";
      const stock = parsearNumero(p.disponibilidad || p.stock || 0);
      
      const item = document.createElement("div");
      item.className = "dropdown-item";
      item.innerHTML = nombre + ' <span style="color: var(--text-secondary); font-size: 12px;">(Stock: ' + formatearNumero(stock) + ')</span>';
      item.dataset.codigo = codigo;
      item.addEventListener("click", function() {
        seleccionarProductoAbastecer(this.dataset.codigo, nombre, stock);
      });
      dropdown.appendChild(item);
    }
  }
  
  dropdown.classList.remove("hidden");
}

// Seleccionar producto del dropdown
function seleccionarProductoAbastecer(codigo, nombre, stock) {
  var codigoInput = document.getElementById("abProductoCodigo");
  var searchInput = document.getElementById("abProductoSearch");
  var infoDiv = document.getElementById("abProductoInfo");
  var dropdown = document.getElementById("abProductoDropdown");
  
  if (codigoInput) codigoInput.value = codigo;
  if (searchInput) searchInput.value = nombre;
  if (infoDiv) {
    infoDiv.innerHTML = '<strong>' + nombre + '</strong><div class="stock-info">Stock actual: ' + formatearNumero(stock) + '</div>';
  }
  if (dropdown) dropdown.classList.add("hidden");
}

window.confirmarAbastecer = async function() {
  var codigoInput = document.getElementById("abProductoCodigo");
  var cantidadInput = document.getElementById("abCantidad");
  
  if (!codigoInput || !cantidadInput) return;
  
  var codigo = codigoInput.value;
  var cantidad = parsearNumero(cantidadInput.value);
  
  if (!codigo) {
    mostrarMensaje("Selecciona un producto", "error");
    return;
  }
  
  if (cantidad <= 0) {
    mostrarMensaje("La cantidad debe ser mayor a 0", "error");
    return;
  }
  
  // Buscar el producto
  var producto = productos.find(function(p) {
    return String(p.codigo) === String(codigo);
  });
  
  if (!producto) {
    mostrarMensaje("Producto no encontrado", "error");
    return;
  }
  
  var nombre = producto.producto || producto.nombre || "";
  
  try {
    mostrarMensaje("Procesando abastecimiento...", "info");
    
    // Guardar abastecimiento offline
    var guardado = await DB.guardarAbastecerOffline({
      codigo: codigo,
      nombre: nombre,
      cantidad: cantidad
    });
    
    if (!guardado) {
      throw new Error("No se pudo guardar el abastecimiento");
    }
    
    // Recargar productos desde SQLite local (no bajar del servidor)
    productos = await DB.getProductosLocal();
    renderizarProductos(productos);
    
    // Cerrar modal
    cerrarModalAbastecer();
    
    // Actualizar contador de pendientes
    await actualizarPanelEstado();
    await actualizarIndicadorSync();
    
    mostrarMensaje(
      "✅ Abastecido: " + nombre + " (+" + formatearNumero(cantidad) + ")",
      "exito",
      3000
    );
    
  } catch (error) {
    console.error("❌ Error en abastecer:", error);
    mostrarMensaje("Error: " + error.message, "error");
  }
};

// Event listeners para abastecer
document.addEventListener("DOMContentLoaded", function() {
  var cancelarBtn = document.getElementById("btnCancelarAbastecer");
  if (cancelarBtn) {
    cancelarBtn.addEventListener("click", cerrarModalAbastecer);
  }
  
  var confirmarBtn = document.getElementById("btnConfirmarAbastecer");
  if (confirmarBtn) {
    confirmarBtn.addEventListener("click", function() {
      window.confirmarAbastecer();
    });
  }
  
  var modalAbastecer = document.getElementById("modalAbastecer");
  if (modalAbastecer) {
    modalAbastecer.addEventListener("click", function(e) {
      if (e.target === modalAbastecer) cerrarModalAbastecer();
    });
  }
  
  // Event listener para búsqueda de productos
  var searchInput = document.getElementById("abProductoSearch");
  if (searchInput) {
    searchInput.addEventListener("input", function() {
      filtrarProductosAbastecer(this.value);
    });
    
    // Mostrar dropdown al recibir foco
    searchInput.addEventListener("focus", function() {
      filtrarProductosAbastecer(this.value);
    });
    
    // Ocultar dropdown al hacer clic fuera
    searchInput.addEventListener("blur", function() {
      // Pequeño retraso para permitir que el clic en el dropdown se procese
      setTimeout(function() {
        var dropdown = document.getElementById("abProductoDropdown");
        if (dropdown) dropdown.classList.add("hidden");
      }, 200);
    });
  }
});

// ============================================
// GENERAR FACTURA ID PARA MODO OFFLINE
// Formato: [5 dígitos fecha serial Excel] + [5 dígitos consecutivo]
// USA LA REFERENCIA DE DATOS.XLSX - no calcula el prefijo
// Si la referencia es 46141, y pasaron 0 días → 46141, si pasó 1 día → 46142
// ============================================
function generarFacturaId() {
  const now = new Date();
  const hoyISO = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0');
  
  // Leer referencia guardada del servidor Y su fecha
  let ultimaReferencia = null;
  let fechaReferencia = null;
  try {
    ultimaReferencia = localStorage.getItem('posmovil_ultima_factura_referencia');
    fechaReferencia = localStorage.getItem('posmovil_fecha_referencia');
  } catch(e) {}
  
  // Si tenemos referencia Y fecha, usarlas
  if (ultimaReferencia && ultimaReferencia.length >= 10 && fechaReferencia) {
    const prefijoRef = ultimaReferencia.substring(0, 5);
    const sufijoRef = parseInt(ultimaReferencia.substring(5, 10)) || 0;
    
    // Calcular cuántos días pasaron desde la referencia
    const refDate = new Date(fechaReferencia);
    const todayDate = new Date(hoyISO);
    const msPerDay = 1000 * 60 * 60 * 24;
    const diasDiff = Math.floor((todayDate - refDate) / msPerDay);
    
    const prefijoHoy = String(Number(prefijoRef) + diasDiff).padStart(5, '0');
    
    if (diasDiff === 0) {
      // Mismo día que la referencia, incrementar sufijo
      const nuevoSufijo = sufijoRef + 1;
      const sufijo = nuevoSufijo.toString().padStart(5, '0');
      const nuevoId = prefijoHoy + sufijo;
      
      try {
        localStorage.setItem('posmovil_ultima_factura_referencia', nuevoId);
        localStorage.setItem('posmovil_ultimo_prefijo', prefijoHoy);
        localStorage.setItem('posmovil_ultimo_sufijo', nuevoSufijo.toString());
      } catch(e) {}
      
      console.log('📋 FacturaID generado: ' + nuevoId + ' (mismo día que referencia)');
      return nuevoId;
    } else {
      // Nuevo día (o días), empezar sufijo en 1
      const sufijo = "00001";
      const nuevoId = prefijoHoy + sufijo;
      
      try {
        localStorage.setItem('posmovil_ultima_factura_referencia', nuevoId);
        localStorage.setItem('posmovil_ultimo_prefijo', prefijoHoy);
        localStorage.setItem('posmovil_ultimo_sufijo', '1');
        localStorage.setItem('posmovil_fecha_referencia', hoyISO);
      } catch(e) {}
      
      console.log('📋 FacturaID generado: ' + nuevoId + ' (nuevo día, díasDiff=' + diasDiff + ')');
      return nuevoId;
    }
  }
  
  // Fallback: sin referencia, usar la fecha actual (primera vez)
  // Aquí SÍ calculamos el prefijo porque no tenemos referencia
  const fechaBase = new Date(1900, 0, 1);
  const diffMs = now - fechaBase;
  const diffDias = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  const fechaSerial = diffDias + 2;
  const prefijoHoy = Math.floor(fechaSerial).toString().padStart(5, '0');
  const sufijo = "00001";
  const nuevoId = prefijoHoy + sufijo;
  
  try {
    localStorage.setItem('posmovil_ultima_factura_referencia', nuevoId);
    localStorage.setItem('posmovil_ultimo_prefijo', prefijoHoy);
    localStorage.setItem('posmovil_ultimo_sufijo', '1');
    localStorage.setItem('posmovil_fecha_referencia', hoyISO);
  } catch(e) {}
  
  console.log('📋 FacturaID generado (sin referencia): ' + nuevoId);
  return nuevoId;
}
// ============================================
// 9. UTILIDADES
// ============================================
// Timeout global para limpiar mensajes automáticos
var mensajeTimeoutId = null;

function mostrarMensaje(texto, tipo, duracion = 3000) {
  // Limpiar timeout anterior si existe
  if (mensajeTimeoutId) {
    clearTimeout(mensajeTimeoutId);
    mensajeTimeoutId = null;
  }

  mensajeDiv.textContent = texto;
  mensajeDiv.className = `mensaje ${tipo}`;
  mensajeDiv.classList.remove("hidden");

  // Si duracion es -1, el mensaje permanece hasta que se llame mostrarMensaje de nuevo
  // Si duracion es 0, el mensaje se muestra pero nunca se oculta automáticamente
  if (duracion > 0) {
    mensajeTimeoutId = setTimeout(() => {
      mensajeDiv.classList.add("hidden");
      mensajeTimeoutId = null;
    }, duracion);
  }
}

window.toggleCarrito = () => {
  const ahora = Date.now();
  if (ahora - ultimoToggleCarritoTs < 220) return;
  ultimoToggleCarritoTs = ahora;

  const contenido = document.getElementById("carritoContenido");
  const icono = document.getElementById("carritoIcono");
  contenido.classList.toggle("abierto");
  icono.classList.toggle("abierto");
};

window.cerrarCarrito = () => {
  const contenido = document.getElementById("carritoContenido");
  const icono = document.getElementById("carritoIcono");
  contenido.classList.remove("abierto");
  icono.classList.remove("abierto");
};

if (modalCancelar) {
  modalCancelar.addEventListener("click", () => cerrarModalVuelto(false));
}

if (modalConfirmar) {
  modalConfirmar.addEventListener("click", () => cerrarModalVuelto(true));
}

if (modalVuelto) {
  modalVuelto.addEventListener("click", (e) => {
    if (e.target === modalVuelto) cerrarModalVuelto(false);
  });
}

document.addEventListener("keydown", (e) => {
  if (!modalVuelto || modalVuelto.classList.contains("hidden")) return;
  if (e.key === "Escape") cerrarModalVuelto(false);
  if (e.key === "Enter") cerrarModalVuelto(true);
});

// ============================================
// 10. FILTRO DE BÚSQUEDA
// ============================================
function reapplySearchFilter() {
  const termino = searchInput.value.toLowerCase().trim();
  if (termino) {
    ultimoTermino = "";
    aplicarFiltro(termino);
  }
}

function toggleClearButton() {
  if (searchInput.value.length > 0) {
    searchClear.classList.remove("hidden");
  } else {
    searchClear.classList.add("hidden");
  }
}

function limpiarBusqueda() {
  searchInput.value = "";
  toggleClearButton();
  renderizarProductos(productos);
}

if (searchClear) {
  searchClear.addEventListener("click", limpiarBusqueda);
}

const FILTRO_DEBOUNCE_MS = 140;
let filtroTimer = null;
let filtroRaf = 0;
let ultimoTermino = "";

function aplicarFiltro(termino) {
  if (termino === ultimoTermino) return;
  ultimoTermino = termino;

  const filtrados = productos.filter(
    (p) =>
      (p.producto || p.nombre || "").toLowerCase().includes(termino) ||
      (p.codigo || "").toString().includes(termino),
  );

  renderizarProductos(filtrados);
}

searchInput.addEventListener("input", (e) => {
  const termino = e.target.value.toLowerCase().trim();
  toggleClearButton();

  if (filtroTimer) {
    clearTimeout(filtroTimer);
  }

  if (!productosContainer.classList.contains("filtrando")) {
    productosContainer.classList.add("filtrando");
  }

  filtroTimer = setTimeout(() => {
    if (filtroRaf) {
      cancelAnimationFrame(filtroRaf);
    }

    filtroRaf = requestAnimationFrame(() => {
      aplicarFiltro(termino);
      setTimeout(() => {
        productosContainer.classList.remove("filtrando");
      }, 160);
    });
  }, FILTRO_DEBOUNCE_MS);
});

// ============================================
// 11. EVENT LISTENERS
// ============================================
efectivoInput.addEventListener("input", calcularPago);
transferenciaInput.addEventListener("input", calcularPago);

// Limpiar campo al hacer foco si tiene el valor por defecto "0"
efectivoInput.addEventListener("focus", function() {
  if (parsearNumero(this.value) === 0) {
    this.value = "";
  }
});
transferenciaInput.addEventListener("focus", function() {
  if (parsearNumero(this.value) === 0) {
    this.value = "";
  }
});

// Restaurar "0" si el campo queda vacío al perder foco
efectivoInput.addEventListener("blur", function() {
  if (this.value === "") {
    this.value = "0";
  } else {
    this.value = formatearNumero(parsearNumero(this.value));
  }
});
transferenciaInput.addEventListener("blur", function() {
  if (this.value === "") {
    this.value = "0";
  } else {
    this.value = formatearNumero(parsearNumero(this.value));
  }
});

btnPagar.addEventListener("click", finalizarVenta);

// ============================================
// INICIALIZACIÓN
// ============================================
document.addEventListener("DOMContentLoaded", async function() {
  console.log("🚀 App iniciada");
  
  // ===== SISTEMA DE ACTIVACIÓN =====
  // Verificar si el dispositivo está activado
  var activado = false;
  if (DB.isActivated) {
    activado = await DB.isActivated();
  }
  console.log("🔐 [Activacion] Dispositivo activado:", activado);

  var proteccionBloqueo = document.getElementById("proteccionBloqueo");

  if (!activado) {
    // Mostrar pantalla de bloqueo
    if (proteccionBloqueo) {
      proteccionBloqueo.classList.remove("hidden");
    }

    // Generar y mostrar el Device ID
    if (DB.generateDeviceId) {
      var deviceId = await DB.generateDeviceId();
      var idEl = document.getElementById("proteccionDeviceId");
      if (idEl) idEl.textContent = deviceId;
    }

    // NO inicializar el resto de la app
    return;
  } else {
    // Ocultar pantalla de bloqueo
    if (proteccionBloqueo) {
      proteccionBloqueo.classList.add("hidden");
    }
  }
  
  // Inicializar SQLite solo en APK
  if (window.Capacitor && window.Capacitor.isNativePlatform()) {
    dbInicializado = await DB.initDatabase();
    console.log("📦 SQLite inicializado:", dbInicializado);
    if (DB.getStorageMode) {
      console.log("📦 Modo de almacenamiento:", DB.getStorageMode());
    }
    
    // Cargar servidor cacheado de sesiones anteriores
    if (dbInicializado && DB.obtenerServidorCacheado) {
      var urlCacheada = await DB.obtenerServidorCacheado();
      if (urlCacheada) {
        window.SERVER_URL = urlCacheada;
        serverUrlCacheado = urlCacheada;
        console.log("📦 Servidor cacheado cargado:", urlCacheada);
        mostrarMensaje("✅ Servidor: " + urlCacheada, "exito", 2000);
      }
    }
    
    // Cargar indicador de pendientes
    if (dbInicializado) {
      await actualizarIndicadorSync();
    }

    // Limpieza de ventas antiguas (fire-and-forget, no bloquea UI)
    if (dbInicializado && DB.limpiarVentasAntiguas) {
      DB.limpiarVentasAntiguas().catch(function(e) {
        console.warn("⚠️ Error en limpieza automática:", e);
      });
    }
  }
  
  await actualizarPanelEstado();
  await cargarProductos();
  actualizarVistaCarrito();
});

function probarModalManual() {
  mostrarModalVuelto(1000, 1100, 100);
}

// ============================================
// MOSTRAR RESUMEN DEL DÍA (offline desde SQLite)
// ============================================
window.mostrarResumen = async function(fechaISO) {
  var modalResumen = document.getElementById("modalResumen");
  if (!modalResumen) return;

  // Mostrar modal
  modalResumen.classList.remove("hidden");

  // Actualizar título con la fecha seleccionada
  var tituloEl = document.getElementById("modalResumenTitulo");
  if (tituloEl) {
    var fechaTexto = fechaISO || new Date().toLocaleDateString('es-ES');
    tituloEl.textContent = "📊 Resumen - " + fechaTexto;
  }

  var totalEl = document.getElementById("resumenTotal");
  var efectivoEl = document.getElementById("resumenEfectivo");
  var transferenciaEl = document.getElementById("resumenTransferencia");
  var listaEl = document.getElementById("listaProductosResumen");

  if (totalEl) totalEl.textContent = "...";
  if (efectivoEl) efectivoEl.textContent = "...";
  if (transferenciaEl) transferenciaEl.textContent = "...";
  if (listaEl) listaEl.innerHTML = '<li class="info">Cargando resumen...</li>';

  // Intentar primero desde SQLite (offline)
  if (dbInicializado && DB.getResumenOffline) {
    console.log("📦 Cargando resumen desde SQLite (offline)");
    
    try {
      var data = await DB.getResumenOffline(fechaISO);
      
      actualizarResumenUI(data);
      return;
    } catch (error) {
      console.error("❌ Error desde SQLite:", error);
    }
  }

  // Fallback al servidor online
  try {
    console.log("🌐 Cargando resumen desde servidor");
    var response = await fetch(obtenerUrlServidor() + "/api/resumen", {
      headers: {
        "x-session-token": window.SESSION_TOKEN || "",
      },
    });

    if (!response.ok) {
      throw new Error("HTTP " + response.status);
    }

    var data = await response.json();
    actualizarResumenUI(data);
  } catch (error) {
    console.error("❌ Error cargando resumen:", error);
    if (listaEl) listaEl.innerHTML = '<li class="error">Error al cargar resumen</li>';
  }
};

// ============================================
// ACTUALIZAR UI DEL RESUMEN
// ============================================
function actualizarResumenUI(data) {
  var totalEl = document.getElementById("resumenTotal");
  var efectivoEl = document.getElementById("resumenEfectivo");
  var transferenciaEl = document.getElementById("resumenTransferencia");
  var listaEl = document.getElementById("listaProductosResumen");

  if (totalEl) totalEl.textContent = formatearMoneda(data.totalIngresado || 0);
  if (efectivoEl) efectivoEl.textContent = formatearMoneda(data.efectivo || 0);
  if (transferenciaEl) transferenciaEl.textContent = formatearMoneda(data.transferencia || 0);

  if (listaEl) {
    if (!data.productosVendidos || data.productosVendidos.length === 0) {
      listaEl.innerHTML = '<li class="info">No hay ventas hoy</li>';
    } else {
      var html = "";
      for (var i = 0; i < data.productosVendidos.length; i++) {
        var p = data.productosVendidos[i];
        html += '<li>';
        html += '<span class="producto-nombre">' + p.nombre + '</span>';
        html += '<span class="producto-cantidad">' + formatearNumero(p.cantidad || 0) + 'u</span>';
        html += '<span class="producto-total">' + formatearMoneda(p.total || 0) + '</span>';
        html += '</li>';
      }
      listaEl.innerHTML = html;
    }
  }
}

// ============================================
// CERRAR MODAL RESUMEN
// ============================================
function cerrarModalResumen() {
  var modalResumen = document.getElementById("modalResumen");
  if (modalResumen) {
    modalResumen.classList.add("hidden");
  }
}

// Agregar event listeners para el modal de resumen
document.addEventListener("DOMContentLoaded", function() {
  var cerrarBtn = document.getElementById("modalCerrarResumen");
  if (cerrarBtn) {
    cerrarBtn.addEventListener("click", cerrarModalResumen);
  }

  var modalR = document.getElementById("modalResumen");
  if (modalR) {
    modalR.addEventListener("click", function(e) {
      if (e.target === modalR) cerrarModalResumen();
    });
  }

  // NUEVO: Event listener para selector de fecha
  var inputFechaResumen = document.getElementById("inputFechaResumen");
  if (inputFechaResumen) {
    inputFechaResumen.addEventListener("change", function(e) {
      var fechaSeleccionada = e.target.value;  // Formato YYYY-MM-DD
      if (fechaSeleccionada) {
        window.mostrarResumen(fechaSeleccionada);
      }
    });
  }
});

// ============================================
// HISTORIAL DE VENTAS
// ============================================

// Mostrar historial de ventas (agrupado por factura)
window.mostrarHistorialVentas = async function(fechaISO) {
  var modalHistorial = document.getElementById("modalHistorial");
  if (!modalHistorial) return;
  
  // Mostrar modal
  modalHistorial.classList.remove("hidden");
  
  // Actualizar fecha por defecto (hoy)
  var inputFecha = document.getElementById("inputFechaHistorial");
  if (inputFecha && !inputFecha.value) {
    var hoy = new Date();
    var hoyISO = hoy.getFullYear() + '-' + 
               String(hoy.getMonth()+1).padStart(2,'0') + '-' + 
               String(hoy.getDate()).padStart(2,'0');
    inputFecha.value = hoyISO;
  }
  
  // Cargar historial
  await cargarHistorial(inputFecha ? inputFecha.value : fechaISO);
};

// Cargar y mostrar el historial
async function cargarHistorial(fechaISO) {
  var contenedor = document.getElementById("contenidoHistorial");
  if (!contenedor) return;
  
  contenedor.innerHTML = '<div class="loading">Cargando historial...</div>';
  
  try {
    if (!DB.getVentasAgrupadasPorFactura) {
      contenedor.innerHTML = '<div class="error">Función no disponible</div>';
      return;
    }
    
    var facturas = await DB.getVentasAgrupadasPorFactura(fechaISO);
    
    if (facturas.length === 0) {
      contenedor.innerHTML = '<div class="info">No hay ventas para esta fecha</div>';
      return;
    }
    
    // Renderizar historial
    var html = "";
    
    for (var i = 0; i < facturas.length; i++) {
      var f = facturas[i];
      var esSynced = f.synced === 1;
      var estadoClass = esSynced ? "synced" : "pendiente";
      var estadoTexto = esSynced ? "✅ Sincronizado" : "⚠️ Pendiente";
      
      html += '<div class="historial-factura ' + estadoClass + '">';
      html += '<div class="historial-factura-header">';
      html += '<div class="historial-factura-info">';
      html += '<span class="historial-factura-id">Factura: ' + f.facturaId + '</span>';
      html += '<span class="historial-factura-fecha">' + (f.fechaHora || "") + '</span>';
      html += '</div>';
      html += '<div class="historial-factura-estado">';
      html += '<span class="estado-badge ' + estadoClass + '">' + estadoTexto + '</span>';
      
      // Botón deshacer (solo si NO está sincronizado)
      if (!esSynced) {
        html += '<button class="btn-deshacer" onclick="deshacerVentaConfirmar(\'' + f.facturaId.replace(/'/g, "\\'") + '\')" title="Deshacer venta">↩️ Deshacer</button>';
      } else {
        html += '<button class="btn-deshacer" disabled title="No se puede deshacer una venta sincronizada">↩️ Deshacer</button>';
      }
      
      html += '</div>';
      html += '</div>';
      
      // Lista de productos
      html += '<div class="historial-productos">';
      html += '<table class="historial-tabla">';
      html += '<thead><tr><th>Producto</th><th>Cant</th><th>P. Unit</th><th>Subtotal</th></tr></thead>';
      html += '<tbody>';
      
      for (var j = 0; j < f.productos.length; j++) {
        var p = f.productos[j];
        html += '<tr>';
        html += '<td>' + (p.nombre || "") + '</td>';
        html += '<td class="text-center">' + formatearNumero(p.cantidad) + '</td>';
        html += '<td class="text-right">' + formatearMoneda(p.precio) + '</td>';
        html += '<td class="text-right">' + formatearMoneda(p.subtotal) + '</td>';
        html += '</tr>';
      }
      
      html += '</tbody></table>';
      html += '</div>'; // cierra historial-productos
      
      // Totales
      html += '<div class="historial-factura-total">';
      html += '<span>Total: ' + formatearMoneda(f.total) + '</span>';
      html += '<span>';
      if (f.efectivo > 0) html += 'Ef: ' + formatearMoneda(f.efectivo) + ' ';
      if (f.transferencia > 0) html += 'Trans: ' + formatearMoneda(f.transferencia);
      html += '</span>';
      html += '</div>';
      
      html += '</div>'; // cierra historial-factura
    }
    
    contenedor.innerHTML = html;
    
  } catch (error) {
    console.error("❌ Error cargando historial:", error);
    contenedor.innerHTML = '<div class="error">Error al cargar historial</div>';
  }
}

// Confirmar y deshacer venta
window.deshacerVentaConfirmar = async function(facturaId) {
  if (!confirm("¿Estás seguro de deshacer esta venta?\n\nSe revertirán los cambios en el stock.")) {
    return;
  }
  
  try {
    mostrarMensaje("Deshaciendo venta...", "info");
    
    if (!DB.deshacerVenta) {
      mostrarMensaje("Función no disponible", "error");
      return;
    }
    
    var resultado = await DB.deshacerVenta(facturaId);
    
    if (resultado.success) {
      mostrarMensaje("✅ Venta deshecha correctamente", "exito", 3000);
      
      // Recargar productos para reflejar cambios en stock
      await cargarProductos();
      
      // Recargar historial
      var inputFecha = document.getElementById("inputFechaHistorial");
      if (inputFecha) {
        await cargarHistorial(inputFecha.value);
      }
      
      // Actualizar panel de estado
      await actualizarPanelEstado();
      
    } else {
      mostrarMensaje("❌ " + (resultado.error || "Error deshaciendo venta"), "error");
    }
    
  } catch (error) {
    console.error("❌ Error deshaciendo venta:", error);
    mostrarMensaje("Error: " + error.message, "error");
  }
};

// Cerrar modal historial
function cerrarModalHistorial() {
  var modalHistorial = document.getElementById("modalHistorial");
  if (modalHistorial) {
    modalHistorial.classList.add("hidden");
  }
}

// Event listeners para historial
document.addEventListener("DOMContentLoaded", function() {
  // Botón cerrar
  var cerrarBtn = document.getElementById("modalCerrarHistorial");
  if (cerrarBtn) {
    cerrarBtn.addEventListener("click", cerrarModalHistorial);
  }
  
  // Cerrar al hacer clic fuera
  var modalH = document.getElementById("modalHistorial");
  if (modalH) {
    modalH.addEventListener("click", function(e) {
      if (e.target === modalH) cerrarModalHistorial();
    });
  }
  
  // Selector de fecha para historial
  var inputFechaHistorial = document.getElementById("inputFechaHistorial");
  if (inputFechaHistorial) {
    inputFechaHistorial.addEventListener("change", function(e) {
      var fechaSeleccionada = e.target.value;
      if (fechaSeleccionada) {
        cargarHistorial(fechaSeleccionada);
      }
    });
  }
});

// ============================================
// SINCRONIZAR VENTAS OFFLINE + CARGAR PRODUCTOS
// ============================================
window.sincronizar = async function() {
  if (!DB.sincronizarVentas) {
    logSyncAPK("Función de sync de ventas no disponible", "error");
    mostrarMensaje("Función sync no disponible", "error");
    return;
  }

  logSyncAPK("Inicio de sincronización manual");

  var btnSync = document.getElementById("btnSync");
  var syncIcon = document.getElementById("syncIcon");
  var syncCount = document.getElementById("syncCount");

  // Auto-detectar servidor si no hay uno configurado
  if (!window.SERVER_URL) {
    console.log("🔍 No hay servidor, auto-detectando...");
    logSyncAPK("No hay servidor configurado, iniciando auto-detección");
    
    // Mostrar modal de progreso
    var modalSync = document.getElementById("modalSyncProgress");
    var syncStatusText = document.getElementById("syncStatusText");
    var btnCancelSync = document.getElementById("btnCancelSync");
    var btnManualIP = document.getElementById("btnManualIP");
    
    if (modalSync) {
      modalSync.classList.remove("hidden");
      syncStatusText.textContent = "Escaneando la red...";
    }

    // Variables de control
    var syncCancelled = false;
    var manualIPSet = false;

    // Manejar botón Cancelar
    if (btnCancelSync) {
      btnCancelSync.onclick = function() {
        syncCancelled = true;
        cancelarEscaneo = true; // Cancelar escaneo en database.js
        if (modalSync) modalSync.classList.add("hidden");
        logSyncAPK("Sincronización cancelada por usuario", "warning");
        mostrarMensaje("Sincronización cancelada", "info", 2000);
      };
    }

    // Manejar botón IP Manual
    if (btnManualIP) {
      btnManualIP.onclick = function() {
        var ipManual = prompt("Introduce la IP del servidor (ej. 10.225.81.54):", "192.168.1.");
        if (ipManual) {
          window.SERVER_URL = "http://" + ipManual + ":3000";
          serverUrlCacheado = window.SERVER_URL;
          // Guardar en cache si está disponible
          if (DB.guardarServidorCacheado) {
            DB.guardarServidorCacheado(window.SERVER_URL);
          }
          manualIPSet = true;
          syncCancelled = true; // Detener escaneo
          cancelarEscaneo = true;
          if (modalSync) modalSync.classList.add("hidden");
          logSyncAPK("Servidor configurado manualmente: " + window.SERVER_URL);
          mostrarMensaje("✅ IP manual: " + window.SERVER_URL, "exito", 2000);
        }
      };
    }

    if (DB.descubrirServidor) {
      var servidorEncontrado = await DB.descubrirServidor();
      
      // Si se canceló o se usó IP manual, salir si no hay IP válida
      if (syncCancelled && !manualIPSet) {
        logSyncAPK("Escaneo cancelado sin IP manual", "warning");
        return; // Cancelado sin IP manual
      }
      
      if (servidorEncontrado && !manualIPSet) {
        window.SERVER_URL = servidorEncontrado;
        serverUrlCacheado = servidorEncontrado;
        console.log("✅ Servidor encontrado:", servidorEncontrado);
        logSyncAPK("Servidor encontrado: " + servidorEncontrado);
        mostrarMensaje("✅ Servidor: " + servidorEncontrado, "exito", 3000);
      } else if (!manualIPSet && !servidorEncontrado) {
        console.log("❌ No se encontró servidor");
        if (modalSync) modalSync.classList.add("hidden");
        logSyncAPK("No se encontró servidor en la red", "error");
        mostrarMensaje("❌ No se encontró el servidor. Si usa HotSpot: active el servidor y verifique que el teléfono también esté conectado a WiFi (no solo la laptop).", "error", 8000);
        return;
      }
    }

    // Ocultar modal si sigue visible
    if (modalSync) modalSync.classList.add("hidden");
  }

  var pendientes = await DB.contarVentasPendientes();
  logSyncAPK("Pendientes de ventas antes de sync: " + pendientes);
  if (pendientes === 0) {
    mostrarMensaje("Sincronizando con servidor...", "info");
  }

  // Capturar URL ANTES de cualquier operación
  var urlServidor = window.SERVER_URL;

  // Siempre intentar sincronizar aunque haya 0 pendientes
  try {
    if (btnSync) btnSync.disabled = true;
    if (syncIcon) syncIcon.textContent = "⏳";
    var mensajeSync = "";
    var overallSuccess = true;
    var resultadoVentas = { success: false, sincronizadas: 0, error: "No ejecutado" };
    function agregarResumenOk(texto) {
      if (!texto) return;
      if (!mensajeSync) {
        mensajeSync = "✅ " + texto;
      } else {
        mensajeSync += " y " + texto;
      }
    }

    // 1) Sincronizar entrada de productos SIEMPRE
    if (DB.sincronizarEntradaProductos) {
      try {
        logSyncAPK("Sincronizando entradas de productos...");
        var resultadoEntradaProductos = await DB.sincronizarEntradaProductos(urlServidor);
        if (resultadoEntradaProductos.success && resultadoEntradaProductos.sincronizadas > 0) {
          agregarResumenOk(resultadoEntradaProductos.sincronizadas + " entradas de productos sincronizadas");
          logSyncAPK("Entradas sincronizadas: " + resultadoEntradaProductos.sincronizadas);
        } else if (!resultadoEntradaProductos.success) {
          overallSuccess = false;
          mensajeSync = "⚠️ Error entrada productos: " + resultadoEntradaProductos.error;
          logSyncAPK("Error sincronizando entradas: " + resultadoEntradaProductos.error, "error");
        }
      } catch (e) {
        overallSuccess = false;
        console.error("❌ Error sincronizando entrada de productos:", e);
        mensajeSync = "⚠️ Excepción entrada productos: " + e.message;
        logSyncAPK("Excepción sincronizando entradas: " + e.message, "error");
      }
    }

    // 2) Sincronizar abastecimientos SIEMPRE (después de entradas y antes de ventas/mermas)
    if (DB.sincronizarAbastecer) {
      try {
        logSyncAPK("Sincronizando abastecimientos...");
        var resultadoAbastecer = await DB.sincronizarAbastecer(urlServidor);
        if (resultadoAbastecer.success && resultadoAbastecer.sincronizadas > 0) {
          agregarResumenOk(resultadoAbastecer.sincronizadas + " abastecimientos");
          logSyncAPK("Abastecimientos sincronizados: " + resultadoAbastecer.sincronizadas);
        } else if (!resultadoAbastecer.success) {
          overallSuccess = false;
          mensajeSync += " | Error abastecimientos: " + resultadoAbastecer.error;
          logSyncAPK("Error sincronizando abastecimientos: " + resultadoAbastecer.error, "error");
        }
      } catch (eAbastecer) {
        overallSuccess = false;
        console.error("❌ Error sincronizando abastecimientos:", eAbastecer);
        mensajeSync += " | Error abastecimientos: " + eAbastecer.message;
        logSyncAPK("Excepción sincronizando abastecimientos: " + eAbastecer.message, "error");
      }
    }

    // 3) Ventas (después de entradas y abastecimientos)
    try {
      logSyncAPK("Sincronizando ventas...");
      resultadoVentas = await DB.sincronizarVentas();

      if (resultadoVentas.success) {
        guardarUltimaSync(new Date());
        agregarResumenOk(resultadoVentas.sincronizadas + " ventas");
        logSyncAPK("Ventas sincronizadas: " + resultadoVentas.sincronizadas);
      } else {
        overallSuccess = false;
        mensajeSync += " | Error ventas: " + (resultadoVentas.error || "No se pudo sincronizar ventas");
        logSyncAPK("Error sincronizando ventas: " + (resultadoVentas.error || "desconocido"), "error");
      }
    } catch (eVentas) {
      overallSuccess = false;
      mensajeSync += " | Excepción ventas: " + (eVentas.message || "desconocido");
      logSyncAPK("Excepción sincronizando ventas: " + (eVentas.message || "desconocido"), "error");
    }

    // 4) Sincronizar mermas SIEMPRE
    if (DB.sincronizarMermas) {
      try {
        logSyncAPK("Sincronizando mermas...");
        var resultadoMermas = await DB.sincronizarMermas(urlServidor);
        if (resultadoMermas.success && resultadoMermas.sincronizadas > 0) {
          agregarResumenOk(resultadoMermas.sincronizadas + " mermas");
          logSyncAPK("Mermas sincronizadas: " + resultadoMermas.sincronizadas);
        } else if (!resultadoMermas.success) {
          overallSuccess = false;
          mensajeSync += " | Error mermas: " + resultadoMermas.error;
          logSyncAPK("Error sincronizando mermas: " + resultadoMermas.error, "error");
        }
      } catch (e) {
        overallSuccess = false;
        console.error("❌ Error sincronizando mermas:", e);
        mensajeSync += " | Error mermas: " + e.message;
        logSyncAPK("Excepción sincronizando mermas: " + e.message, "error");
      }
    }

    // Modo offline depende del estado global
    modoOffline = !overallSuccess;

    if (!mensajeSync) {
      mensajeSync = overallSuccess ? "✅ Sincronización completada" : "⚠️ Sincronización finalizada con advertencias";
    }

    logSyncAPK("Sincronización finalizada");
    mostrarMensaje(mensajeSync, overallSuccess ? "exito" : "warning", 4000);
    
    // Actualizar UI
    await cargarProductos();
    await actualizarIndicadorSync();
    await actualizarPanelEstado();
    
  } catch (error) {
    console.error("❌ Error sync:", error);
    modoOffline = true;
    logSyncAPK("Error general de sincronización: " + (error.message || "desconocido"), "error");
    mostrarMensaje(error.message || "Error al sincronizar", "warning", 3000);
    await cargarProductos();
    await actualizarIndicadorSync();
    await actualizarPanelEstado();
  } finally {
    if (btnSync) btnSync.disabled = false;
    if (syncIcon) syncIcon.textContent = "🔄";
    logSyncAPK("Botón de sincronizar reactivado");
  }
};

// ============================================
// ACTUALIZAR INDICADOR DE PENDIENTES (ventas + mermas + entradas + abastecimientos)
// ============================================
async function actualizarIndicadorSync() {
  if (!DB.contarVentasPendientes && !DB.contarMermasPendientes && !DB.contarEntradaProductosPendientes) return 0;

  var totalPendientes = await contarPendientesTotales();

  var syncCount = document.getElementById("syncCount");
  if (syncCount) {
    syncCount.textContent = totalPendientes > 0 ? String(totalPendientes) : "";
  }

  if (estadoPendientesEl) {
    estadoPendientesEl.textContent = String(totalPendientes);
  }

  return totalPendientes;
}

// ============================================
// MENÚ DESPLEGABLE
// ============================================
function toggleMenu() {
  const menu = document.getElementById("menuDesplegable");
  if (menu) menu.classList.toggle("hidden");
}

// Cerrar menú si tocás afuera
document.addEventListener("click", function (e) {
  const menu = document.getElementById("menuDesplegable");
  const btn = document.getElementById("btnMenu");
  if (menu && !menu.contains(e.target) && !btn.contains(e.target)) {
    menu.classList.add("hidden");
  }
});

// ============================================
// NUEVO PRODUCTO - MOSTRAR FORMULARIO
// ============================================
window.mostrarFormularioProducto = async function() {
  const modal = document.getElementById("modalNuevoProducto");
  const menu = document.getElementById("menuDesplegable");
  
  if (!modal) return;
  
  // Cerrar menú desplegable
  if (menu) menu.classList.add("hidden");
  
  // Generar código automático
  try {
    if (DB.getUltimoCodigoProducto) {
      var codigo = await DB.getUltimoCodigoProducto();
      document.getElementById("npCodigo").value = codigo;
    }
  } catch (e) {
    console.error("Error generando código:", e);
    document.getElementById("npCodigo").value = "Pr_00001";
  }
  
  // Poner fecha actual
  var hoy = new Date();
  var fechaISO = hoy.getFullYear() + '-' + 
               String(hoy.getMonth()+1).padStart(2,'0') + '-' + 
               String(hoy.getDate()).padStart(2,'0');
  document.getElementById("npFecha").value = fechaISO;
  
  // Limpiar campos anteriores
  document.getElementById("npNombre").value = "";
  document.getElementById("npCantidad").value = "0";
  document.getElementById("npPrecioVenta").value = "0";
  document.getElementById("npPrecioCosto").value = "0";
  
  // Mostrar modal
  modal.classList.remove("hidden");
  document.getElementById("npNombre").focus();
};

// ============================================
// NUEVO PRODUCTO - CERRAR FORMULARIO
// ============================================
function cerrarFormularioProducto() {
  const modal = document.getElementById("modalNuevoProducto");
  if (modal) modal.classList.add("hidden");
}

// ============================================
// EVENT LISTENERS PARA FORMULARIO DE NUEVO PRODUCTO
// ============================================
document.addEventListener("DOMContentLoaded", function() {
  // Botón Cancelar
  var btnCancelar = document.getElementById("btnCancelarProducto");
  if (btnCancelar) {
    btnCancelar.addEventListener("click", cerrarFormularioProducto);
  }
  
  // Botón Guardar
  var btnGuardar = document.getElementById("btnGuardarProducto");
  if (btnGuardar) {
    btnGuardar.addEventListener("click", async function() {
      var codigo = document.getElementById("npCodigo").value.trim();
      var nombre = document.getElementById("npNombre").value.trim();
      var cantidad = parsearNumero(document.getElementById("npCantidad").value);
      var precioVenta = parsearNumero(document.getElementById("npPrecioVenta").value);
      
      if (!nombre) {
        mostrarMensaje("El nombre es requerido", "error");
        return;
      }
      
      if (precioVenta <= 0) {
        mostrarMensaje("El precio debe ser mayor a 0", "error");
        return;
      }
      
      try {
        var guardado = await DB.guardarEntradaProductoCompleto({
          codigo: codigo,
          nombre: nombre,
          precioVenta: precioVenta,
          precioCosto: parsearNumero(document.getElementById("npPrecioCosto").value),
          cantidad: cantidad
        });
        
        if (guardado) {
          modoOffline = true;
          productos.push({
            codigo: codigo,
            producto: nombre,
            precio: precioVenta,
            disponibilidad: cantidad
          });
          renderizarProductos(productos);
          cerrarFormularioProducto();
          mostrarMensaje("✅ " + codigo + " agregado", "exito", 3000);
          actualizarPanelEstado();
        } else {
          mostrarMensaje("Error guardando", "error");
        }
      } catch (e) {
        mostrarMensaje("Error: " + e.message, "error");
      }
    });
  }
  
  // Cerrar modal al hacer clic fuera
  var modalNuevo = document.getElementById("modalNuevoProducto");
  if (modalNuevo) {
    modalNuevo.addEventListener("click", function(e) {
      if (e.target === modalNuevo) cerrarFormularioProducto();
    });
  }
  
  // Cerrar modal con ESC
  document.addEventListener("keydown", function(e) {
    if (!modalNuevo) return;
    if (!modalNuevo.classList.contains("hidden") && e.key === "Escape") {
      cerrarFormularioProducto();
    }
  });

  // Limpiar campos numéricos al hacer foco si tienen "0"
  var npCantidad = document.getElementById("npCantidad");
  var npPrecioVenta = document.getElementById("npPrecioVenta");
  var npPrecioCosto = document.getElementById("npPrecioCosto");

  [npCantidad, npPrecioVenta, npPrecioCosto].forEach(function(input) {
    if (!input) return;
    
    input.addEventListener("focus", function() {
      if (parsearNumero(this.value) === 0) {
        this.value = "";
      }
    });

    input.addEventListener("blur", function() {
      if (this.value === "") {
        this.value = "0";
      } else {
        this.value = formatearNumero(parsearNumero(this.value));
      }
    });
  });
});

// ============================================
// SISTEMA DE ACTIVACIÓN DE DISPOSITIVOS
// ============================================

// Mostrar modal de protección (desde el menú)
window.mostrarProteccion = async function() {
  var menu = document.getElementById("menuDesplegable");
  if (menu) menu.classList.add("hidden");

  var modal = document.getElementById("modalProteccion");
  if (!modal) return;

  // Cargar Device ID en el modal
  if (DB.generateDeviceId) {
    var deviceId = await DB.generateDeviceId();
    var idEl = document.getElementById("modalProteccionDeviceId");
    if (idEl) idEl.value = deviceId;
  }

  // Limpiar campos
  var claveInput = document.getElementById("modalProteccionNuevaClave");
  if (claveInput) claveInput.value = "";
  var errorEl = document.getElementById("modalProteccionError");
  if (errorEl) errorEl.classList.add("hidden");
  var cargandoEl = document.getElementById("modalProteccionCargando");
  if (cargandoEl) cargandoEl.classList.add("hidden");

  modal.classList.remove("hidden");
};

// Cerrar modal de protección
window.cerrarModalProteccion = function() {
  var modal = document.getElementById("modalProteccion");
  if (modal) modal.classList.add("hidden");
};

// Copiar Device ID desde la pantalla de bloqueo
window.copiarDeviceId = async function() {
  var idEl = document.getElementById("proteccionDeviceId");
  if (!idEl) return;

  try {
    await navigator.clipboard.writeText(idEl.textContent);
    mostrarMensaje("✅ ID copiado al portapapeles", "exito", 2000);
  } catch (e) {
    // Fallback: seleccionar el texto
    var range = document.createRange();
    range.selectNodeContents(idEl);
    var selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    mostrarMensaje("📋 Seleccioná el ID y copialo", "info", 2000);
  }
};

// Copiar Device ID desde el modal
window.copiarModalDeviceId = async function() {
  var inputEl = document.getElementById("modalProteccionDeviceId");
  if (!inputEl) return;

  try {
    await navigator.clipboard.writeText(inputEl.value);
    mostrarMensaje("✅ ID copiado al portapapeles", "exito", 2000);
  } catch (e) {
    inputEl.select();
    mostrarMensaje("📋 Seleccioná el ID y copialo", "info", 2000);
  }
};

// ============================================
// EJECUTAR ACTIVACIÓN (desde pantalla de bloqueo)
// ============================================
window.ejecutarActivacion = async function() {
  var claveInput = document.getElementById("proteccionClaveInput");
  var btnActivar = document.getElementById("btnActivar");
  var errorEl = document.getElementById("proteccionError");
  var cargandoEl = document.getElementById("proteccionCargando");
  var exitoEl = document.getElementById("proteccionExito");
  var idEl = document.getElementById("proteccionDeviceId");

  if (!claveInput || !btnActivar || !errorEl || !cargandoEl || !exitoEl) return;

  var clave = claveInput.value.trim();
  if (!clave) {
    errorEl.textContent = "✕ Ingresá la clave de activación";
    errorEl.classList.remove("hidden");
    return;
  }

  // Ocultar error previo
  errorEl.classList.add("hidden");

  // Mostrar cargando
  btnActivar.disabled = true;
  cargandoEl.classList.remove("hidden");

  try {
    var valida = await DB.verificarClave(clave);

    if (valida) {
      // Guardar activación permanente
      await DB.guardarActivacion();

      // Mostrar éxito
      cargandoEl.classList.add("hidden");
      exitoEl.classList.remove("hidden");

      // Recargar la app después de 1.5s (ocultar bloqueo, mostrar app)
      setTimeout(function() {
        location.reload();
      }, 1500);
    } else {
      cargandoEl.classList.add("hidden");
      btnActivar.disabled = false;
      errorEl.textContent = "✕ Clave incorrecta. Verificá con el administrador.";
      errorEl.classList.remove("hidden");
      claveInput.value = "";
      claveInput.focus();
    }
  } catch (e) {
    console.error("❌ [Activacion] Error:", e);
    cargandoEl.classList.add("hidden");
    btnActivar.disabled = false;
    errorEl.textContent = "✕ Error al verificar: " + e.message;
    errorEl.classList.remove("hidden");
  }
};

// ============================================
// EJECUTAR RE-ACTIVACIÓN (desde modal, ya activado)
// ============================================
window.ejecutarReactivacion = async function() {
  var claveInput = document.getElementById("modalProteccionNuevaClave");
  var errorEl = document.getElementById("modalProteccionError");
  var cargandoEl = document.getElementById("modalProteccionCargando");
  var btnReactivar = document.getElementById("modalProteccionActivar");

  if (!claveInput || !errorEl || !cargandoEl || !btnReactivar) return;

  var clave = claveInput.value.trim();
  if (!clave) {
    errorEl.textContent = "✕ Ingresá la clave de activación";
    errorEl.classList.remove("hidden");
    return;
  }

  errorEl.classList.add("hidden");
  btnReactivar.disabled = true;
  cargandoEl.classList.remove("hidden");

  try {
    var valida = await DB.verificarClave(clave);

    if (valida) {
      await DB.guardarActivacion();
      cargandoEl.classList.add("hidden");
      mostrarMensaje("✅ Dispositivo re-activado correctamente", "exito", 3000);
      cerrarModalProteccion();
    } else {
      cargandoEl.classList.add("hidden");
      btnReactivar.disabled = false;
      errorEl.textContent = "✕ Clave incorrecta";
      errorEl.classList.remove("hidden");
      claveInput.value = "";
      claveInput.focus();
    }
  } catch (e) {
    console.error("❌ [Activacion] Error:", e);
    cargandoEl.classList.add("hidden");
    btnReactivar.disabled = false;
    errorEl.textContent = "✕ Error: " + e.message;
    errorEl.classList.remove("hidden");
  }
};

// ============================================
// AUTO-FORMATO DE CLAVE (agrega guiones automáticos)
// ============================================
document.addEventListener("DOMContentLoaded", function() {
  var inputs = document.querySelectorAll(".proteccion-input");
  inputs.forEach(function(input) {
    input.addEventListener("input", function(e) {
      var val = this.value.replace(/[^0-9A-Fa-f]/g, "").toUpperCase();
      var formatted = "";
      for (var i = 0; i < val.length && i < 16; i++) {
        if (i > 0 && i % 4 === 0) formatted += "-";
        formatted += val[i];
      }
      this.value = formatted;
    });
  });
});

// Cerrar modal de protección al hacer clic fuera
document.addEventListener("DOMContentLoaded", function() {
  var modalProteccion = document.getElementById("modalProteccion");
  if (modalProteccion) {
    modalProteccion.addEventListener("click", function(e) {
      if (e.target === modalProteccion) cerrarModalProteccion();
    });
  }

  // Enter en el input de clave de bloqueo → activar
  var bloqueoInput = document.getElementById("proteccionClaveInput");
  if (bloqueoInput) {
    bloqueoInput.addEventListener("keydown", function(e) {
      if (e.key === "Enter") {
        e.preventDefault();
        ejecutarActivacion();
      }
    });
  }

  // Enter en el input de clave del modal → reactivar
  var modalInput = document.getElementById("modalProteccionNuevaClave");
  if (modalInput) {
    modalInput.addEventListener("keydown", function(e) {
      if (e.key === "Enter") {
        e.preventDefault();
        ejecutarReactivacion();
      }
    });
  }
});


