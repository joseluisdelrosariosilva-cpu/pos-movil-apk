# webapp-beta — POS Móvil

Sistema de Punto de Venta (POS) diseñado para usarse desde un dispositivo móvil en red local. La aplicación web se comunica con un archivo Excel que actúa como intermediario de datos entre el frontend y un sistema de gestión en Excel con macros VBA.

> **Flujo:** Móvil → API Express → Excel (`datos.xlsx`) → Macros de Excel procesan las ventas.

---

## Características

- 📱 **Frontend mobile-first** con tema oscuro, búsqueda en tiempo real y carrito flotante
- 📱 **APK Android nativa** con Capacitor 8 y SQLite local offline
- 📡 **Sincronización batch** — un solo request sincroniza productos, abastecimientos, ventas, mermas, gastos y elaboraciones con remapeo automático de códigos
- 🧪 **Elaboración por lotes** — producir stock desde recetas con registro local y sincronización
- 🔒 **Sesión única** vinculada a IP con token y expiración renovable (30 min)
- 📊 **Persistencia en Excel** — sin base de datos tradicional
- 🚦 **Control por flags** — macros de Excel pueden iniciar/detener el servidor
- ⚡ **ES Modules** con Express 5

---

## Requisitos

- **Node.js** v18+
- **Windows** (se usa `xlsx-populate` y macros VBA de Excel)
- **Red local** para acceso desde dispositivo móvil
- **`data/datos.xlsx`** con las hojas `Productos`, `Pendientes` y `Recetas`

---

## Instalación

### Desarrollo local

```bash
cd webapp-beta
git clone <repo-url>
npm install
```

### APK Android (con Capacitor)

```bash
# Sincronizar y compilar APK
npm run build:android
```

También podés usar `compilar.bat` directamente (configura `JAVA_HOME` y `ANDROID_HOME` automáticamente).

---

## Uso

### Iniciar el servidor

```bash
# Desarrollo (con auto-reload)
npm run dev

# Producción
npm start

# Android
npm run build:android   # Compilar APK
npm run open:android    # Abrir en Android Studio
```

### Acceder desde el móvil

```
http://{IP-del-servidor}:3000
```

### Control desde Excel (macros VBA)

| Acción | Mecanismo |
|--------|-----------|
| Iniciar servidor | Macro ejecuta `npm start` vía shell |
| Verificar estado | Macro lee `flags/servidor_activo.flag` |
| Detener servidor | Macro crea `flags/detener.flag` |

---

## Configuración

El proyecto usa variables de entorno. Para desarrollo funciona con los valores por defecto.

```bash
# Copiar template
copy .env.example .env
```

| Variable | Default | Descripción |
|----------|---------|-------------|
| `PORT` | `3000` | Puerto del servidor |
| `SESSION_EXPIRATION_MS` | `1800000` (30 min) | Tiempo de expiración de sesión |

---

## Estructura del Proyecto

```
webapp-beta/
├── src/
│   ├── server.js                 # Entry point (Express)
│   ├── routes/
│   │   ├── auth.routes.js        # Autenticación (sesión/token)
│   │   ├── pos.routes.js         # Productos, Recetas, Ingredientes
│   │   ├── ventas.routes.js      # Ventas
│   │   ├── resumen.routes.js     # Resumen / dashboard
│   │   └── sync-completo.routes.js  # Sincronización batch (APK → Excel)
│   ├── controllers/
│   │   ├── pos.controller.js     # Lee productos, recetas e ingredientes desde Excel
│   │   ├── ventas.controller.js  # Escribe ventas en Excel
│   │   ├── resumen.controller.js
│   │   └── sync-completo.controller.js  # Procesa todo en un solo request batch
│   ├── middlewares/
│   │   └── auth.middleware.js    # Validación de token de sesión
│   ├── utils/
│   │   ├── excelHelper.js        # CRUD con Excel (xlsx-populate + mutex)
│   │   ├── server-control.js     # Gestión de flags (activo/detener)
│   │   ├── session.util.js       # Sesión en memoria (IP + token + expiración)
│   │   ├── token.util.js         # Generación y validación de tokens
│   │   └── date.util.js          # Helpers de fecha (hora local)
│   ├── public/
│   │   ├── index.html            # UI principal del POS
│   │   ├── css/style.css         # Estilos (tema oscuro, responsive)
│   │   └── js/
│   │       ├── app.js            # Lógica del frontend
│   │       └── database.js       # Manejo de datos (SQLite local + localStorage)
│   │                            # Tablas: productos, ventas_pending, mermas_pending,
│   │                            # entrada_productos_pending, abastecer_pending,
│   │                            # gastos_pending, elaboraciones_pending, config,
│   │                            # recetas, ingredientes
│   └── views/
│       └── ocupado.html          # Página "POS Ocupado"
├── android/                      # Proyecto Android nativo (Capacitor)
├── data/
│   └── datos.xlsx                # Base de datos intermediaria
├── dist/                         # Builds empaquetados
├── flags/                        # Control del servidor
├── capacitor.config.json         # Configuración de Capacitor
├── compilar.bat                  # Script para compilar APK
├── .env.example                  # Template de variables
├── .gitignore
└── package.json
```

---

## API Endpoints

| Endpoint | Método | Auth | Descripción |
|----------|--------|------|-------------|
| `GET /` | GET | No | Página principal (inyecta `SESSION_TOKEN`) |
| `GET /api/estado-publico` | GET | No | Health-check público para descubrimiento de servidor desde la APK |
| `GET /api/test` | GET | No | Prueba de servidor (público) |
| `GET /api/estado` | GET | Sí | Estado de flags del servidor |
| `GET /api/estado-sesion` | GET | Sí | Detalle de sesión activa |
| `GET /api/estado-publico` | GET | No | Indica si hay sesión activa (público) |
| `POST /api/cerrar-sesion` | POST | Sí | Cierra la sesión activa |
| `GET /api/mutex` | GET | Sí | Diagnóstico del mutex de Excel |
| `POST /api/mutex/reiniciar` | POST | Sí | Reinicia la cola del mutex si está trabado |
| `GET /api/productos` | GET | Flexible* | Lista productos desde Excel |
| `GET /api/recetas` | GET | Flexible* | Lista recetas desde Excel (Nombre, CantLote, PrecioVenta, PrecioCosto) |
| `GET /api/ingredientes` | GET | Flexible* | Lista ingredientes de recetas desde Excel (Ingrediente, Cantidad, Unidad, Receta) |
| `POST /api/ventas` | POST | Sí | Registra una venta en Excel |
| `POST /api/sync/completo` | POST | Sí | Sincronización batch: productos, abastecimientos, ventas, mermas, gastos y elaboraciones en un solo request con remapeo automático de códigos. Devuelve `ultimaFactura` (último ID de factura en Pendientes) para que la APK mantenga su referencia offline |
| `GET /api/resumen?desde=YYYY-MM-DD&hasta=YYYY-MM-DD` | GET | Sí | Resumen de ventas por rango de fechas (si no se envía desde, usa el día actual) |

Todas las peticiones autenticadas envían el token en el header `x-session-token`.

> **Nota:** Las rutas marcadas como *Flexible* (`/productos`, `/recetas`, `/ingredientes`, `/sync/completo`) permiten acceso sin token de sesión para compatibilidad con la APK sin sesión previa.

---

## Flujo de una Venta

```
1. Usuario abre navegador → http://{IP}:3000
2. Servidor genera token de sesión (primera visita)
3. Frontend carga productos desde Excel
4. Usuario busca producto (filtro en tiempo real)
5. Usuario toca "+" → se agrega al carrito
6. Usuario abre carrito, ajusta cantidades, ingresa pago
7. Usuario toca "Pagar" → POST /api/ventas
8. Servidor valida pago, verifica stock, escribe en Excel
9. Frontend limpia carrito y recarga productos
10. Macros de Excel leen "Pendientes", procesan ventas, marcan Procesado=true
```

---

## Sincronización Batch (APK → Excel)

La APK Android acumula datos offline y los sincroniza con el servidor en un solo request batch vía `POST /api/sync/completo`. El orden de procesamiento es determinista y atómico:

```
1. Productos nuevos → hoja "Productos" + "Entrada"
   ↪ Si hay conflicto de código, se remapea automáticamente (Pr_00001, Pr_00002...)
2. Abastecimientos → hoja "Abastecimiento"
   ↪ Usa el mapa de remapeo de códigos del paso 1
   ↪ Suma al stock en hoja "Productos"
3. Ventas → hoja "Pendientes"
   ↪ Usa el mapa de remapeo de códigos
   ↪ Si el sufijo del FacturaID es ≤ al máximo existente para ese mismo prefijo (día),
     se remapea automáticamente al siguiente sufijo disponible (preserva el prefijo original)
   ↪ Descuenta del stock
4. Mermas → hoja "Merma"
   ↪ Usa el mapa de remapeo de códigos
   ↪ Descuenta del stock
5. Gastos → hoja "Gastos" (se crea si no existe)
6. Elaboraciones → hoja "Elaboracion" (se crea si no existe)
```

Todo dentro de un mutex (`conExcelLock`): si **algo** falla, **nada** se guarda. El servidor devuelve:

- `remap.codigos` — mapa de códigos de producto remapeados (old → new)
- `remap.facturas` — mapa de facturas remapeadas por conflicto (old → new)
- `ultimaFactura` — el ID de la última factura en Pendientes (un string de 10 dígitos), para que la APK mantenga su referencia de IDs offline

### Referencia de FacturaID offline

La APK necesita generar FacturaIDs incluso sin conexión. Para evitar duplicados:

1. Después de cada sync exitoso, la APK guarda `ultimaFactura` del servidor como `posmovil_ultima_factura_referencia` en localStorage
2. También actualiza `posmovil_ultima_sync` con el timestamp del sync (incluso si no había datos pendientes)
3. Al generar un ID offline, la función `generarFacturaId()` usa esa referencia para calcular el prefijo (fecha serial) y el sufijo consecutivo, incrementando desde el último conocido
4. Si no hay referencia (primera vez), arranca desde sufijo `00001` para el día actual

---

## Arquitectura Offline (APK Android)

La APK usa una **arquitectura de doble persistencia**: los datos se guardan en SQLite local y simultáneamente en `localStorage` como fallback.

```
┌─────────────────────────────────────────────────┐
│                  Frontend (APK)                  │
│                                                  │
│   ┌──────────┐      ┌────────────────────────┐   │
│   │ IndexedDB │◄────►│      database.js       │   │
│   │ (sql.js)  │      │ ┌──────────────────┐  │   │
│   └──────────┘      │ │   crearWrapper    │  │   │
│                      │ │   SQLite (Cap 8) │  │   │
│   ┌──────────┐      │ ├──────────────────┤  │   │
│   │localStor.│◄────►│ │activarModoLocal()│  │   │
│   └──────────┘      │ └──────────────────┘  │   │
│                      └────────────────────────┘   │
└─────────────────────────────────────────────────┘
```

### Inicialización (offline-first)

La APK **no toca la red al iniciar**. Siempre abre en modo offline cargando datos desde SQLite local:

1. **Capacitor 8** via `@capacitor-community/sqlite` → SQLite nativo
   - `createConnection()` → `open()` → `crearTablasSQLite()`
   - Fallo → `activarModoLocal()` (vuelca todo a localStorage)
2. **Capacitor <8 / Cordova** via `window.SQLite.createDatabase()`
   - Fallo → `activarModoLocal()`
3. **Sin Capacitor** (navegador web) → localStorage directo
4. Se cargan los **servidores conocidos** desde SQLite (lista persistente de hasta 10 URLs de sesiones anteriores)
5. Se cargan productos desde SQLite (`cargarProductos(true)`) — **sin fetch al servidor**
6. Solo cuando el usuario toca **"🔄 Sincronizar"** se intenta conectar al servidor

### Descubrimiento de servidor en red local

Cuando se presiona el botón Sync, la APK descubre automáticamente el servidor Express en la red local:

1. **URL conocida** — Si hay una `SERVER_URL` de la sesión actual, la prueba con fetch a `/api/estado-publico` (timeout 2s). Si falla, se limpia para no reintentarla.
2. **Servidores conocidos** — Prueba la lista persistente (hasta 10 URLs guardadas en SQLite/localStorage). Las auto-detectadas tienen prioridad sobre las manuales.
3. **Escaneo de red** — Si ningún conocido responde, detecta la IP local del dispositivo (RTCPeerConnection / Network Information API) y escanea el rango auto-detectado más hasta 6 rangos comunes `/24` en lotes de 3 concurrentes, con batches de 25 IPs paralelas y timeout de 300ms por request.
4. **IP Manual** — El usuario puede ingresar una IP manualmente desde el modal, con validación de formato y preview en vivo de la URL.

Los servidores encontrados se guardan automáticamente en la lista persistente como tipo `"auto"`. Las IPs ingresadas manualmente se guardan como `"manual"` al final de la lista.

El modal de sincronización muestra:
- Progreso del escaneo en tiempo real
- **Error inline** si no se encuentra servidor (dentro del modal, no como toast detrás)
- Vista de **IP manual** con preview de URL en vivo y validación de formato

### Esquema SQLite

```sql
-- Productos, ventas, gastos, etc. se persisten como JSON en cada tabla
CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT, timestamp INTEGER);
CREATE TABLE IF NOT EXISTS productos (data TEXT);
CREATE TABLE IF NOT EXISTS ventas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
CREATE TABLE IF NOT EXISTS mermas_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
CREATE TABLE IF NOT EXISTS entrada_productos_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
CREATE TABLE IF NOT EXISTS abastecer_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
CREATE TABLE IF NOT EXISTS gastos_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
CREATE TABLE IF NOT EXISTS elaboraciones_pending (id INTEGER PRIMARY KEY AUTOINCREMENT, synced INTEGER DEFAULT 0, data TEXT);
```

### Persistencia robusta

- Cada operación escribe en **SQLite + localStorage** simultáneamente
- Si SQLite falla, `localStorage` sigue funcionando como respaldo
- Al leer, prioriza SQLite; si falla, cae a localStorage
- El batch de sincronización se construye desde SQLite (o localStorage si SQLite no está disponible)

---

## Política de Fechas

**Regla de oro:** TODAS las fechas se generan en **hora local del dispositivo**, sin conversión UTC.

| Componente | Mecanismo | Ejemplo |
|------------|-----------|---------|
| App (frontend JS) | `fechaLocalISO()` — construye el string con `getFullYear/getMonth/getDate` locales | `2026-05-31T23:00:00.000` |
| Servidor (Express) | `date.util.js` — `fechaLocalISO()` | mismo formato |
| Excel (Entrada) | `formatearFechaDesdeISO()` extrae la fecha de creación real del producto | `31/05/2026` |

> ⚠️ **No usar `new Date().toISOString()`** para fechas de negocio. `toISOString()` convierte a UTC, lo que en zonas con offset negativo (ej: Argentina UTC-3) desplaza la fecha un día al operar después de las 21:00.

---

## Estructura de `datos.xlsx`

### Hoja "Productos"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `codigo` | Identificador del producto |
| B | `producto` | Nombre del producto |
| C | `disponibilidad` | Stock disponible |

### Hoja "Recetas"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `nombre` | Nombre de la receta |
| B | `cant_lote` | Cantidad producida por lote |
| C | `precio_venta` | Precio de venta por unidad |
| D | `precio_costo` | Costo por unidad (Inversión/Unidad) |
| F | `ingrediente` | Ingrediente de la receta (1 fila por ingrediente) |
| G | `cantidad` | Cantidad del ingrediente |
| H | `unidad` | Unidad de medida del ingrediente |
| I | `receta` | Nombre de la receta a la que pertenece |

### Hoja "Elaboracion"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `NombreReceta` | Nombre de la receta elaborada |
| B | `Lotes` | Cantidad de lotes producidos |

> Se escribe desde la APK al usar la función **🧪 Elaborar** del menú.

> Esta hoja se escribe desde VBA (`SincronizarRecetasEnDatos`) al activar el servidor,
> y se sincroniza a la app al presionar "🔄 Sincronizar" desde el menú.
> `PrecioCosto` mapea a la columna `Inversión/Unidad` (`COL_REC_INV_UNIDAD`) de la tabla Recetas en "Elaboración de Productos".
> Los ingredientes (columnas F:I) se leen de la tabla `TBL_INGREDIENTES_RECETA` en la hoja "Ingredientes" del libro Gestion Plus.

### Hoja "Pendientes"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `FacturaID` | Formato: `[fecha serial 5 dígitos][consecutivo 5 dígitos]` |
| B | `FechaHora` | ISO local: `2026-04-06T14:11:43.886` |
| C | `CodigoProducto` | Código del producto |
| D | `Nombre` | Nombre del producto |
| E | `CantVendida` | Cantidad vendida |
| F | `PrecioUnitario` | Precio unitario |
| G | `Subtotal` | Cantidad × Precio |
| H | `Efectivo` | Monto pagado en efectivo |
| I | `Transferencia` | Monto pagado por transferencia |
| J | `Procesado` | `false` por defecto (macros lo ponen en `true`) |

### Hoja "Entrada"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `codigo` | Código del producto (puede ser remapeado si hay conflicto) |
| B | `producto` | Nombre del producto |
| C | `fecha` | Fecha de creación en formato DD/MM/YYYY |
| D | `stock_inicial` | Stock inicial |
| E | `precio_venta` | Precio de venta |
| F | `precio_costo` | Precio de costo |

### Hoja "Abastecimiento"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `codigo` | Código del producto a reabastecer |
| B | `cantidad` | Cantidad agregada al stock |
| C | `fecha_hora` | Fecha y hora del abastecimiento |

### Hoja "Merma"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `codigo` | Código del producto |
| B | `nombre` | Nombre del producto |
| C | `cantidad` | Cantidad de merma |

### Hoja "Gastos"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `Fecha` | Fecha del gasto |
| B | `Descripcion` | Concepto del gasto |
| C | `Monto` | Monto del gasto |

---

## Autenticación

- **Modelo:** Single-session, vinculada a IP, con token y expiración.
- **Inicio:** Primer visitante a `/` recibe un token único inyectado en el HTML.
- **Uso:** Cada petición API envía `x-session-token` en el header.
- **Expiración:** 30 minutos de inactividad (se renueva con cada petición válida).
- **IP binding:** Si otro dispositivo intenta acceder, ve la página "POS Ocupado".
- **Cierre:** `POST /api/cerrar-sesion` o apagar el servidor.

---

## Distribución

### Opción A: GitHub (desarrolladores)

```bash
git clone <repo-url>
cd webapp-beta
npm install
npm start
```

### Opción B: APK Android (usuarios finales)

Compilar con Capacitor:
- Genera una APK nativa de Android en `android/app/build/outputs/apk/debug/`.
- La APK se conecta al servidor Express en la red local.
- Usa SQLite local (`@capacitor-community/sqlite`) como caché offline.

---

## Tecnologías

| Tecnología | Versión | Propósito |
|------------|---------|-----------|
| Node.js | v18+ | Runtime |
| Express | 5.2.1 | Servidor HTTP / API |
| xlsx-populate | 1.21.0 | Lectura/escritura de Excel |
| Capacitor | 8.x | App Android nativa |
| @capacitor-community/sqlite | 8.x | SQLite local offline |
| dotenv | latest | Variables de entorno |
| nodemon | 3.1.x | Hot reload (dev) |

---

## Licencia

MIT
