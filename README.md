# webapp-beta — POS Móvil

Sistema de Punto de Venta (POS) diseñado para usarse desde un dispositivo móvil en red local. La aplicación web se comunica con un archivo Excel que actúa como intermediario de datos entre el frontend y un sistema de gestión en Excel con macros VBA.

> **Flujo:** Móvil → API Express → Excel (`datos.xlsx`) → Macros de Excel procesan las ventas.

---

## Características

- 📱 **Frontend mobile-first** con tema oscuro, búsqueda en tiempo real y carrito flotante
- 📱 **APK Android nativa** con Capacitor (SQLite local offline)
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
│   │   ├── pos.routes.js         # Productos, Recetas
│   │   ├── ventas.routes.js      # Ventas
│   │   ├── gastos.routes.js      # Gastos
│   │   ├── abastecimientos.routes.js  # Abastecimientos
│   │   ├── entrada-productos.routes.js # Entrada de productos
│   │   ├── mermas.routes.js      # Mermas
│   │   └── resumen.routes.js     # Resumen / dashboard
│   ├── controllers/
│   │   ├── pos.controller.js     # Lee productos y recetas desde Excel
│   │   ├── ventas.controller.js  # Escribe ventas en Excel
│   │   ├── gastos.controller.js
│   │   ├── abastecimientos.controller.js
│   │   ├── entrada-productos.controller.js
│   │   ├── mermas.controller.js
│   │   └── resumen.controller.js
│   ├── middlewares/
│   │   └── auth.middleware.js    # Validación de token de sesión
│   ├── utils/
│   │   ├── excelHelper.js        # CRUD con Excel (xlsx-populate)
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
│   │                            # gastos_pending, elaboraciones_pending, config, recetas
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
| `GET /api/test` | GET | Sí | Prueba de servidor |
| `GET /api/estado` | GET | Sí | Estado de flags del servidor |
| `GET /api/estado-sesion` | GET | Sí | Detalle de sesión activa |
| `GET /api/estado-publico` | GET | No | Indica si hay sesión activa (público) |
| `POST /api/cerrar-sesion` | POST | Sí | Cierra la sesión activa |
| `GET /api/productos` | GET | Flexible* | Lista productos desde Excel |
| `GET /api/recetas` | GET | Flexible* | Lista recetas desde Excel (Nombre, CantLote, PrecioVenta) |
| `POST /api/ventas` | POST | Sí | Registra una venta en Excel |
| `POST /api/gastos` | POST | Sí | Registra gastos desde APK |
| `POST /api/abastecimientos` | POST | Sí | Sincroniza abastecimientos desde APK |
| `POST /api/entrada-productos` | POST | Sí | Sincroniza productos nuevos desde APK |
| `POST /api/mermas` | POST | Sí | Registra mermas desde APK |
| `GET /api/resumen` | GET | Sí | Resumen de ventas del día |

Todas las peticiones autenticadas envían el token en el header `x-session-token`.

> **Nota:** Las rutas marcadas como *Flexible* (`/productos`, `/recetas`, `/ventas`, `/resumen`, `/sync/completo`) permiten acceso sin token de sesión para compatibilidad con la APK sin sesión previa.

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

### Hoja "Elaboracion"

| Columna | Campo | Descripción |
|---------|-------|-------------|
| A | `NombreReceta` | Nombre de la receta elaborada |
| B | `Lotes` | Cantidad de lotes producidos |

> Se escribe desde la APK al usar la función **🧪 Elaborar** del menú.

> Esta hoja se escribe desde VBA (`SincronizarRecetasEnDatos`) al activar el servidor,
> y se sincroniza a la app al presionar "🔄 Sincronizar" desde el menú.
> `PrecioCosto` mapea a la columna `Inversión/Unidad` (`COL_REC_INV_UNIDAD`) de la tabla Recetas en "Elaboración de Productos".

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
