# Mapa de Migración — Gestion Plus VBA

## Estructura del proyecto

```
Gestion Plus archivos/
├── 00_Globales/          # Constantes y variables globales (único punto)
│   ├── modConstantes.bas  # TODAS las constantes: hojas, tablas, rutas, passwords
│   │                      # [Fase3] Añadidas COL_TRAB_NOMBRE/CARGO/PORCENTAJE/SAL_MIN
│   ├── modGlobales.bas    # TODAS las variables públicas
│   └── modLicencia.bas    # Lógica de licencia (extraído de ThisWorkbook)
├── 01_Utilidades/        # Funciones transversales
│   ├── modConversor.bas   # Conversión de unidades de ingredientes
│   ├── modFormateo.bas    # Formateo de números y fechas
│   ├── modValidacionCampos.bas  # Validación de campos vacíos en formularios
│   ├── modValidacionNumerica.bas # Validación de entrada numérica (KeyPress)
│   └── modSeguridad.bas   # Protección/desprotección de libro y hojas
│                          # [Fase3] Nuevas: VerificarPasswordAdmin, CambiarPasswordAdmin,
│                          #          ObtenerPasswordAdminAlmacenada, MostrarHojasAdministrativas
├── 02_Gestion_BD/        # Operaciones de base de datos
│   ├── modBusquedas.bas   # Búsquedas y verificación de repeticiones
│   │                      # [Fase3] ComprobarRepeticionFacturacion desacoplada (recibe parámetros)
│   │                      # [Fase3] Nuevas: CargarTodosLosProductos, BuscarProductos,
│   │                      #          CargarDescripcionesUnicas (genérica), NavegarAFilaOriginal,
│   │                      #          BuscarGastos, BuscarVentas, BuscarFacturas
│   ├── modEliminarRegistros.bas  # Eliminación de registros
│   ├── modStock.bas       # Actualización de stock y reabastecimiento
│   └── modAgregarRegistros.bas  # Inserción de gastos, fechas, facturas, ventas
├── 03_Inventario/        # Lógica de inventario
│   ├── modMermas.bas      # Cálculo de mermas
│   ├── modCalculosInventario.bas  # Precio UBase, monto inversión, precio ingrediente
│   └── modGeneradorCodigo.bas  # Generación de códigos de producto (Pr_00001)
├── 04_Ventas/            # Facturación y ventas
│   └── modFacturacion.bas # Cálculo de total y generación de ID de factura
│                          # [Fase3] Funciones desacopladas (reciben parámetros en vez de referencias a form)
│                          # [Fase3] Nueva: RegistrarVentaCompleta (productos + factura + stock + gastos)
├── 05_Trabajadores/      # Gestión de trabajadores
│   └── modTrabajadores.bas  # PagarTrabajador, ModificarTrabajador, AgregarTrabajador (+), ModificarTrabajadorEnTabla (+)
├── 06_Respaldos/         # Respaldos y limpieza
│   └── modRespaldos.bas   # Respaldo, respaldo completo, diario, final
├── 07_POS_Movil/         # Servidor y sincronización (extraído de serverControl.bas)
│   ├── modServerControl.bas
│   ├── modImportacion.bas
│   └── modSincronizacion.bas
├── 08_Recetas/           # Recetas y elaboración
│   ├── modRecetas.bas     # CRUD de recetas, búsqueda por ingrediente
│   │                      # [Fase3] Nuevas: AgregarReceta, EliminarReceta, EliminarIngredientesReceta,
│   │                      #          EliminarRecetaYIngredientes, AgregarIngredienteReceta,
│   │                      #          CargarIngredientesReceta (array 2D)
│   └── modElaboracion.bas # Validación stock, descuento, registro, costos
│                          # [Bugfix Fase3] ValidarStockElaboracion y DescontarStockIngredientes
│                          # ahora convierten unidades vía modConversor antes de comparar/descontar.
│                          # RegistrarProductoElaborado acepta código opcional (codigoExterno).
├── 09_Dashboard/         # Dashboard
│   └── modDashboard.bas   # Refresco, verificación expiración, formas, ocultar
├── 10_Calendario/        # Calendario (antes ModuloCalendario)
│   └── modCalendario.bas
│
├── Formularios/          # ORIGINALES — formularios (se mantienen como interfaz)
├── Hojas/                # ORIGINALES — code-behind de hojas (eventos)
│   ├── ThisWorkbook.cls  # (MIGRADO) delega en modLicencia, modSeguridad, modRespaldos
│   ├── Hoja3.cls         # (MIGRADO) delega en modSeguridad
│   ├── Hoja6.cls         # (MIGRADO) delega en modDashboard
│   ├── Hoja9.cls         # (MIGRADO) selection change — mantenido
│   ├── Hoja10.cls        # (MIGRADO) delega en modElaboracion, modConversor
│   ├── Hoja11.cls        # (MIGRADO) selection change — mantenido
│   └── Hoja1-12.cls      # Sin código (excepto las listadas)
└── Modulos/              # 3 ORIGINALES CONSERVADOS (form-callers + UDFs) — resto eliminados (Fase 3)
```

## Correspondencia Original → Reorganizado

### Módulos .bas

| Original | Estado | Reorganizado |
|----------|--------|-------------|
| `calling.bas` | ✅ Conservado | Form-callers UI; `ReabastecerProducto` reasignado a `modStock` |
| `Módulo1.bas` | ✅ Conservado | Form-callers UI + `fila_sel, cargo, pago`; delegaciones reasignadas a modEliminarRegistros/modTrabajadores/modDashboard |
| `descuento_almacen.bas` | ✅ Conservado | UDFs para fórmulas de Excel (`ObtenerMermasProducto`, `actualizar_almacen`) |
| `Módulo2.bas` | ❌ Eliminado | Todo → `modRespaldos` |
| `Módulo3.bas` | ❌ Eliminado | Botón `eliminar_ingrediente` reasignado a `modEliminarRegistros` |
| `ModuloCalendario.bas` | ❌ Eliminado | Todo → `10_Calendario/modCalendario` (forms llaman `modCalendario` directo) |
| `serverControl.bas` | ❌ Eliminado | Botones reasignados a `modImportacion`/`modMermas` |
| `Validar_campos_vacios.bas` | ❌ Eliminado | `modValidacionCampos` desde forms directo |
| `Validar_vacios_2.bas` | ❌ Eliminado | `modValidacionCampos` desde forms directo |
| `Validar_vacios_3.bas` | ❌ Eliminado | `modValidacionCampos` desde forms directo |
| `ID_Factura.bas` | ❌ Eliminado | Sin botones |
| `Caluculo_de_total.bas` | ❌ Eliminado | Sin botones |
| `Conversor.bas` | ❌ Eliminado | Forms llaman `modConversor` directo |
| `Generador_codigo.bas` | ❌ Eliminado | Sin botones |
| `Eliminar_producto_seleccionado.bas` | ❌ Eliminado | Botones reasignados a `modEliminarRegistros` |
| `Calculo_precio_UBase.bas` | ❌ Eliminado | Sin botones |
| `busqueda_factura.bas` | ❌ Eliminado | Sin botones |

### Hojas .cls (code-behind)

| Original | Estado | Delegación |
|----------|--------|-----------|
| `ThisWorkbook.cls` | ✅ Migrado | Eventos mantienen; lógica → `modLicencia`, `modSeguridad`, `modRespaldos` |
| `Hoja3.cls` | ✅ Migrado | `protegerLibro/desprotegerLibro` → `modSeguridad` |
| `Hoja6.cls` | ✅ Migrado | `Actualizar_Formas` → `modDashboard.ActualizarFormasDashboard` |
| `Hoja9.cls` | ✅ Mantenido | SelectionChange puro (solo UI) |
| `Hoja10.cls` | ✅ Migrado | ColorCoding (UI); recálculo → `modElaboracion.ActualizarRecetasPorIngrediente` |
| `Hoja11.cls` | ✅ Mantenido | SelectionChange puro (solo UI) |
| `Hoja1,2,4,5,7,8,12.cls` | ✅ Vacías | Sin código |

### Formularios .frm (Fase 3 — REFACTORIZACIÓN COMPLETA ✅)

Todos los formularios delegan la lógica de negocio en los módulos reorganizados. Solo mantienen eventos de UI (MouseMove, Change, KeyPress), orquestación y estado visual.

| Formulario | Estado | Delegación |
|-----------|--------|-----------|
| `Gasto_Form.frm` | ✅ Refactorizado | `modAgregarRegistros.AgregarGasto`, `.AgregarFecha`, `modValidacionNumerica` |
| `Trabajador.frm` | ✅ Refactorizado | `modTrabajadores.AgregarTrabajador`, `modValidacionNumerica` |
| `Trabajador_Mod.frm` | ✅ Refactorizado | `modTrabajadores.ModificarTrabajadorEnTabla`, constantes `COL_TRAB_*` |
| `ingresar_form.frm` | ✅ Refactorizado | `modValidacionCampos.ValidarVaciosProducto`, `modAgregarRegistros.AgregarProductoUnico/AgregarFecha`, `modGeneradorCodigo` |
| `ingresar_ingr_nuevo_form.frm` | ✅ Refactorizado | `modValidacionCampos.ValidarVaciosIngredienteNuevo`, `modBusquedas.ComprobarRepeticionIngredienteNuevo/Modificar`, `modCalculosInventario.CalcularPrecioUBase` |
| `ingresar_ingr_existente_form.frm` | ✅ Refactorizado | `modValidacionCampos.ValidarVaciosIngredienteExistente`, `modConversor.ConvertirUnidad`, `modRecetas.CargarUnidadesPorIngrediente`, `modCalculosInventario.CalcularPrecioUBase` |
| `Elaborar.frm` | ✅ Refactorizado | `modElaboracion.ValidarStockElaboracion`, `.DescontarStockIngredientes`, `.RegistrarProductoElaborado`, `modRecetas.CargarRecetasEnCombo`, `modGeneradorCodigo`, `modFormateo` |
| `Ingr_Receta.frm` | ✅ Refactorizado | `modRecetas.AgregarReceta/EliminarRecetaYIngredientes/AgregarIngredienteReceta/CargarIngredientesReceta/CargarIngredientesEnCombo/CargarUnidadesPorIngrediente`, `modConversor.ConvertirUnidad`, `modCalculosInventario.ObtenerPrecioPorUnidadBase`, `modValidacionNumerica`, `modFormateo`, `modConstantes` |
| `facturacion.frm` | ✅ Refactorizado | `modFacturacion.CalcularTotalFacturacion/GenerarIdFactura/RegistrarVentaCompleta`, `modBusquedas.BuscarProductos/CargarTodosLosProductos/ComprobarRepeticionFacturacion`, `modStock.RestarCantidadInventario`, `modAgregarRegistros.AgregarGastoInversion/AgregarFecha`, `modCalculosInventario.CalcularMontoInversion`, `modValidacionNumerica`, `modFormateo`, `modConstantes` |
| `Busqueda.frm` | ✅ Refactorizado | `modBusquedas.BuscarGastos/CargarDescripcionesUnicas/NavegarAFilaOriginal`, constantes `HOJA_GASTOS`, `TBL_GASTOS`, `COL_GAS_*` |
| `Busqueda_Ventas.frm` | ✅ Refactorizado | `modBusquedas.BuscarVentas/CargarDescripcionesUnicas/NavegarAFilaOriginal`, constantes `HOJA_VENTAS`, `TBL_VENTAS` |
| `Busqueda_Facturacion.frm` | ✅ Refactorizado | `modBusquedas.BuscarFacturas/CargarDescripcionesUnicas/NavegarAFilaOriginal`, constantes `HOJA_FACTURAS`, `TBL_FACTURAS`. Bugfix: navegaba a "Historial de ventas" en vez de "Registro de facturas" |
| `frmCalendario.frm` | ✅ Refactorizado | Eventos UI + `modCalendario` (en vez de `ModuloCalendario` original) |
| `Login.frm` | ✅ Refactorizado | `modSeguridad.VerificarPasswordAdmin/CambiarPasswordAdmin/MostrarHojasAdministrativas/DesprotegerLibro/ProtegerLibro`, constantes `PASS_ADMIN`, `HOJA_*` |
| `frmPOSMovil.frm` | ✅ Refactorizado | `modServerControl.ServidorEstaActivo/ObtenerIPLocal/DesactivarPOSMovil/ContarVentasPendientes`, `modImportacion.ImportarVentas`. Timer UI se mantiene en el form |

## Convenciones del proyecto

1. **Nombres**: `mod` + dominio (ej: `modStock`, `modFacturacion`)
2. **Constantes**: SIEMPRE desde `modConstantes` — NO hardcodear nombres de hojas/tablas
3. **Variables globales**: SOLO en `modGlobales` — no declarar `Public` en otro módulo
4. **Fachadas**: Los módulos originales ya migrados deben ser solo fachadas que delegan
5. **Hojas .cls**: Solo eventos; toda lógica de negocio en módulos .bas
6. **Formularios .frm**: Idealmente solo UI; lógica delega a módulos
7. **Protección**: Siempre usar `modSeguridad.ProtegerHoja/DesprotegerHoja` y constantes de `modConstantes`
8. **Errores**: On Error Resume Next solo donde sea necesario; On Error GoTo para errores esperados
