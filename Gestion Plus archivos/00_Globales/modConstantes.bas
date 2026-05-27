Attribute VB_Name = "modConstantes"
' ============================================================================
' MODULO: modConstantes
' PROPOSITO: Centraliza TODAS las constantes del proyecto (nombres de hojas,
' tablas, contraseñas, rutas) para facilitar el mantenimiento.
' ============================================================================
' INSTRUCCIONES:
'   - NO hardcodees nombres de hojas, tablas o contraseñas en el codigo.
'   - Usa estas constantes SIEMPRE.
'   - Si cambia una contraseña o nombre, solo se cambia aqui.
' ============================================================================

Option Explicit

' --- CONTRASEÑAS -----------------------------------------------------------
' ATENCION: Estas contraseñas estan expuestas en el codigo fuente.
'           En una version futura, mover a un sistema mas seguro.
Public Const PASS_HOJA      As String = "theSheetProtectpassisthe011rd"
Public Const PASS_LIBRO     As String = "theBookProtectpassisthe010nd"
Public Const PASS_ADMIN     As String = "859674"

' --- NOMBRES DE HOJAS ------------------------------------------------------
Public Const HOJA_ALMACEN       As String = "Gestión de Almacén"
Public Const HOJA_VENTAS        As String = "Historial de ventas"
Public Const HOJA_GASTOS        As String = "Historial de Gastos"
Public Const HOJA_FACTURAS      As String = "Registro de facturas"
Public Const HOJA_INGREDIENTES  As String = "Gestión de Ingredientes"
Public Const HOJA_TRABAJADORES  As String = "Gestión de Trabajadores"
Public Const HOJA_RECETAS       As String = "Elaboración de Productos"
Public Const HOJA_INGR_RECETA   As String = "Ingredientes"
Public Const HOJA_FECHAS        As String = "Fechas"
Public Const HOJA_DASHBOARD     As String = "Dashboard"
Public Const HOJA_HIDDEN        As String = "Hidden"
Public Const HOJA_BLANCO        As String = "Blanco"

' --- NOMBRES DE TABLAS (ListObjects) ---------------------------------------
Public Const TBL_INVENTARIO         As String = "INVENTARIO"
Public Const TBL_VENTAS             As String = "HISTORIALVENTAS"
Public Const TBL_GASTOS             As String = "Gastos"
Public Const TBL_FACTURAS           As String = "HISTORIALVENTAS4"
Public Const TBL_INGREDIENTES       As String = "INGREDIENTES"
Public Const TBL_TRABAJADORES       As String = "trabajadores"
Public Const TBL_RECETAS            As String = "Recetas"
Public Const TBL_INGREDIENTES_RECETA As String = "Ingredientes_Receta"
Public Const TBL_FECHAS_UNICAS      As String = "FechasUnicas"
Public Const TBL_PRODUCTOS_UNICOS   As String = "Productos_Unicos"

' --- RUTAS Y ARCHIVOS ------------------------------------------------------
Public Const CARPETA_WEBAPP        As String = "webapp-beta\"
Public Const CARPETA_RESPALDOS     As String = "\Respaldos\"
Public Const ARCHIVO_DATOS         As String = "data\datos.xlsx"
Public Const ARCHIVO_VENTAS_LOG    As String = "data\ventas.log"
Public Const CARPETA_FLAGS         As String = "flags\"
Public Const ARCHIVO_FLAG_ACTIVO   As String = "servidor_activo.flag"
Public Const ARCHIVO_FLAG_DETENER  As String = "detener.flag"

' --- COLUMNAS DE TABLAS (mapeo) --------------------------------------------
' TBL_INVENTARIO: Gestión de Almacén
'   1=Código, 2=Nombre, 3=Fecha, 4=Cant.Inicial, 5=Cant.Actual,
'   6=PrecioVenta, 7=PrecioCosto, 8=Fondo, 9=IngresoEsp, 10=GananciaEsp
Public Const COL_INV_CODIGO      As Long = 1
Public Const COL_INV_NOMBRE      As Long = 2
Public Const COL_INV_FECHA       As Long = 3
Public Const COL_INV_CANT_INI    As Long = 4
Public Const COL_INV_CANT_ACT    As Long = 5
Public Const COL_INV_PRECIO_V    As Long = 6
Public Const COL_INV_PRECIO_C    As Long = 7
Public Const COL_INV_FONDO       As Long = 8
Public Const COL_INV_INGRESO     As Long = 9
Public Const COL_INV_GANANCIA    As Long = 10

' TBL_GASTOS: Historial de Gastos
'   1=Fecha, 2=Categoría, 3=Descripción, 4=Monto
Public Const COL_GAS_FECHA       As Long = 1
Public Const COL_GAS_CATEGORIA   As Long = 2
Public Const COL_GAS_DESCRIPCION As Long = 3
Public Const COL_GAS_MONTO       As Long = 4

' TBL_INGREDIENTES: Gestión de Ingredientes
'   1=Nombre, 2=Código/Alterno, 3=Stock, 4=UnidadBase, 5=Precio, 6=Fondo, 7=Monto
Public Const COL_ING_NOMBRE       As Long = 1
Public Const COL_ING_STOCK        As Long = 3
Public Const COL_ING_UNIDAD_BASE  As Long = 4
Public Const COL_ING_PRECIO       As Long = 5
Public Const COL_ING_FONDO        As Long = 6

' TBL_RECETAS: Elaboración de Productos
'   1=Receta, 2=Lote, 3=Inversión/Lote, 4=Precio/Lote,
'   5=Ganancia/Lote, 6=Inversión/Unidad, 7=Precio/Unidad, 8=Ganancia/Unidad
Public Const COL_REC_NOMBRE       As Long = 1
Public Const COL_REC_LOTE         As Long = 2
Public Const COL_REC_INV_LOTE     As Long = 3
Public Const COL_REC_PRECIO_LOTE  As Long = 4
Public Const COL_REC_GAN_LOTE     As Long = 5
Public Const COL_REC_INV_UNIDAD   As Long = 6
Public Const COL_REC_PRECIO_UNID  As Long = 7
Public Const COL_REC_GAN_UNIDAD   As Long = 8

' TBL_INGREDIENTES_RECETA: Ingredientes
'   1=Receta, 2=Ingrediente, 3=Cantidad, 4=Unidad
Public Const COL_IR_RECETA       As Long = 1
Public Const COL_IR_INGREDIENTE  As Long = 2
Public Const COL_IR_CANTIDAD     As Long = 3
Public Const COL_IR_UNIDAD       As Long = 4

' TBL_TRABAJADORES: Gestión de Trabajadores
'   1=Nombre, 2=Cargo, 3=Porcentaje Salario, 4=Salario Mínimo
Public Const COL_TRAB_NOMBRE      As Long = 1
Public Const COL_TRAB_CARGO       As Long = 2
Public Const COL_TRAB_PORCENTAJE  As Long = 3
Public Const COL_TRAB_SAL_MIN     As Long = 4

' --- CATEGORÍAS DE GASTOS --------------------------------------------------
Public Const CAT_FIJO        As String = "Fijo"
Public Const CAT_VARIABLE    As String = "Variable"
Public Const CAT_INVERSION   As String = "Inversión"
Public Const CAT_MERMA       As String = "Merma"
