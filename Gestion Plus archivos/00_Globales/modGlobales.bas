Attribute VB_Name = "modGlobales"
' ============================================================================
' MÓDULO: modGlobales
' PROPÓSITO: Único punto de declaración para todas las variables públicas
' del proyecto Gestion Plus.
' ============================================================================
' INSTRUCCIONES:
'   - No declarar variables Public fuera de este módulo.
'   - Si necesitas una nueva variable global, agrégala aquí con un comentario
'     que explique su propósito y dónde se usa.
' ============================================================================

Option Explicit

' --- NAVEGACIÓN Y SELECCIÓN ------------------------------------------------
Public fila_sel          As Integer   ' Fila seleccionada en tabla (Trabajador_Mod)
Public fila_a_modificar  As Integer   ' Fila activa para modificar (calling, formularios)
Public seleccion_destino_calendario As Integer  ' Controla qué TextBox recibe la fecha del calendario

' --- TRABAJADORES ----------------------------------------------------------
Public cargo             As String    ' Cargo del trabajador al pagar salario
Public pago              As Double    ' Monto a pagar al trabajador

' --- REPETICIÓN EN FORMULARIOS --------------------------------------------
Public repeticion        As Boolean   ' Repetición en facturación
Public repeticion2       As Boolean   ' Repetición en ingrediente nuevo
Public repeticion3       As Boolean   ' Repetición en ingrediente modificar

' --- VALIDACIÓN DE CAMPOS VACÍOS -----------------------------------------
Public contador          As Integer   ' Contador de campos vacíos (ingresar_form)
Public contador2         As Integer   ' Contador de campos vacíos (ingrediente nuevo)
Public contador3         As Integer   ' Contador de campos vacíos (ingrediente existente)

' --- RECETAS --------------------------------------------------------------
Public modoModificacion  As Boolean   ' True = modificando receta existente
Public recetaAModificar  As String    ' Nombre de la receta a modificar

' --- GENERACIÓN DE CÓDIGOS ------------------------------------------------
Public Codigo_Producto   As String    ' Último código generado (formato "Pr_00001")
