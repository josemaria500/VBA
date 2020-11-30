VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Tasks As Collection
Private ws As Worksheet
Private pcol_inicio As Long
Private pfila_inicio As Long
Private pcol_ID As Long
Private pcol_Desc As Long
Private pcol_Duracion As Long
Private pcol_ID_Pre As Long
Private p_loaded As Boolean
    
'fila inicio grafico
Public Property Get fila_inicio() As Long
    fila_inicio = pfila_inicio
End Property
Public Property Let fila_inicio(value As Long)
    pfila_inicio = value
End Property
    
'columna inicio grafico
Public Property Get col_inicio() As Long
    col_inicio = pcol_inicio
End Property
Public Property Let col_inicio(value As Long)
    pcol_inicio = value
End Property

'columna posicion datos ID
Public Property Get col_ID() As Long
    col_ID = pcol_ID
End Property
Public Property Let col_ID(value As Long)
    pcol_ID = value
End Property

'columna posicion datos Descripcion
Public Property Get col_Desc() As Long
    col_Desc = pcol_Desc
End Property
Public Property Let col_Desc(value As Long)
    pcol_Desc = value
End Property

'columna posicion datos Duracion
Public Property Get col_Duracion() As Long
    col_Duracion = pcol_Duracion
End Property
Public Property Let col_Duracion(value As Long)
    pcol_Duracion = value
End Property

'columna posicion datos Duracion
Public Property Get col_ID_Pre() As Long
    col_ID_Pre = pcol_ID_Pre
End Property
Public Property Let col_ID_Pre(value As Long)
    pcol_ID_Pre = value
End Property
'Pasamos wsData donde estan los datos
Public Sub load_data(wsData As Worksheet)
    Set ws = wsData
     
    Call crea_coleccion_tareas
    Call calculo_duracion_previa
    Call calcula_left_rectangulo
    'variable para indicar que estan cargados los datos y se puede dibujar el grafico Gantt
    p_loaded = True
End Sub

Sub crea_coleccion_tareas()
    Dim task As CTask
    Dim cell As Range
    Dim rango As Range
    Dim numTask As Long
    
    Set Tasks = New Collection
    Set rango = ws.Range("A" & CStr(fila_inicio), ws.Range("A" & CStr(fila_inicio)).End(xlDown))
    
    For Each cell In rango
        Set task = New CTask
        task.ID = cell.Offset(0, col_ID - 1)
        task.Descripcion = cell.Offset(0, col_Desc - 1)
        task.duracion = cell.Offset(0, col_Duracion - 1)
        task.ID_Precedente = cell.Offset(0, col_ID_Pre - 1)
        task.Top = cell.Top
        task.Height = cell.Height
        task.Color = cell.Interior.Color
        task.Width = calculo_width(task.duracion)
        Tasks.Add task
    Next cell
End Sub
'Conversion del color a RGB
Private Function fRGB(Color) As Variant
    fRGB = RGB(Color Mod 256, Color / 256 Mod 256, Color / 65536 Mod 256)
End Function
 'calcula el ancho del rectangulo. Este metodo no acumula errores en la representacion
Function calculo_width(duracion)
    Dim duracion_ent As Long
    Dim duracion_dec As Double
    
    duracion_ent = Fix(duracion)
    duracion_dec = duracion - duracion_ent
    calculo_width = ws.Cells(fila_inicio, col_inicio + duracion_ent).Left - ws.Cells(fila_inicio, col_inicio).Left _
                            + duracion_dec * ws.Cells(fila_inicio, col_inicio).Width
End Function

'Para posicionar los rectangulos es necesario conocer la duracion total de las todas las tareas que se
' realizan. Con el valor obtenido calculamos depsues el valor Left del rectangulo
Private Sub calculo_duracion_previa()
    Dim task As CTask
    Dim oTask As CTask
    Dim pTask As CTask
    Dim sTask As CTask
    Dim duracion_acumulada As Double
    Dim cero_IDPrecedente As Boolean
    
    For Each oTask In Tasks
        'desciende hasta ID_precedente = 0  y va acumulando duracion, que luego se la asigna a la tarea
        duracion_acumulada = 0
        If oTask.ID_Precedente = 0 Then
            cero_IDPrecedente = True
        Else
            cero_IDPrecedente = False
        End If
        Set sTask = oTask
        Do While Not cero_IDPrecedente
            For Each pTask In Tasks
                If pTask.ID = sTask.ID_Precedente Then
                    duracion_acumulada = duracion_acumulada + pTask.duracion
                    If pTask.ID_Precedente = 0 Then cero_IDPrecedente = True
                    Set sTask = pTask
                    Exit For
                End If
            Next pTask
        Loop
        oTask.DuracionPrevias = duracion_acumulada
        Next oTask
End Sub
 
 'Con el valor de duracion total de las tareas previas calculamos el valor de Left del rectangulo.
 'Este metodo no acumula errores de representación
Private Sub calcula_left_rectangulo()
    Dim duracion_ent As Long
    Dim duracion_dec As Double
    Dim task As CTask
    
    For Each task In Tasks
        duracion_ent = Fix(task.DuracionPrevias)
        duracion_dec = task.DuracionPrevias - duracion_ent
        task.Left = ws.Cells(fila_inicio, col_inicio + duracion_ent).Left + duracion_dec * ws.Cells(fila_inicio, col_inicio).Width
    Next task
End Sub

Sub draw()
    Dim rect As Shape
    Dim oTask As CTask
    
    'chequea si primero a cargado los datos y efectuado los calculos necesarios
    If Not p_loaded Then
        MsgBox "Antes de poder dibujar el grafico debes llamar a:" & vbCrLf & vbCrLf & "Gantt.load_data" & vbCrLf & _
                      vbCrLf & "Revisa el código en el módulo principal", vbExclamation, "Gantt demo"
        End
    End If
    
    'borrar todos los rectagulos
    For Each rect In data.Shapes
        If rect.Top > data.Cells(fila_inicio - 1, 1).Top Then rect.Delete
    Next rect
    
    'dibuja los rectangulos de las tareas
    For Each oTask In Tasks
        Set rect = ws.Shapes.AddShape(msoShapeRectangle, oTask.Left, oTask.Top, oTask.Width, oTask.Height)
        With rect
            .Name = oTask.ID
            .Fill.ForeColor.RGB = fRGB(oTask.Color)
            .Fill.Transparency = 0.2
        End With
    Next oTask
End Sub
