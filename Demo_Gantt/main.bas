Attribute VB_Name = "main"
Option Explicit

Sub main()
    Dim gantt As CGantt
    
    Set gantt = New CGantt
    'Zona donde se inicia la representacion
    gantt.fila_inicio = 3
    gantt.col_inicio = 5
    'Columnas donde esta la informacion
    gantt.col_ID = 1
    gantt.col_Desc = 2
    gantt.col_Duracion = 3
    gantt.col_ID_Pre = 4
    
    'carga de datos para cada tarea
    gantt.load_data data
    'dibuja diagrama Gantt
    gantt.draw
    
    Set gantt = Nothing
End Sub
