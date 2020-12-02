Attribute VB_Name = "main"
Option Explicit

Sub main()
    Dim rng As Range
    Dim cell As Range
    Dim lat As String
    Dim lon As String
    Dim data As String
    Dim parameters() As String
    Dim i As Integer
    Dim col_inicio As Integer
    Dim row_inicio As Integer
    'seleccionamos rango donde estan los datos, para despues recorrerlos
    Set rng = cities.Range("A2", cities.Range("A2").End(xlDown))
    
    row_inicio = 2
    col_inicio = 5
    For Each cell In rng
        'seleccionamos latitud
        lat = CStr(cell.Offset(0, 1))
        'seleccionamos longitud
        lon = CStr(cell.Offset(0, 2))
        'consulta API y obtenemos JSON
        data = api_call(lat, lon)
        'extraemos los datos de JSON
        parameters = get_parameters(data)
        'mostramos los datos
        weather.Cells(row_inicio, col_inicio) = cell
        For i = 0 To 3
            weather.Cells(row_inicio + 4 * (i + 1), col_inicio) = parameters(i)
        Next i
                col_inicio = col_inicio + 3
    Next cell
    weather.Range("B3") = Date
    weather.Range("B4") = Time
End Sub

Private Function api_call(lat As String, lon As String) As String
    'Tools > References:  Microsoft XML v6.0
    Dim apiURL As String
    Dim requestString As String
    Dim endpoint As String
    Dim latitude As String
    Dim longitude As String
    Dim exclude As String
    Dim units As String
    Dim appid As String
    Dim request As MSXML2.ServerXMLHTTP60
    Dim answer As String
    
    apiURL = "https://api.openweathermap.org"
    endpoint = "/data/2.5/onecall?"
    latitude = lat
    longitude = lon
    exclude = "minutely,hourly,daily,alerts"
    units = "metric"
    appid = "your API-KEY here"

    requestString = apiURL & endpoint & "lat=" & latitude & "&lon=" & longitude & "&exclude=" & exclude & "&units=" & units & "&appid=" & appid
   
    Set request = New ServerXMLHTTP60
    request.Open "GET", requestString, False
    request.send
    api_call = request.responseText
End Function
 
 Function get_parameters(json As String) As String()
  'Tools > References:  Microsoft Script Control 1.0
    Dim description As String
    Dim temperature As Double
    Dim humidity As Double
    Dim pressure As Double
    Dim answer As String
    Dim parameters(0 To 3) As String
   
    description = runScript(json, "s.current.weather[0].description")
    temperature = runScript(json, "s.current.temp")
    humidity = runScript(json, "s.current.humidity")
    pressure = runScript(json, "s.current.pressure")
    
    parameters(0) = description
    parameters(1) = temperature & " °C"
    parameters(2) = humidity & "% R.H"
    parameters(3) = pressure & " mbar"
    
    get_parameters = parameters
 End Function
 
 Private Function runScript(response As String, query As String) As Variant
    'Tools > References:  Microsoft Script Control 1.0
    Dim script As MSScriptControl.scriptControl
    Dim data As Object
    
    Set script = New MSScriptControl.scriptControl
    script.Language = "JScript"
    Set data = script.Eval("(" + response + ")")
    With script
        .AddCode "function myfuncion(s) {return (" & query & ");}"
        runScript = .Run("myfuncion", data)
    End With
End Function
