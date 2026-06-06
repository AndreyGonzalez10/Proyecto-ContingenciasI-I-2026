Attribute VB_Name = "macro_esperanzaVida"
Sub CalcularEsperanzaVida(nombreHoja As String, colSex As Integer, colFecha As Integer, colSalida As Integer)
    Dim i As Long
    Dim edad As Integer
    Dim ev As Double
    Dim sexo As String
    
    Dim tablaH As Range
    Dim tablaM As Range
    Dim ws As Worksheet
    
    Set ws = Worksheets(nombreHoja)
    
    Set tablaH = Worksheets("Hombres").Range("A1").CurrentRegion
    Set tablaM = Worksheets("Mujeres").Range("A1").CurrentRegion
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        sexo = ws.Cells(i, colSex).Value
        edad = edadCalc(ws.Cells(i, colFecha).Value)
        ev = ex(edad, sexo, 2025, tablaH, tablaM)
        ws.Cells(i, colSalida).Value = ev
    Next i
            
End Sub



Sub EsperanzaVida_Empleados()
 Call CalcularEsperanzaVida("Empleados Activos", 4, 5, 6)
End Sub


Sub EsperanzaVida_Pensionados()
 Call CalcularEsperanzaVida("Pensionados", 2, 3, 5)
End Sub



