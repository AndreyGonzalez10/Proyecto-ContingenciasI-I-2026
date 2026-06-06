Attribute VB_Name = "pro_demo_muertes_activos"
' Modulo 2
' Hecho por Anthonny Flores C32975
' Proyeccion demografica anual de muertes de empleados activos por sexo
' Proyecto Final - Contingencias de Vida I
' Escuela de Matematica, UCR - I Semestre 2026


Option Explicit

' PuntoC_MuertesActivos
' Genera la hoja "c) Muertes Activos" con las muertes esperadas
' de empleados activos por sexo para cada año del horizonte
' Muertes(t) = sum_i P(activo en t) * qx(edad_i, año_t, sexo_i)
Sub PuntoC_MuertesActivos()

    Call CargarTablaMortalidad

    Dim wsEmp As Worksheet
    Set wsEmp = ThisWorkbook.Sheets("Empleados Activos")

    ' Limpiar y crear hoja de resultado
    Dim nombreHoja As String
    nombreHoja = "c) Proyeccion Muertes Activos"

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(nombreHoja).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsR.Name = nombreHoja

    Const AÑO_CORTE  As Integer = 2025
    Const EDAD_JUB_M As Integer = 65
    Const EDAD_JUB_F As Integer = 63

    Dim ultima As Long
    ultima = wsEmp.Cells(wsEmp.Rows.Count, 1).End(xlUp).Row

    Dim nEmp As Integer
    nEmp = ultima - 1

    ' Leer datos en arrays para no acceder a celdas dentro del loop
    Dim sexArr() As String
    Dim nacArr() As Long
    ReDim sexArr(1 To nEmp)
    ReDim nacArr(1 To nEmp)

    Dim i As Integer
    For i = 1 To nEmp
        sexArr(i) = CStr(wsEmp.Cells(i + 1, 4).Value)
        nacArr(i) = CLng(wsEmp.Cells(i + 1, 5).Value)
    Next i

    ' Horizonte maximo: año en que se jubila el empleado mas joven
    Dim añoMax As Integer
    añoMax = AÑO_CORTE

    Dim edadCorte As Integer, jubEdad As Integer, añoJub As Integer

    For i = 1 To nEmp
        edadCorte = EdadAlFinDeAño(nacArr(i), AÑO_CORTE)
        If LCase(Trim(sexArr(i))) = "masculino" Then
            jubEdad = EDAD_JUB_M
        Else
            jubEdad = EDAD_JUB_F
        End If
        añoJub = AÑO_CORTE + (jubEdad - edadCorte)
        If añoJub > añoMax Then añoMax = añoJub
    Next i

    Dim nAños As Integer
    nAños = añoMax - AÑO_CORTE

    ' Arrays donde se acumulan las muertes esperadas por año
    Dim muertesH() As Double
    Dim muertesM() As Double
    ReDim muertesH(0 To nAños - 1)
    ReDim muertesM(0 To nAños - 1)

    ' Variables del loop principal
    Dim codS As Integer, edadI As Integer, jub As Integer
    Dim prob As Double, qx As Double
    Dim t As Integer, añoT As Integer, edadT As Integer

    For i = 1 To nEmp

        codS = CodSexo(sexArr(i))
        edadI = EdadAlFinDeAño(nacArr(i), AÑO_CORTE)

        If codS = 1 Then jub = EDAD_JUB_M Else jub = EDAD_JUB_F

        ' Ignorar empleados que ya alcanzaron la edad de jubilacion al corte
        If edadI >= jub Then GoTo SigEmp

        ' Al corte el empleado esta activo con probabilidad 1
        prob = 1#

        For t = 0 To nAños - 1

            añoT = AÑO_CORTE + t
            edadT = edadI + t

            ' Sale del grupo activo al jubilarse
            If edadT >= jub Then Exit For

            qx = ObtenerQx(codS, edadT, añoT)

            ' Muertes esperadas en el año t
            If codS = 1 Then
                muertesH(t) = muertesH(t) + prob * qx
            Else
                muertesM(t) = muertesM(t) + prob * qx
            End If

            ' Actualizar prob de seguir activo para el año siguiente
            prob = prob * (1# - qx)

        Next t

SigEmp:
    Next i

    ' Encabezados
    wsR.Cells(1, 1).Value = "Año"
    wsR.Cells(1, 2).Value = "Muertes Hombres"
    wsR.Cells(1, 3).Value = "Muertes Mujeres"
    wsR.Cells(1, 4).Value = "Total Muertes"

    ' Resultados
    For t = 0 To nAños - 1
        wsR.Cells(t + 2, 1).Value = AÑO_CORTE + t
        wsR.Cells(t + 2, 2).Value = Round(muertesH(t), 4)
        wsR.Cells(t + 2, 3).Value = Round(muertesM(t), 4)
        wsR.Cells(t + 2, 4).Value = Round(muertesH(t) + muertesM(t), 4)
    Next t

    ' Formato de la hoja
    With wsR.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(54, 96, 146)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsR.Columns("A:D").AutoFit
    wsR.Columns("B:D").NumberFormat = "#,##0.0000"

End Sub
