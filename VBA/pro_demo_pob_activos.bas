Attribute VB_Name = "pro_demo_pob_activos"
' Modulo 1
' Hecho por Anthonny Flores C32975 (no se si aqui ponemos los nombres de todos)
' Proyeccion demografica anual de empleados activos por sexo
' Proyecto Final - Contingencias de Vida I
' Escuela de Matematica, UCR - I Semestre 2026


Option Explicit

' Diccionario global para la tabla de mortalidad SP-2015
' Clave: "sex_edad_año" ? Valor: qx
Dim dictQx As Object

' CargarTablaMortalidad
' Lee toda la hoja "Mortalidad" una sola vez y la guarda en
' un diccionario para evitar buscar fila por fila cada vez
Sub CargarTablaMortalidad()

    Set dictQx = CreateObject("Scripting.Dictionary")

    Dim wsMort As Worksheet
    Set wsMort = ThisWorkbook.Sheets("qxhasta2150-v2018")

    Dim ultima As Long
    ultima = wsMort.Cells(wsMort.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim clave As String

    For i = 2 To ultima
        ' Formato clave: sex_edad_year (ej: "1_65_2025")
        clave = CStr(wsMort.Cells(i, 1).Value) & "_" & _
                CStr(wsMort.Cells(i, 2).Value) & "_" & _
                CStr(wsMort.Cells(i, 4).Value)

        If Not dictQx.Exists(clave) Then
            dictQx.Add clave, wsMort.Cells(i, 6).Value
        End If
    Next i

End Sub

' ObtenerQx
' Consulta la probabilidad de muerte qx del diccionario
' sex: 1 = Masculino, 2 = Femenino
' edad: edad entera de la persona
' año: año del calendario (tabla dinamica)
Function ObtenerQx(sex As Integer, edad As Integer, año As Integer) As Double

    Dim clave As String
    clave = CStr(sex) & "_" & CStr(edad) & "_" & CStr(año)

    If dictQx.Exists(clave) Then
        ObtenerQx = dictQx(clave)
    Else
        ObtenerQx = 0
    End If

End Function

' EdadAlFinDeAño
' Calcula la edad cumplida al 31 de diciembre del año indicado
' fechaNac: numero serial de Excel de la fecha de nacimiento
' año: año objetivo del calculo
Function EdadAlFinDeAño(fechaNac As Long, año As Integer) As Integer

    Dim fFin As Date
    Dim fNac As Date

    fFin = DateSerial(año, 12, 31)
    fNac = CDate(fechaNac)

    EdadAlFinDeAño = DateDiff("yyyy", fNac, fFin) - _
                     IIf(Format(fFin, "mmdd") < Format(fNac, "mmdd"), 1, 0)

End Function

' CodSexo
' Convierte el texto del sexo al codigo numerico de la tabla SP-2015
' "Masculino" ? 1, "Femenino" ? 2
Function CodSexo(s As String) As Integer

    If LCase(Trim(s)) = "masculino" Then
        CodSexo = 1
    Else
        CodSexo = 2
    End If

End Function

' PuntoA_PoblacionActivos
' Genera la hoja "a) Pob Activos" con la cantidad esperada de
' empleados activos por sexo para cada año del horizonte,
' usando la tabla dinamica SP-2015 y la edad de jubilacion
' como unica salida distinta a la muerte (65H / 63M)
Sub PuntoA_PoblacionActivos()

    Call CargarTablaMortalidad

    Dim wsEmp As Worksheet
    Set wsEmp = ThisWorkbook.Sheets("Empleados Activos")

    ' Limpiar y crear hoja de resultado
    Dim nombreHoja As String
    nombreHoja = "a) Proyeccion Poblacion Activos"

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

    ' Buscar el año en que se jubila el empleado mas joven
    ' ese es el horizonte maximo de la proyeccion
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

    ' Arrays donde se acumula la poblacion activa esperada por año
    Dim pobH() As Double
    Dim pobM() As Double
    ReDim pobH(0 To nAños)
    ReDim pobM(0 To nAños)

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

        If codS = 1 Then
            pobH(0) = pobH(0) + prob
        Else
            pobM(0) = pobM(0) + prob
        End If

        For t = 0 To nAños - 1

            añoT = AÑO_CORTE + t
            edadT = edadI + t

            ' Sale del grupo activo al jubilarse
            If edadT >= jub Then Exit For

            qx = ObtenerQx(codS, edadT, añoT)

            ' P(activo en t+1) = P(activo en t) * (1 - qx)
            prob = prob * (1# - qx)

            ' Solo acumula si en t+1 aun no se jubila
            If (edadT + 1) < jub Then
                If codS = 1 Then
                    pobH(t + 1) = pobH(t + 1) + prob
                Else
                    pobM(t + 1) = pobM(t + 1) + prob
                End If
            End If

        Next t

SigEmp:
    Next i

    ' Encabezados
    wsR.Cells(1, 1).Value = "Año"
    wsR.Cells(1, 2).Value = "Hombres Activos"
    wsR.Cells(1, 3).Value = "Mujeres Activas"
    wsR.Cells(1, 4).Value = "Total Activos"

    ' Resultados
    For t = 0 To nAños
        wsR.Cells(t + 2, 1).Value = AÑO_CORTE + t
        wsR.Cells(t + 2, 2).Value = Round(pobH(t), 2)
        wsR.Cells(t + 2, 3).Value = Round(pobM(t), 2)
        wsR.Cells(t + 2, 4).Value = Round(pobH(t) + pobM(t), 2)
    Next t

    ' Formato de la hoja
    With wsR.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(54, 96, 146)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsR.Columns("A:D").AutoFit
    wsR.Columns("B:D").NumberFormat = "#,##0.00"

End Sub

