Attribute VB_Name = "empleados_activos"
Private Function qx_c(edad As Integer, sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range) As Double
    Dim filaEdad As Variant
    Dim colAnio As Variant
    Dim tabla As Range
    
    If sexo = "Masculino" Then
        Set tabla = t_hombres
    Else
        Set tabla = t_mujeres
    End If
    
    filaEdad = Application.Match(edad, tabla.Columns(1), 0)
    If IsError(filaEdad) Then qx_c = 0: Exit Function
    
    colAnio = Application.Match(anio, tabla.Rows(1), 0)
    If IsError(colAnio) Then qx_c = 0: Exit Function
    
    qx_c = tabla.Cells(filaEdad, colAnio).Value
End Function

Private Function kpx_c(edad As Integer, sexo As String, k As Integer, anioBase As Integer, t_hombres As Range, t_mujeres As Range) As Double
    Dim j As Integer
    Dim prob As Double
    prob = 1
    For j = 0 To k - 1
        prob = prob * (1 - qx_c(edad + j, sexo, anioBase + j, t_hombres, t_mujeres))
    Next j
    kpx_c = prob
End Function

Function ProbActivoEnAnio(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, empleados As Range) As Double

    Dim t As Integer
    t = anio - 2025
    If t < 0 Then ProbActivoEnAnio = 0: Exit Function

    Dim total As Double
    total = 0

    Dim i As Long
    For i = 1 To empleados.Rows.Count
        Dim sexoEmp As String
        Dim fechaNac As Date
        sexoEmp = CStr(empleados.Cells(i, 4).Value)
        fechaNac = CDate(empleados.Cells(i, 5).Value)

        If sexoEmp = sexo Then
            Dim edadBase As Integer
            edadBase = edadCalc(fechaNac)

            Dim edadJub As Integer
            If sexo = "Masculino" Then edadJub = 65 Else edadJub = 63

            If edadBase < edadJub And t <= edadJub - edadBase Then
                total = total + CDbl(kpx_c(edadBase, sexoEmp, t, 2025, t_hombres, t_mujeres))

            End If
        End If
    Next i

    ProbActivoEnAnio = total
End Function

Function MuertesEnAnio(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, empleados As Range) As Double

    Dim t As Integer
    t = anio - 2025
    If t < 0 Then MuertesEnAnio = 0: Exit Function

    Dim total As Double
    total = 0

    Dim i As Long
    For i = 1 To empleados.Rows.Count
        Dim sexoEmp As String
        Dim fechaNac As Date
        sexoEmp = CStr(empleados.Cells(i, 4).Value)
        fechaNac = CDate(empleados.Cells(i, 5).Value)

        If sexoEmp = sexo Then
            Dim edadBase As Integer
            edadBase = edadCalc(fechaNac)

            Dim edadJub As Integer
            If sexo = "Masculino" Then edadJub = 65 Else edadJub = 63

            If edadBase < edadJub And edadBase + t < edadJub Then
                total = total + kpx_c(edadBase, sexoEmp, t, 2025, t_hombres, t_mujeres) * _
                                qx_c(edadBase + t, sexoEmp, anio, t_hombres, t_mujeres)
            End If
        End If
    Next i

    MuertesEnAnio = total
End Function

