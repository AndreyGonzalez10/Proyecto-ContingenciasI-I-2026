Attribute VB_Name = "emple_pensionados"

Private Function qx_p(edad As Integer, sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range) As Double
    Dim filaEdad As Variant
    Dim colAnio As Variant
    Dim tabla As Range

    If sexo = "Masculino" Then
        Set tabla = t_hombres
    Else
        Set tabla = t_mujeres
    End If

    filaEdad = Application.Match(edad, tabla.Columns(1), 0)
    If IsError(filaEdad) Then qx_p = 0: Exit Function

    colAnio = Application.Match(anio, tabla.Rows(1), 0)
    If IsError(colAnio) Then qx_p = 0: Exit Function

    qx_p = tabla.Cells(filaEdad, colAnio).Value
End Function


Private Function kpx_p(edad As Integer, sexo As String, k As Integer, anioBase As Integer, t_hombres As Range, t_mujeres As Range) As Double
    Dim j As Integer
    Dim prob As Double
    prob = 1
    For j = 0 To k - 1
        prob = prob * (1 - qx_p(edad + j, sexo, anioBase + j, t_hombres, t_mujeres))
    Next j
    kpx_p = prob
End Function


Private Function GrupoA(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, pensionados As Range) As Double
    Dim t As Integer
    t = anio - 2025
    If t < 0 Then GrupoA = 0: Exit Function

    Dim total As Double
    total = 0

    Dim i As Long
    For i = 1 To pensionados.Rows.Count
        Dim sexoPen As String
        Dim fechaNac As Date
        sexoPen = CStr(pensionados.Cells(i, 2).Value)
        fechaNac = CDate(pensionados.Cells(i, 3).Value)

        If sexoPen = sexo Then
            Dim edadBase As Integer
            edadBase = edadCalc(fechaNac)
            total = total + kpx_p(edadBase, sexoPen, t, 2025, t_hombres, t_mujeres)
        End If
    Next i

    GrupoA = total
End Function


Private Function GrupoB(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, empleados As Range) As Double
    Dim t As Integer
    t = anio - 2025
    If t < 0 Then GrupoB = 0: Exit Function

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

            Dim tJub As Integer
            tJub = edadJub - edadBase

            If edadBase < edadJub And tJub >= 0 And tJub <= t Then
                Dim anioJub As Integer
                anioJub = 2025 + tJub

                Dim aniosComoPen As Integer
                aniosComoPen = t - tJub

                total = total + kpx_p(edadJub, sexoEmp, aniosComoPen, anioJub, t_hombres, t_mujeres)
            End If
        End If
    Next i

    GrupoB = total
End Function


Function ProbPensionadoEnAnio(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, pensionados As Range, empleados As Range) As Double
    ProbPensionadoEnAnio = GrupoA(sexo, anio, t_hombres, t_mujeres, pensionados) + _
                           GrupoB(sexo, anio, t_hombres, t_mujeres, empleados)
End Function

Function MuertesPensionadosEnAnio(sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range, pensionados As Range, empleados As Range) As Double
    Dim t As Integer
    t = anio - 2025
    If t < 0 Then MuertesPensionadosEnAnio = 0: Exit Function

    Dim total As Double
    total = 0

    Dim i As Long
    For i = 1 To pensionados.Rows.Count
        Dim sexoPen As String
        Dim fechaNac As Date
        sexoPen = CStr(pensionados.Cells(i, 2).Value)
        fechaNac = CDate(pensionados.Cells(i, 3).Value)

        If sexoPen = sexo Then
            Dim edadBase As Integer
            edadBase = edadCalc(fechaNac)
            total = total + kpx_p(edadBase, sexoPen, t, 2025, t_hombres, t_mujeres) * _
                            qx_p(edadBase + t, sexoPen, anio, t_hombres, t_mujeres)
        End If
    Next i

    Dim j As Long
    For j = 1 To empleados.Rows.Count
        Dim sexoEmp As String
        Dim fechaNacE As Date
        sexoEmp = CStr(empleados.Cells(j, 4).Value)
        fechaNacE = CDate(empleados.Cells(j, 5).Value)

        If sexoEmp = sexo Then
            Dim edadBaseE As Integer
            edadBaseE = edadCalc(fechaNacE)

            Dim edadJub As Integer
            If sexo = "Masculino" Then edadJub = 65 Else edadJub = 63

            Dim tJub As Integer
            tJub = edadJub - edadBaseE

            If edadBaseE < edadJub And tJub >= 0 And tJub <= t Then
                Dim anioJub As Integer
                anioJub = 2025 + tJub

                Dim aniosComoPen As Integer
                aniosComoPen = t - tJub

                total = total + kpx_p(edadJub, sexoEmp, aniosComoPen, anioJub, t_hombres, t_mujeres) * _
                                qx_p(edadJub + aniosComoPen, sexoEmp, anio, t_hombres, t_mujeres)
            End If
        End If
    Next j

    MuertesPensionadosEnAnio = total
End Function

