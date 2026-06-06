Attribute VB_Name = "mod_mortalidad"
'Calcula la edad al 31/12/20251 con DateSerial'
Function edadCalc(fechaNacimiento As Date) As Integer
    Dim fechaCorte As Date
    fechaCorte = DateSerial(2025, 12, 31)
    
    edadCalc = DateDiff("yyyy", fechaNacimiento, fechaCorte)
    
    If DateSerial(2025, Month(fechaNacimiento), Day(fechaNacimiento)) > fechaCorte Then
        edadCalc = edadCalc - 1
    End If
End Function

'Calcula los px=1-qx usando una tabla de mortalidad (asumiendo que tiene columnas x,y,qx_hombre, qx_mujer)'
Function qx(edad As Integer, sexo As String, anio As Integer, t_hombres, t_mujeres As Range) As Double
    Dim filaEdad As Variant
    Dim colAnio As Variant
    Dim tabla As Range
    
    If sexo = "Masculino" Then
        Set tabla = t_hombres
    Else
        Set tabla = t_mujeres
    End If
    
    filaEdad = Application.Match(edad, tabla.Columns(1), 0)
    
    If IsError(filaEdad) Then
        MsgBox "No se encontró la edad"
        qx = -1
        Exit Function
    End If
    
    colAnio = Application.Match(anio, tabla.Rows(1), 0)
    If IsError(colAnio) Then
        MsgBox "No se encontró el anio"
        qx = -1
        Exit Function
    End If
    
    qx = tabla.Cells(filaEdad, colAnio).Value
    
End Function

'Calcula la probabilidad de sobrevivir 1 anio'
Function px(edad As Integer, sexo As String, anio As Integer, t_hombres As Range, t_mujeres As Range) As Double
    px = 1 - qx(edad, sexo, anio, t_hombres, t_mujeres)
End Function

'Calcula la probabilidad de sobrevivir k anios'

Function kpx(edad As Integer, sexo As String, k As Integer, anioBase As Integer, t_hombres As Range, t_mujeres As Range)
    Dim j As Integer
    Dim prob As Double
    
    prob = 1
    
    For j = 0 To k - 1
        prob = prob * px(edad + j, sexo, anioBase + j, t_hombres, t_mujeres)
    Next j
    
    kpx = prob
End Function

'Funcion de esperanza de vida para una persona'
Function ex(edad As Integer, sexo As String, anioBase As Integer, t_hombres As Range, t_mujeres As Range) As Double
    Dim edadMax As Integer
    Dim k As Integer
    Dim suma As Double
    
    edadMax = Application.Max(t_hombres.Columns(1))
    suma = 0
    
    For k = 1 To edadMax - edad
        suma = suma + kpx(edad, sexo, k, anioBase, t_hombres, t_mujeres)
    Next k
    
    ex = suma
End Function


'Funcion para calcular las primas por edad y sexo'
'Function prima_edad_sexo(edad as Integer, sexo as String, salario as Double, anno as Integer, anno_corte as Integer, tabla as Range, edad_pensionM as Integer, edad
'
