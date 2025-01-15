Public Function Array_AddColumn(Arry As Variant, Optional TieneHeader As Boolean = True, Optional HeaderName As String = "ColumnaAgregada", _
                                Optional DatoUnico As Boolean = False, Optional ValorUnico As String = "", Optional DatosDiferentes As Boolean = False, _
                                Optional OrArryDatosNx1 As Variant = Empty, Optional OrVectorDeDatosN As Variant = Empty, Optional ColumnPositionToAdd As Integer = -1) As Variant
    Dim n1&, n2&, i&, m1&, m2&, m3&, j&, j1&, k1%
    
    ' Validación del array de entrada
    On Error GoTo ErrorHandler
    If IsEmpty(Arry) Or Not IsArray(Arry) Then Err.Raise 1001, , "El array de entrada no es válido."
    n1 = LBound(Arry, 1): n2 = UBound(Arry, 1)
    m1 = LBound(Arry, 2): m2 = UBound(Arry, 2)

    ' Validación de la posición de la columna
    If ColumnPositionToAdd < -1 Or ColumnPositionToAdd > m2 + 1 Then Err.Raise 1002, , "La posición de la columna es inválida."

    ' Delegar al final si la posición es -1
    If ColumnPositionToAdd = -1 Then
        Array_AddColumn = Array_AddColumnToEnd(Arry, TieneHeader, HeaderName, DatoUnico, ValorUnico, DatosDiferentes, OrArryDatosNx1, OrVectorDeDatosN)
        Exit Function
    End If

    ' Redimensionar el nuevo array
    m3 = m2 - m1 + 1
    ReDim ArryNew(n1 To n2, m1 To m3 + 1)

    ' Copiar encabezados
    If TieneHeader Then
        For j = m1 To m2
            j1 = IIf(j < ColumnPositionToAdd, j, j + 1)
            ArryNew(1, j1) = Arry(1, j)
        Next j
        ArryNew(1, ColumnPositionToAdd) = HeaderName
        k1 = 2
    Else
        k1 = 1
    End If

    ' Agregar datos
    For i = k1 To n2
        For j = m1 To m2
            j1 = IIf(j < ColumnPositionToAdd, j, j + 1)
            ArryNew(i, j1) = Arry(i, j)
        Next j

        If DatoUnico Then
            ArryNew(i, ColumnPositionToAdd) = ValorUnico
        ElseIf DatosDiferentes Then
            If Not IsEmpty(OrArryDatosNx1) Then
                ArryNew(i, ColumnPositionToAdd) = OrArryDatosNx1(i, 1)
            ElseIf Not IsEmpty(OrVectorDeDatosN) Then
                ArryNew(i, ColumnPositionToAdd) = OrVectorDeDatosN(i, 1)
            End If
        End If
    Next i

    Array_AddColumn = ArryNew
    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Array_AddColumn = Arry
End Function

Public Function Array_AddColumnToEnd(Arry As Variant, Optional TieneHeader As Boolean = True, Optional HeaderName As String = "ColumnaAgregada", _
                                     Optional DatoUnico As Boolean = False, Optional ValorUnico As String = "", Optional DatosDiferentes As Boolean = False, _
                                     Optional ArryDatosNx1 As Variant = Empty, Optional VectorDeDatosN As Variant = Empty) As Variant
    Dim n1&, n2&, i&, m1&, m2&, j&, k1%
    
    ' Validación del array de entrada
    On Error GoTo ErrorHandler
    If IsEmpty(Arry) Or Not IsArray(Arry) Then Err.Raise 1001, , "El array de entrada no es válido."
    n1 = LBound(Arry, 1): n2 = UBound(Arry, 1)
    m1 = LBound(Arry, 2): m2 = UBound(Arry, 2)

    ' Redimensionar el nuevo array
    ReDim ArryNew(n1 To n2, m1 To m2 + 1)

    ' Copiar encabezados
    If TieneHeader Then
        For j = m1 To m2
            ArryNew(1, j) = Arry(1, j)
        Next j
        ArryNew(1, m2 + 1) = HeaderName
        k1 = 2
    Else
        k1 = 1
    End If

    ' Agregar datos
    For i = k1 To n2
        For j = m1 To m2
            ArryNew(i, j) = Arry(i, j)
        Next j

        If DatoUnico Then
            ArryNew(i, m2 + 1) = ValorUnico
        ElseIf DatosDiferentes Then
            If Not IsEmpty(ArryDatosNx1) Then
                ArryNew(i, m2 + 1) = ArryDatosNx1(i, 1)
            ElseIf Not IsEmpty(VectorDeDatosN) Then
                ArryNew(i, m2 + 1) = VectorDeDatosN(i, 1)
            End If
        End If
    Next i

    Array_AddColumnToEnd = ArryNew
    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Array_AddColumnToEnd = Arry
End Function

Public Function Array_ColumnExtract(Arry As Variant, nCol&, Optional nRowIni& = 1, Optional ToVector As Boolean = False) As Variant
    Dim n0&, n1&, i&, nRow&, Arry2 As Variant

    ' Validar parámetros de entrada
    If IsEmpty(Arry) Then Exit Function
    If nCol < LBound(Arry, 2) Or nCol > UBound(Arry, 2) Then Exit Function
    If nRowIni < 1 Or nRowIni > UBound(Arry, 1) Then Exit Function

    ' Inicializar límites y calcular filas a extraer
    n0 = LBound(Arry, 1)
    n1 = UBound(Arry, 1)
    nRow = n1 - n0 + 1 - (nRowIni - 1)

    ' Extraer columna según formato requerido
    If ToVector Then
        If nRow > 0 Then
            ReDim Arry2(1 To nRow)
            For i = n0 + (nRowIni - 1) To n1
                Arry2(i - n0 - (nRowIni - 1) + 1) = Arry(i, nCol)
            Next i
        Else
            ReDim Arry2(0 To 0)
        End If
    Else
        If nRow > 0 Then
            ReDim Arry2(1 To nRow, 1 To 1)
            For i = n0 + (nRowIni - 1) To n1
                Arry2(i - n0 - (nRowIni - 1) + 1, 1) = Arry(i, nCol)
            Next i
        Else
            ReDim Arry2(0 To 0, 0 To 0)
        End If
    End If

    ' Asignar el resultado
    Array_ColumnExtract = Arry2
End Function
