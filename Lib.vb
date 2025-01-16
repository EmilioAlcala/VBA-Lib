Option Explicit
'Sub Btn_Config_Actualizar():    Call Config_Actualizar("Cfg", True):    End Sub
' ------ Indice de funciones ---------
' Public Function Array_AddColumn(Arry, Optional TieneHeader As Boolean = True, Optional HeaderName$ = "ColumnaAgregada", _
' Public Function Array_AddColumnToEnd(Arry, Optional TieneHeader As Boolean = True, Optional HeaderName$ = "ColumnaAgregada", Optional DatoUnico As Boolean = False, Optional ValorUnico$ = "", Optional DatosDiferentes As Boolean = False, Optional ArryDatosNx1 As Variant = Empty, Optional VectorDeDatosN As Variant = Empty)
' Public Function Array_ColumnExtract(Arry As Variant, nCol&, Optional nRowIni& = 1, Optional ToVector As Boolean = False) As Variant
' Public Function Array_Compare(Arry1, Arry2, Optional Fila1& = 1, Optional Fila2& = 1, Optional Columna1& = 1, Optional Columna2& = 1, Optional Msg$) As Boolean
' Public Function Array_ConsolidarDataRd(DataRd() As ReadXlsDatShts, Optional AddFileNameColumn As Boolean = True, Optional AddAutoIdColumn As Boolean = True) As ReadXlsDatShts
' Public Function Array_DeleteEmptyColumns(Arry As Variant)
' Public Function Array_DeleteEmptyRows(Arry As Variant)
' Public Function Array_DeleteEmptyRowsByColumn(Arry As Variant, Optional nColumn& = 1)
' Public Function Array_DeleteFirstRows(Arry As Variant, nRows&)
' Public Function Array_GetColumnNumbers(ArryDB As Variant, VectorFldLst As Variant)
' Public Function Array_HeaderColumnNumber(Arry As Variant, Elem As Variant, Optional nRow& = 1)
' Public Function Array_HeaderColumnNumberVector(Arry As Variant, ElemVector As Variant, Optional nRow& = 1)
' Public Function Array_Join(Arrys, Optional OptBase% = 1)
'Private Function Array_Join1D(Arrys As Variant, Optional OptBase% = 1) As Variant
'Private Function Array_Join2D(Arrys As Variant, Optional OptBase% = 1, Optional RemoveRow1 As Boolean = False) As Variant
' Public Function Array_JoinMultipleDataRd(DataRd() As ReadXlsDatShts, Optional AddFileNameColumn As Boolean = True, Optional AddAutoIdColumn As Boolean = True)
' Public Function Array_nDimensions(ByVal vArray As Variant) As Long
' Public Function Array_RowExtract(Arry As Variant, Optional nRow& = 1) As Variant
' Public Function Array_Transpose(Arry As Variant, Optional Base1 As Boolean = True) As Variant
' Public Function Array_TrimRow(Arry As Variant, nRow&)
' Public Function Array_VLookUp(Arry2D As Variant, LstCampos_Llenar As Variant, ShtName_Extraer$, LstCampo_Extraer As Variant, _
' Public Function ArrayDB_AutoIndex(Arry As Variant, nColumn&, Optional IniRow As Integer = 2)
' Public Function ArrayDB_FilterByColumn(Data As Variant, Hdr As Variant, ColumnName As String, ArryConditionsOk As Variant)
' Public Function ArrayDB_ImportFields(Header2D As Variant, Data2D As Variant, LstFields$, _
' Public Function ArrayDB_Join(Arrys As Variant, Optional OptBase% = 1, Optional IncludeHeader As Boolean = True) As Variant
' Public Function ArrayDB_ToSheet(ShtName$, Optional Header As Variant = Empty, Optional Data As Variant = Empty, _
' Public Function ArrayDB_ToSheetSimple(Arry2D, ShtName, Optional ActualizarDatos As Boolean = True, Optional ActualizarHeader As Boolean = False, _
' Public Function ArrayDB_xImport_Data(DataRd As Variant, LstFields$, nCol As Variant, Optional FillIndex As Boolean = False, Optional nIdxCol% = 1)
' Public Function ArrayDB_xImport_GetColumns(DataRd, LstFields$, Optional XlsFile$, Optional Msg$) As Variant
' Public Function ArrayDB_xImport_GetPosition(Arry2D, Elem)
' Public Function ArrayDB_xImport_TestHdr(DataRd, LstFields$, XlsFile$, Optional ReportBlkHdrs As Boolean = False) As String
' Public Function Collection_Exists(Key$, coll As Collection) As Boolean
' Public Function Config_Actualizar(Optional ShtName$ = "Cfg", Optional UpdateCode As Boolean = True)
' Public Function Config_FullFileVariable(VarName$) As SeparaPathFile
'Private Function Config_GenClassCode()
' Public Function Config_List(LstName$, _
' Public Function Config_ListBuscarEquivalente(Element$, Lst1_Name$, Lst2_Name$) As String
' Public Function Config_ListDateFields(LstFields$)
' Public Function Config_ListEquivalent(Element As String, Lst1, Lst2)
' Public Function Config_SetVariable(VarName$, Valor As Variant)
'Private Function Config_TipoVariable(Variable, Valor) As String
' Public Function Config_Variable(VarName$, Optional CfgName$ = "Cfg")
' Public Function Create_EmptyArray_FromSheet(Filas, Columnas)
' Public Function Create_IdxColl_ForDB(Arry2D As Variant, ColumName$, ShtName$, Optional MsgDup$ = "")
' Public Function DataSheet_ChecarHayDatos(ShtName) As Boolean
' Public Function DataSheet_DefineDBName(Wbk As Workbook, ShtName$)
' Public Function DataSheet_DeleteData(ShtName$, Optional CreateDBname As Boolean = False, Optional FilaDatos% = 2, Optional DeleteRows As Boolean = False, Optional DeleteUsedRgRows As Boolean = True)
' Public Function DataSheet_FilterCancel(ShtName$)
' Public Function FE_DateToExcelData(Fecha As Date) As Variant
' Public Function FE_DateToStr(Fecha As Date) As String
' Public Function FE_GetDate(Fecha As Variant) As Date
' Public Function Find_InArrayDB(Id$, ArrayDB As Variant, IdxColl As Collection)
'Private Function Find_Index(Clave$, IdxColl As Collection) As Long
'Private Function Format_AlignColumns(ShtName$, sLstColumnsAlign, Optional hAlign As Excel.Constants = xlGeneral, Optional vAlign As Excel.Constants = xlCenter)
'Private Function Format_Columns(ShtName$, sLstColumnNames$, NumFormat$)
' Public Function GetRowForFilter(RgName$, ColNameToFilter$, Condition As Variant, Optional UseType As String = "String")
' Public Function OS_CreateDirectory(RootPath As String, SubFolder As String)
' Public Function OS_DialogoAbrirArchivo(FullFileName$, TituloDialogo$, DialogFilter$) As String
' Public Function OS_DialogoSelectFolder(Optional IniPath$ = "") As String
' Public Function OS_GetFile2(FullFileName) As SeparaPathFile
' Public Function OS_StrFileList(Path_Ini$, FileName$, Optional SearchSubFolders As Boolean = False, Optional FullPath As Boolean = True)
' Public Function OS_SubFolderFileListArray(Path_Ini$, FileName$, Optional SearchSubFolders As Boolean = False, Optional FullPath As Boolean = True)
' Public Function ReadSheet_HeaderAndData(ShtNameToRead As String, _
' Public Function ReadSheet_ImportColumns(ShtNameToRead As String, LstFields As String, _
' Public Function ReadSheet_ImportFields(ShtRead$, ShtDestino$) As ReadXlsDatShts
' Public Function ReadSheet_NamedRangeToArray(Wbk As Workbook, ShtName$, Optional Msg$ = "")
' Public Function ReadSheet_ToArray(ShtNameToRead As String, _
' Public Function ReadSheet_UsedRangeToArray(Wbk As Workbook, ShtName$, Optional Msg$ = "")
' Public Function ReadSheet_UsedRangeToArrayDB(Wbk As Workbook, ShtName$, Optional Msg$ = "") As ReadXlsDatShts
' Public Function ReadXlsSheet_HeaderAndData(FullFile$, _
' Public Function ReadXlsSheet_ImportColumns(XlsFullName$, ShtNameToRead$, LstFields$, Optional WbkPass$ = "", _
' Public Function ReadXlsSheet_ImportColumns_MultipShts(XlsFullName$, VectorShtNameToRead, VectorLstFields, Optional VectorWbkPass As Variant = Empty, _
'Private Function ReadXlsSheet_DataImportFields(ShtNameToRead$, DataRd As ReadXlsDatShts, LstFields$, Optional TestHdrs As Boolean = True, Optional ErrMsg$ = "") As ReadXlsDatShts
' Public Function ReadXlsSheet_ImportColumns_MultipWbks2(VectorXlsFullName As Variant, VectorShtNameToRead As Variant, VectorLstFields As Variant, _
' Public Function ReadXlsSheet_ToArray(XlsFullName$, Optional ShtNameToRead$ = "", Optional WbkPass$ = "", _
' Public Function ReadXlsSheet_UsedRangeToArray(XlsFullName, Optional ShtNameToRead$ = "", Optional WbkPass$ = "", Optional Msg$ = "")
' Public Function ReadXlsSheetsAll_ToMultipleArray(XlsFullName, Optional WbkPass As String = "", Optional Msg As String = "") As ReadXlsDatShts()
' Public Function SheetExists(Wbk As Workbook, ShtName$) As Boolean
' Public Function STR_AnchoFijo(s1, Width As Long, Optional ReplaceStr = " ", Optional Brackets As Boolean = False, Optional NumEspacios& = 0)
'Private Function VBA_CrearModulo(Wbk As Workbook, VbModuleName$, StrContenidoModulo$, Optional CodeType As vbext_ComponentType = vbext_ct_StdModule) As VBIDE.VBComponent   '
' Public Function STR_SinAcentosMinusc(s1$)
' Public Function Texto_BuscarFrases(Txt$, VectorDeTxt As Variant) As Boolean
' Public Function Vector_FindElementPosition(Elem, ArryVector) As Long
' Public Function Vector_SetDataHeaderIdx(ByRef Array2D As Variant, Header2D As Variant, FieldName, Value As Variant, Optional Fecha As Boolean = False)
' Public Function Xls_CopySheetsToNewWbk(ShtsArray As Variant, FileName As String, FullPath As String, _
' Public Function Xls_LimpiarNombres(Wbk As Workbook, LstName As String, Optional DeleteName As Boolean = False)
' Public Function Xls_OcultarColumna(Sht As Worksheet, nCol As Long)
' Public Function Xls_OrdenarDB(ShtName$, RgName$, VectorColumnNames As Variant, VectorXlSortOrder As Variant, Optional Wbk As Workbook)
' Public Function Xls_ShowSheets(ShtToShw)

'--------------------------------- Array_AddColumn ---- Verificada ChatGPT
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

'--------------------------------- Array_AddColumnToEnd ---- Verificada ChatGPT
Public Function Array_AddColumnToEnd(Arry As Variant, Optional TieneHeader As Boolean = True, _
                                     Optional HeaderName As String = "ColumnaAgregada", _
                                     Optional DatoUnico As Boolean = False, _
                                     Optional ValorUnico As String = "", _
                                     Optional DatosDiferentes As Boolean = False, _
                                     Optional ArryDatosNx1 As Variant = Empty, _
                                     Optional VectorDeDatosN As Variant = Empty) As Variant
    Dim n1 As Long, n2 As Long, m1 As Long, m2 As Long
    Dim i As Long, j As Long, StartRow As Long
    Dim ArryNew As Variant

    ' Validación inicial del array de entrada
    On Error GoTo ErrorHandler
    If IsEmpty(Arry) Or Not IsArray(Arry) Then Err.Raise 1001, , "El array de entrada no es válido."
    If Not IsArray(Arry) Or LBound(Arry, 1) > UBound(Arry, 1) Then Err.Raise 1002, , "El array no tiene datos válidos."

    ' Determinar límites del array de entrada
    n1 = LBound(Arry, 1): n2 = UBound(Arry, 1)
    m1 = LBound(Arry, 2): m2 = UBound(Arry, 2)

    ' Validar consistencia de DatosDiferentes
    If DatosDiferentes Then
        If IsEmpty(ArryDatosNx1) And IsEmpty(VectorDeDatosN) Then _
            Err.Raise 1003, , "Debe proporcionar ArryDatosNx1 o VectorDeDatosN para DatosDiferentes."
        If Not IsEmpty(ArryDatosNx1) And UBound(ArryDatosNx1, 1) <> n2 Then _
            Err.Raise 1004, , "ArryDatosNx1 no coincide con las filas del array de entrada."
        If Not IsEmpty(VectorDeDatosN) And UBound(VectorDeDatosN) <> n2 Then _
            Err.Raise 1005, , "VectorDeDatosN no coincide con las filas del array de entrada."
    End If

    ' Redimensionar el nuevo array con una columna adicional
    ReDim ArryNew(n1 To n2, m1 To m2 + 1)

    ' Copiar encabezados si corresponde
    StartRow = IIf(TieneHeader, 2, 1)
    For i = n1 To n2
        For j = m1 To m2
            ArryNew(i, j) = Arry(i, j)
        Next j
        If i = n1 And TieneHeader Then
            ArryNew(i, m2 + 1) = HeaderName
        ElseIf DatoUnico Then
            ArryNew(i, m2 + 1) = ValorUnico
        ElseIf DatosDiferentes Then
            If Not IsEmpty(ArryDatosNx1) Then
                ArryNew(i, m2 + 1) = ArryDatosNx1(i, 1)
            ElseIf Not IsEmpty(VectorDeDatosN) Then
                ArryNew(i, m2 + 1) = VectorDeDatosN(i)
            End If
        End If
    Next i

    ' Devolver el array final
    Array_AddColumnToEnd = ArryNew
    Exit Function

ErrorHandler:
    ' Manejo de errores detallado
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error en Array_AddColumnToEnd"
    Array_AddColumnToEnd = Arry
End Function


'--------------------------------- Array_ColumnExtract ---- Verificada ChatGPT
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

'--------------------------------- Array_Compare
Public Function Array_Compare(Arry1, Arry2, Optional Fila1& = 1, Optional Fila2& = 1, Optional Columna1& = 1, Optional Columna2& = 1, Optional Msg$) As Boolean
    Dim n1&, m1&, n2&, m2&, n3&, m3&, n4&, m4&
    n1 = LBound(Arry1, 1)
    n2 = LBound(Arry2, 1)
    n3 = UBound(Arry1, 1)
    n4 = UBound(Arry2, 1)

    m1 = LBound(Arry1, 2)
    m2 = LBound(Arry2, 2)
    m3 = UBound(Arry1, 2)
    m4 = UBound(Arry2, 2)

    Msg = ""
    Array_Compare = False
    If n1 <> n2 Or n3 <> n4 Then
        Msg = Msg & "Diferente número de filas..." & vbCrLf
    End If
    If m1 <> m2 Or m3 <> m4 Then
        Msg = Msg & "Diferente número de columnas..." & vbCrLf
    End If
    If Msg = "" Then
        Dim i&, j&
        Array_Compare = True
        For i = Fila1 To Fila2
            For j = Columna1 To Columna2
                If Arry1(i, j) <> Arry2(i, j) Then
                    Msg = Msg & "Diferencia (" & i & ", " & j & ") >> " & Arry1(i, j) & " / " & Arry2(i, j) & vbCrLf
                    Array_Compare = False
                End If
            Next j
        Next
    End If
End Function

'--------------------------------- Array_ConsolidarDataRd
Public Function Array_ConsolidarDataRd(DataRd() As ReadXlsDatShts, Optional AddFileNameColumn As Boolean = True, Optional AddAutoIdColumn As Boolean = True) As ReadXlsDatShts
    Dim sLst$, sFolder$, ShtNameToRead$, i%, n%, Header, Data, Ret As ReadXlsDatShts
    sLst = "Consolidar"
    n = UBound(DataRd)
    For i = 1 To n
        sLst = sLst & "|" & DataRd(i).FileName
    Next i
    Header = DataRd(1).Header
    If AddAutoIdColumn Then
        Header = Lib.Array_AddColumn(Header, True, "Id", False, "", False, Empty, Empty, 1)
    End If

    sFolder = DataRd(1).Folder
    ShtNameToRead = DataRd(1).ShtName
    If AddFileNameColumn Then
        Header = Me.Array_AddColumnToEnd(Header, True, "Archivo", True, "")
    End If
    Data = Me.Array_JoinMultipleDataRd(DataRd:=DataRd, AddFileNameColumn:=AddFileNameColumn, AddAutoIdColumn:=AddAutoIdColumn)

    Ret.Data = Data
    Ret.Header = Header
    Ret.FileName = sLst
    Ret.Folder = sFolder
    Ret.ShtName = ShtNameToRead

    Array_ConsolidarDataRd = Ret
End Function

'--------------------------------- Array_DeleteEmptyColumns
Public Function Array_DeleteEmptyColumns(Arry As Variant)
    'Se asume Arry que es base 1
    Dim n&, m%, k%, nCol As Variant, q%, j%, i&, ColEmpty$, RetArry As Variant
    
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)
    k = 0
    ' nCol(0) no será tomada en cuenta
    ReDim nCol(0 To m) As Integer
    q = 0
    For j = 1 To m
        ColEmpty = ""
        For i = 1 To n
            If VarType(Arry(i, j)) = vbError Then
                ColEmpty = ColEmpty & "#Err"
                Arry(i, j) = "#Err"
            Else
                ColEmpty = ColEmpty & Arry(i, j)
            End If
            If Len(ColEmpty) > 0 Then
                Exit For
            End If
        Next i
        If ColEmpty <> "" Then
            q = q + 1
            nCol(q) = j
        End If
    Next j
    If q > 0 Then
        If q = m Then
            RetArry = Arry
        Else
            ReDim RetArry(1 To n, 1 To q)
            For i = 1 To n
                For j = 1 To q
                    RetArry(i, j) = Arry(i, nCol(j))
                Next j
            Next i
        End If
    Else
        ReDim RetArry(0, 0)
    End If
    Array_DeleteEmptyColumns = RetArry
End Function

'--------------------------------- Array_DeleteEmptyRows
Public Function Array_DeleteEmptyRows(Arry As Variant)
    ' Se asume que Arry es base 1
    Dim n&, m%, k&, nFil As Variant, i&, LinEmpty$, j%
    Dim RetArry As Variant
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)
    k = 0
    ' El elemento 0 no se tomará en cuenta
    ReDim nFil(0 To n)
    For i = 1 To n
        LinEmpty = ""
        For j = 1 To m
            'Debug.Print j
            If VarType(Arry(i, j)) = vbError Then
                LinEmpty = LinEmpty & "#Err"
                Arry(i, j) = "#Err"
            Else
                LinEmpty = LinEmpty & Arry(i, j)
            End If
            If Len(LinEmpty) > 0 Then
                Exit For
            End If
        Next j
        If LinEmpty <> "" Then
            k = k + 1
            nFil(k) = i
        End If
    Next i
    If k > 0 Then
        If k = n Then
            RetArry = Arry
        Else
            'ReDim RetArry(1 To k, 1 To m)
            RetArry = Create_EmptyArray_FromSheet(k, m)
            For i = 1 To k
                For j = 1 To m
                    RetArry(i, j) = Arry(nFil(i), j)
                Next j
            Next i
        End If
    Else
        ReDim RetArry(0, 0)
    End If
    Array_DeleteEmptyRows = RetArry
End Function

'--------------------------------- Array_DeleteEmptyRowsByColumn
Public Function Array_DeleteEmptyRowsByColumn(Arry As Variant, Optional nColumn& = 1)
    ' Se asume que Arry es base 1
    Dim n&, m%, k&, nFil As Variant, nFilx As Variant, kx&
    Dim i&, j%, RetArry As Variant
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)
    k = 0
    ' El elemento 0 no se tomará en cuenta
    ReDim nFil(0 To n) As Long
    ReDim nFilx(0 To n) As Long
    For i = 1 To n
        If VarType(Arry(i, nColumn)) = vbError Then
            Arry(i, nColumn) = "#Err"
            k = k + 1
            nFil(k) = i

        ElseIf Trim(Arry(i, nColumn)) = "" Then
            kx = kx + 1
            nFilx(kx) = i
            
        Else
            k = k + 1
            nFil(k) = i
            
        End If
    Next i
    
    If k > 0 Then
        If k = n Then
            RetArry = Arry
        Else
            'ReDim RetArry(1 To k, 1 To m)
            RetArry = Create_EmptyArray_FromSheet(k, m)
            For i = 1 To k
                For j = 1 To m
                    RetArry(i, j) = Arry(nFil(i), j)
                Next j
            Next i
        End If
    Else
        ReDim RetArry(0, 0)
    End If
    Array_DeleteEmptyRowsByColumn = RetArry
End Function

'--------------------------------- Array_DeleteFirstRows
Public Function Array_DeleteFirstRows(Arry As Variant, nRows&)
    ' Se asume que Arry es base 1
    Dim n&, m%, k&, i&, j%, RetArry As Variant
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)
    k = n - nRows
    If k > 0 Then
        If k = n Then
            RetArry = Arry
        Else
            'ReDim RetArry(1 To k, 1 To m)
            RetArry = Create_EmptyArray_FromSheet(k, m)
            For i = nRows + 1 To n
                For j = 1 To m
                    RetArry(i - nRows, j) = Arry(i, j)
                Next j
            Next i
        End If
    Else
        ReDim RetArry(0, 0)
    End If
    Array_DeleteFirstRows = RetArry
End Function

'--------------------------------- Array_GetColumnNumbers
Public Function Array_GetColumnNumbers(ArryDB As Variant, VectorFldLst As Variant)
    Dim k&, x1%, i&, nC As Variant
    If LBound(VectorFldLst) = 0 Then
        x1 = 1
    Else
        x1 = 0
    End If
    ReDim nC(0 To UBound(VectorFldLst) + x1)
    k = 0
    For i = LBound(VectorFldLst) To UBound(VectorFldLst)
        k = k + 1
        nC(k) = Lib.Array_HeaderColumnNumber(ArryDB, VectorFldLst(i), 1)
    Next i
    Array_GetColumnNumbers = nC
End Function

'--------------------------------- Array_HeaderColumnNumber
Public Function Array_HeaderColumnNumber(Arry As Variant, Elem As Variant, Optional nRow& = 1)
    Dim m1&, m2&, j&
    m1 = LBound(Arry, 2)
    m2 = UBound(Arry, 2)
    For j = m1 To m2
        If Arry(nRow, j) = Elem Then
            Array_HeaderColumnNumber = j
            Exit Function
        End If
    Next j
    Array_HeaderColumnNumber = -1
End Function

'--------------------------------- Array_HeaderColumnNumberVector
Public Function Array_HeaderColumnNumberVector(Arry As Variant, ElemVector As Variant, Optional nRow& = 1)
    Dim i%, m%, m1&, m2&, j&, nColumns As Variant, Elem As String, Fnd As Boolean
    m1 = LBound(Arry, 2)
    m2 = UBound(Arry, 2)
    
    m = UBound(ElemVector)
    ReDim nColumns(0 To m)
    For i = 0 To m
        Elem = ElemVector(i)
        nColumns(i) = Lib.Array_HeaderColumnNumber(Arry:=Arry, Elem:=Elem, nRow:=1)
        Fnd = False
        For j = m1 To m2
            If Arry(nRow, j) = Elem Then
                Fnd = True
                nColumns(i) = j
                Exit For
            End If
        Next j
        If Not Fnd Then
            nColumns(i) = -1
        End If
    Next i
    Array_HeaderColumnNumberVector = nColumns
End Function

'--------------------------------- Array_Join
Public Function Array_Join(Arrys, Optional OptBase% = 1)
    Dim nDim1%, NewArry As Variant
    nDim1 = Me.Array_nDimensions(Arrys(0))
    Select Case nDim1
        Case 1:     NewArry = Array_Join1D(Arrys, OptBase)
        Case 2:     NewArry = Array_Join2D(Arrys, OptBase, False)
    End Select
    Array_Join = NewArry
End Function

'--------------------------------- Array_Join1D (Private)
Private Function Array_Join1D(Arrys As Variant, Optional OptBase% = 1) As Variant
    Dim nElementos&, n&, i&, kOptBase&, ArryResult As Variant, Ini&, Fin&, a&

    ' Configurar valores según OptBase (base 0 o base 1)
    If OptBase = 0 Then
        kOptBase = 1
    Else
        kOptBase = 0
    End If

    ' Calcular número total de elementos en todos los vectores
    nElementos = 0
    For n = LBound(Arrys) To UBound(Arrys)
        nElementos = nElementos + (UBound(Arrys(n)) - LBound(Arrys(n)) + 1)
    Next n

    ' Redimensionar el vector de resultado
    ReDim ArryResult(1 - kOptBase To nElementos - kOptBase)

    ' Rellenar el nuevo vector con los valores de los vectores originales
    a = 1 - kOptBase
    For n = LBound(Arrys) To UBound(Arrys)
        ' Determinar el rango de elementos a copiar
        Ini = LBound(Arrys(n))
        Fin = UBound(Arrys(n))

        ' Copiar los datos de cada vector al vector nuevo
        For i = Ini To Fin
            ArryResult(a) = Arrys(n)(i)
            a = a + 1
        Next i
    Next n

    ' Devolver el vector unido
    Array_Join1D = ArryResult
End Function

'--------------------------------- Array_Join2D (Private)
Private Function Array_Join2D(Arrys As Variant, Optional OptBase% = 1, Optional RemoveRow1 As Boolean = False) As Variant
    Dim nFilas&, nColumnas&, n&, i&, j&, m&, a&, b&, a1&, b1&, kOptBase&
    Dim ArryResult As Variant, Ini&, Fin&
    
    ' Configurar valores según OptBase
    If OptBase = 0 Then
        kOptBase = 1
    Else
        kOptBase = 0
    End If

    ' Inicializar número total de filas y columnas
    nColumnas = UBound(Arrys(0), 2) - LBound(Arrys(0), 2) + 1
    nFilas = 0

    ' Calcular número total de filas en todas las matrices
    For n = LBound(Arrys) To UBound(Arrys)
        If RemoveRow1 Then
            nFilas = nFilas + (UBound(Arrys(n), 1) - LBound(Arrys(n), 1))
        Else
            nFilas = nFilas + (UBound(Arrys(n), 1) - LBound(Arrys(n), 1) + 1)
        End If
    Next n

    ' Redimensionar la matriz de resultado
    ReDim ArryResult(1 - kOptBase To nFilas - kOptBase, 1 - kOptBase To nColumnas - kOptBase)

    ' Rellenar la nueva matriz con las matrices originales
    a = 1 - kOptBase
    For n = LBound(Arrys) To UBound(Arrys)
        ' Determinar el rango de filas a copiar
        If RemoveRow1 Then
            Ini = LBound(Arrys(n), 1) + 1
            Fin = UBound(Arrys(n), 1)
        Else
            Ini = LBound(Arrys(n), 1)
            Fin = UBound(Arrys(n), 1)
        End If

        ' Copiar los datos de cada matriz a la nueva matriz
        For i = Ini To Fin
            b = 1 - kOptBase
            For j = LBound(Arrys(n), 2) To UBound(Arrys(n), 2)
                ArryResult(a, b) = Arrys(n)(i, j)
                b = b + 1
            Next j
            a = a + 1
        Next i
    Next n

    ' Devolver la nueva matriz unida
    Array_Join2D = ArryResult
End Function

'--------------------------------- Array_JoinMultipleDataRd
Public Function Array_JoinMultipleDataRd(DataRd() As ReadXlsDatShts, Optional AddFileNameColumn As Boolean = True, Optional AddAutoIdColumn As Boolean = True)
    Dim k&, i&, n&, m&, Arry1, TieneHeader As Boolean, FileName$, Arry2
    Dim n1&, m1&, x&, a&, b&, j&, k1&

    n = UBound(DataRd)
    n1 = 0
    For k = 1 To n
        Arry2 = DataRd(k).Data
        n1 = n1 + UBound(Arry2)
    Next k
    m = UBound(Arry2, 2)
    m1 = m
    If AddAutoIdColumn Then
        m1 = m1 + 1
    End If

    If AddFileNameColumn Then
        m1 = m1 + 1
        k1 = 1
    Else
        k1 = 0
    End If

    ReDim Arry1(1 To n1, 1 To m1)
    x = 0
    For k = 1 To n
        FileName = DataRd(k).FileName
        Arry2 = DataRd(k).Data
        a = LBound(Arry2)
        b = UBound(Arry2)
        For i = a To b
            x = x + 1
            For j = LBound(Arry2, 2) To UBound(Arry2, 2)
                Arry1(x, j + k1) = Arry2(i, j)
            Next j
            If AddFileNameColumn Then
                Arry1(x, m1) = FileName
            End If
        Next i
    Next k
    If AddAutoIdColumn Then
        For i = LBound(Arry1) To UBound(Arry1)
            Arry1(i, 1) = i
        Next i
    End If
    Array_JoinMultipleDataRd = Arry1
End Function

'--------------------------------- Array_nDimensions
Public Function Array_nDimensions(ByVal vArray As Variant) As Long
    Dim DimNum&, ErrorCheck As Variant
    On Error GoTo FinalDimension
    For DimNum = 1 To 100
        ErrorCheck = LBound(vArray, DimNum)
    Next
FinalDimension:
        Array_nDimensions = DimNum - 1
End Function

'--------------------------------- Array_RowExtract
Public Function Array_RowExtract(Arry As Variant, Optional nRow& = 1) As Variant
    Dim m0&, m1&, j&  ', Arry2
    m0 = LBound(Arry, 2)
    m1 = UBound(Arry, 2)

    If nRow > 0 Then
        ReDim Arry2(1 To 1, m0 To m1)
        For j = m0 To m1
            Arry2(1, j) = Arry(nRow, j)
        Next j
    Else
        ReDim Arry2(0 To 0, 0 To 0)
    End If

    Array_RowExtract = Arry2
End Function

'--------------------------------- Array_Transpose
Public Function Array_Transpose(Arry As Variant, Optional Base1 As Boolean = True) As Variant
    Dim m0&, n0&, m1&, n1&, m&, n&, i&, k1&, k2&, j%
    m0 = LBound(Arry, 1)
    n0 = LBound(Arry, 2)
    m1 = UBound(Arry, 1)
    n1 = UBound(Arry, 2)

    ReDim Arry2(n0 To n1, m0 To m1)

    If Base1 Then
        k1 = 0
        If m0 = 0 Then k1 = 1

        k2 = 0
        If n0 = 0 Then k2 = 1

        ReDim Arry2(n0 + k2 To n1 + k2, m0 + k1 To m1 + k1)
        For i = m0 To m1
            For j = n0 To n1
                Arry2(j + k2, i + k1) = Arry(i, j)
            Next j
        Next i
    Else
        ReDim Arry2(n0 To n1, m0 To m1)
        For i = m0 To m1
            For j = n0 To n1
                Arry2(j, i) = Arry(i, j)
            Next j
        Next i
    End If
    Array_Transpose = Arry2
End Function

'--------------------------------- Array_TrimRow
Public Function Array_TrimRow(Arry As Variant, nRow&)
    Dim n&, m%, j%, m0%
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)
    m0 = LBound(Arry, 2)
    For j = m0 To m
        Arry(nRow, j) = Trim(Arry(nRow, j))
    Next j
End Function

'--------------------------------- Array_VLookUp
Public Function Array_VLookUp(Arry2D As Variant, LstCampos_Llenar As Variant, ShtName_Extraer$, LstCampo_Extraer As Variant, _
            Optional Remplazar As Boolean = False, Optional TieneHdr As Boolean = True, Optional Header As Variant = Empty, Optional DNF_Value$ = "DNF")
    Dim n&, ColumnaClave$, Wbk As Workbook, DBExtraeData As Variant, IdxDB As New Collection, nC1 As Variant, nC2 As Variant
    Dim cRefLlenar%, cRefExtraer%, RowDat As Variant, Elem1$, i&, k%, Idx, a%
    n = UBound(Arry2D, 1)
    'm = UBound(Arry2D, 2)
    
    ColumnaClave = LstCampo_Extraer(0)
    Set Wbk = Application.ThisWorkbook
    DBExtraeData = Lib.ReadSheet_UsedRangeToArray(Wbk:=Wbk, ShtName:=ShtName_Extraer)
    'DBExtraeDataHdr = Lib.Array_RowExtract(DB, 1)
    'DBExtraeData = Lib.Array_DeleteFirstRows(DB, 1)
    Set IdxDB = Lib.Create_IdxColl_ForDB(Arry2D:=DBExtraeData, ColumName:=ColumnaClave, ShtName:=ShtName_Extraer)
    '==== Extraer los numeros de columnas de LstCampos_Llenar ====
    If TieneHdr Then
        nC1 = Lib.Array_GetColumnNumbers(Arry2D, LstCampos_Llenar)
        a = 2
    Else
        nC1 = Lib.Array_GetColumnNumbers(Header, LstCampos_Llenar)
        a = 1
    End If
    '==== Extraer los numeros de columnas de LstCampo_Extraer ====
    nC2 = Lib.Array_GetColumnNumbers(DBExtraeData, LstCampo_Extraer)
    
    cRefLlenar = nC1(1)
    cRefExtraer = nC2(1)
    
    For i = a To n
        Elem1 = "" & Trim(Arry2D(i, cRefLlenar))
        RowDat = Lib.Find_InArrayDB(Id:=Elem1, ArrayDB:=DBExtraeData, IdxColl:=IdxDB)
        If IsEmpty(RowDat(1, 1)) Then
            If Remplazar = True Then
                For k = 2 To UBound(nC1)
                    Arry2D(i, nC1(k)) = DNF_Value
                Next k
            Else
                For k = 2 To UBound(nC1)
                    If IsEmpty(Arry2D(i, nC1(k))) Or Arry2D(i, nC1(k)) = "" Then
                        Arry2D(i, nC1(k)) = RowDat(1, nC2(k))
                    End If
                Next k
            End If
        Else
            If Remplazar = True Then
                For k = 2 To UBound(nC1)
                    Arry2D(i, nC1(k)) = RowDat(1, nC2(k))
                Next k
            Else
                For k = 2 To UBound(nC1)
                    If IsEmpty(Arry2D(i, nC1(k))) Or Arry2D(i, nC1(k)) = "" Then
                        Arry2D(i, nC1(k)) = RowDat(1, nC2(k))
                    End If
                Next k
            End If
        End If
    Next
    
End Function

'--------------------------------- ArrayDB_AutoIndex
Public Function ArrayDB_AutoIndex(Arry As Variant, nColumn&, Optional IniRow As Integer = 2)
    Dim i&, n&, k%
    n = UBound(Arry, 1)
    If IniRow = 2 Then
        k = 1
    ElseIf IniRow = 1 Then
        k = 0
    End If
    For i = IniRow To n
        Arry(i, nColumn) = i - k
    Next i
    'ArrayDB_AutoIndex = Arry
End Function

'--------------------------------- ArrayDB_FilterByColumn
Public Function ArrayDB_FilterByColumn(Data As Variant, Hdr As Variant, ColumnName As String, ArryConditionsOk As Variant)
    Dim n%, m%, i%, j%, k%, c%, nCol%, ColumnValue As Variant, DatNew As Variant
    
    n = UBound(Data, 1)
    m = UBound(Data, 2)
    nCol = Lib.Array_HeaderColumnNumber(Hdr, ColumnName)
    
    k = 0
    For i = 1 To n
        ColumnValue = Data(i, nCol)
        For c = LBound(ArryConditionsOk) To UBound(ArryConditionsOk)
            If ColumnValue = ArryConditionsOk(c) Then
                k = k + 1
                Exit For
            End If
        Next c
    Next i
    
    If k > 0 Then
        DatNew = Lib.Create_EmptyArray_FromSheet(k, m)
        k = 0
        For i = 1 To n
            ColumnValue = Data(i, nCol)
            For c = LBound(ArryConditionsOk) To UBound(ArryConditionsOk)
                If ColumnValue = ArryConditionsOk(c) Then
                    k = k + 1
                    For j = 1 To m
                        DatNew(k, j) = Data(i, j)
                    Next j
                    Exit For
                End If
            Next c
        Next i
    End If
    ArrayDB_FilterByColumn = DatNew
End Function

'--------------------------------- ArrayDB_ImportFields
Public Function ArrayDB_ImportFields(Header2D As Variant, Data2D As Variant, LstFields$, _
                                            Optional XlsFileName$ = "File", _
                                            Optional TestHdrs As Boolean = True, _
                                            Optional MsgCampos$ = "")
    Dim nCol, Data, TestHeader$  ', LstImportFields$, LstFormatFields$
    'LstImportFields = LstFields & "_Import"
    'LstFormatFields = LstFields & "_Format"

    '=== Verificar que no hay nombres de columna a importar repetidos (Reporta los que estan en blanco y los repetidos a importar)
    If TestHdrs Then
        TestHeader = ArrayDB_xImport_TestHdr(Header2D, LstFields, XlsFileName)
    End If

    '=== Checar los campos a importar y su posición (No se permiten Encabezados repetidos)
    nCol = ArrayDB_xImport_GetColumns(Header2D, LstFields, XlsFileName, MsgCampos)

    '=== Armar la tabla con los campos configurados
    Data = ArrayDB_xImport_Data(Data2D, LstFields, nCol)

    ArrayDB_ImportFields = Data
End Function

'--------------------------------- ArrayDB_Join
Public Function ArrayDB_Join(Arrys As Variant, Optional OptBase% = 1, Optional IncludeHeader As Boolean = True) As Variant
    Dim Hdr As Variant, Dat As Variant, Ret As Variant
    Hdr = Lib.Array_RowExtract(Arrys(0))
    Dat = Array_Join2D(Arrys:=Arrys, OptBase:=OptBase, RemoveRow1:=True)
    If IncludeHeader Then
        Ret = Array_Join2D(Arrys:=Array(Hdr, Dat), OptBase:=OptBase, RemoveRow1:=True)
    Else
        Ret = Dat
    End If
    ArrayDB_Join = Ret
End Function

'--------------------------------- ArrayDB_ToSheet
Public Function ArrayDB_ToSheet(ShtName$, Optional Header As Variant = Empty, Optional Data As Variant = Empty, _
        Optional DBHeaderAndData As Variant = Empty, _
        Optional FilaIni& = 1, _
        Optional ColumnaIni& = 1, _
        Optional AutoFit As Boolean = False, _
        Optional AutoFilter As Boolean = False, _
        Optional ProtectSheet As Boolean = False, _
        Optional DefName As Boolean = True, _
        Optional ColorHeader As Boolean = False, _
        Optional RGBColor& = 15783870, _
        Optional sLstFmtColFecha$ = "", _
        Optional FecFormat$ = "dd/mm/yyyy", _
        Optional sLstFmtColDouble$ = "", _
        Optional NumFormat$ = "Comma", _
        Optional sLstFmtAlign$ = "", _
        Optional hAlign As Excel.Constants = xlGeneral, _
        Optional vAlign As Excel.Constants = xlBottom, _
        Optional sLstColNames1$ = "", _
        Optional Fmt1$ = "", _
        Optional sLstColNames2$ = "", _
        Optional Fmt2$ = "", _
        Optional sLstColNames3$ = "", _
        Optional Fmt3$ = "")
    '============ Borrar todos los datos anteriores de la hoja ============   (Gris = 13158600)  (Azul = 15783870)

    Dim n&, m&, m1&, Sht As Worksheet, Rg As Range, RgHeader As Range, RgData As Range, RgName As Name, StrRango$
    Set Sht = ThisWorkbook.Sheets(ShtName)
    Sht.Unprotect

    If Not IsEmpty(DBHeaderAndData) Then
        Header = Me.Array_RowExtract(DBHeaderAndData, 1)
        Data = Me.Array_DeleteFirstRows(DBHeaderAndData, 1)
    End If


    Set Rg = Sht.UsedRange
    Set RgHeader = Rg.Rows(1)

    '=== Comparar el header de entrada con el Header de la hoja acutal, si difiere borra en la hoja y pone el header de entrada ===
    Dim ShtHdr, MsgComp$
    ShtHdr = RgHeader.Value
    n = UBound(ShtHdr, 1)
    m = UBound(ShtHdr, 2)

    If Array_Compare(ShtHdr, Header, 1, n, 1, m, MsgComp) Then
        'esta bien el header de la hoja
    Else
        RgHeader.ClearContents

        Set RgHeader = RgHeader.Cells(1, 1).Resize(1, UBound(Header, 2))
        RgHeader.Value = Header
        RgHeader.WrapText = False
    End If

    '=== si el conteo de filas es mas de 1 (tiene header y datos) borra el contenido de los datos ===
    If Rg.Rows.Count > 1 Then
        Set RgData = Range(Rg.Rows(2), Rg.Rows(Rg.Rows.Count))
        RgData.ClearContents
    End If




    n = UBound(Data, 1)
    m = UBound(Data, 2)

    m1 = UBound(Header, 2)

    If m1 <> m Then
        Dim MsgDifCol$
        MsgDifCol = "No coinciden las columnas del Header con la Data..."
        MsgBox MsgDifCol
        Stop
    End If


    If n > 1 Then
        '============ Poner los datos en la hoja ============
'        Set RgHeader = Sht.Cells(FilaIni, ColumnaIni).Resize(1, m)
'        RgHeader.Value = Header
'        RgHeader.WrapText = False

        Set RgData = Sht.Cells(FilaIni + 1, ColumnaIni).Resize(n, m)
        RgData.Value = Data
        RgData.WrapText = False

        If AutoFit = True Then
            RgHeader.Columns.EntireColumn.AutoFit
        End If


        'If AutoFilter Then RgHeader.AutoFilter

        '============ Generar el "Nombre" de rango ============
        If DefName Then
            For Each RgName In ThisWorkbook.Names
                If RgName.Name = "DB" & ShtName Then
                    RgName.Delete
                    Exit For
                End If
            Next
            StrRango = "=OFFSET(" & ShtName & "!R1C1,0,0,COUNTA(" & "" & ShtName & "!C1),COUNTA(" & "" & ShtName & "!R1))"
            ThisWorkbook.Names.Add Name:="DB" & ShtName, RefersToR1C1:=StrRango
        End If


        '============ Establecer el color del Header ============
        If ColorHeader = True Then
            Set RgHeader = Rg.Rows(1)
            RgHeader.Interior.Pattern = xlSolid
            RgHeader.Interior.PatternColorIndex = xlAutomatic
            RgHeader.Interior.Color = RGBColor
            RgHeader.Interior.TintAndShade = 0
            RgHeader.Interior.PatternTintAndShade = 0
        End If

        '============ Formatear las columnas de fecha (Solo cuando sLstFmtColFecha <> "" ) ============
        If sLstFmtColFecha <> "" Then
            Call Format_Columns(ShtName, sLstFmtColFecha, FecFormat)
        End If

        '============ Formatear las columnas de numericas con coma (Solo cuando sLstFmtColDouble <> "" ) ============
        If sLstFmtColDouble <> "" Then
            Call Format_Columns(ShtName, sLstFmtColDouble, NumFormat)
        End If

        '============ Formatear las columnas con alineación (Solo cuando sLstFmtAlign <> "" ) ============
        If sLstFmtAlign <> "" Then
            Call Format_AlignColumns(ShtName, sLstFmtAlign, hAlign, vAlign)
        End If

        If sLstColNames1 <> "" Then
            Call Format_Columns(ShtName, sLstColNames1, Fmt1)
        End If

        If sLstColNames2 <> "" Then
            Call Format_Columns(ShtName, sLstColNames2, Fmt2)
        End If

        If sLstColNames3 <> "" Then
            Call Format_Columns(ShtName, sLstColNames3, Fmt3)
        End If


    End If
    If AutoFilter Then
        If Sht.AutoFilterMode Then
            'RgHeader.AutoFilter
        Else
            RgHeader.AutoFilter
        End If

    End If

    '============ Proteger la Hoja ============
    If ProtectSheet Then
        Sht.Protect
    End If

End Function

'--------------------------------- ArrayDB_ToSheetSimple
Public Function ArrayDB_ToSheetSimple(Arry2D, ShtName, Optional ActualizarDatos As Boolean = True, Optional ActualizarHeader As Boolean = False, _
                                        Optional FilaHdr& = 1, Optional ColumnaIni& = 1, Optional Proteger As Boolean = False)
    Dim Data, Header, Sht As Worksheet, n&, m&, Rg As Range

    Set Sht = ThisWorkbook.Sheets(ShtName)
    Sht.Unprotect Password:=""

    If ActualizarDatos Then
        Data = Me.Array_DeleteFirstRows(Arry2D, 1)
        n = UBound(Data, 1)
        m = UBound(Data, 2)
        Set Rg = Sht.Cells(FilaHdr + 1, ColumnaIni).Resize(n, m)
        Rg.Value = Data
    End If

    If ActualizarHeader Then
        Header = Me.Array_RowExtract(Arry2D, 1)
        n = 1
        m = UBound(Header, 2)
        Set Rg = Sht.Cells(FilaHdr, ColumnaIni).Resize(1, m)
        Rg.Value = Header
    End If

    Sht.Unprotect Password:=""

    If Proteger Then
        Sht.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=""
    End If

End Function

'--------------------------------- ArrayDB_xImport_Data
Public Function ArrayDB_xImport_Data(DataRd As Variant, LstFields$, nCol As Variant, Optional FillIndex As Boolean = False, Optional nIdxCol% = 1)
    Dim Fields, Formato, n&, m%, i&, j%, nC%, Fmt$, TestError As Boolean, MsgData$
    '=== Corregir los errores de los datos en la tabla que se leyó ===

    '=== Armar la tabla con los campos configurados
    Fields = Cfg.Range(LstFields)
    Formato = Cfg.Range(LstFields & "_Format")
    n = UBound(DataRd)
    m = UBound(Fields)
    If n < 1 Then
        ReDim Data(0 To 0, 1 To m)
    Else
        For i = 1 To n
            For j = 1 To m
                If nCol(j) > 0 Then
                    TestError = IsError(DataRd(i, nCol(j)))
                    If TestError Then
                        MsgData = MsgData & "DataRd error en: (" & i & " , " & j & ")" & vbCrLf
                        DataRd(i, nCol(j)) = CStr(DataRd(i, nCol(j)))
                    End If
                End If
            Next j
        Next i

        ReDim Data(1 To n, 1 To m)
'        '=== Encabezados configurados ===
'        For j = 1 To m
'            Data(1, j) = Fields(j, 1)
'        Next j
        If n >= 1 Then
            '=== Datos ===
            For j = 1 To m
                nC = nCol(j)
                If nC = -9 Then
                    
                    For i = 1 To n
                        Data(i, j) = i
                    Next i
                    
                ElseIf nC > 0 Then

                    Fmt = LCase(Formato(j, 1))

                    Select Case Fmt

                        Case "string", "s"

                            For i = 1 To n
                                Data(i, j) = "" & CStr(DataRd(i, nC))
                            Next i

                        Case "date"

                            For i = 1 To n
                                If IsDate(DataRd(i, nC)) Then
                                    Data(i, j) = CDate(DataRd(i, nC))
                                Else
                                    Data(i, j) = DataRd(i, nC)
                                End If
                            Next i

                        Case "double"

                            For i = 1 To n
                                If IsNumeric(DataRd(i, nC)) Then
                                    Data(i, j) = CDbl(DataRd(i, nC))
                                Else
                                    Data(i, j) = DataRd(i, nC)
                                End If
                            Next i

                        Case "integer"

                            For i = 1 To n
                                If IsNumeric(DataRd(i, nC)) Then
                                    Data(i, j) = CInt(DataRd(i, nC))
                                Else
                                    Data(i, j) = DataRd(i, nC)
                                End If
                            Next i

                        Case "nss"

                            For i = 1 To n
                                Data(i, j) = Format(DataRd(i, nC), "00-####-####-#")
                            Next i

                    End Select

                End If
            Next j
        End If
    End If
    ArrayDB_xImport_Data = Data
End Function

'--------------------------------- ArrayDB_xImport_GetColumns
Public Function ArrayDB_xImport_GetColumns(DataRd, LstFields$, Optional XlsFile$, Optional Msg$) As Variant
    Dim Fields, ImportFields, Msg1$, n%, m%, i%, Elem$, nCol As Variant, Cfg  As Worksheet
    Set Cfg = ThisWorkbook.Sheets("Cfg")
    Fields = Cfg.Range(LstFields)
    ImportFields = Cfg.Range(LstFields & "_Import")
    Msg1 = ""
    n = UBound(Fields)
    m = UBound(ImportFields)

    If n <> m Then
        Msg1 = "[" & XlsFile & "]" & vbCrLf & vbCrLf & "Num de encabezados a importar No corresponde..."
        Msg = Msg1
        'MsgBox Msg1, vbInformation
        ReDim nCol(0 To 0)
        'Stop
    Else
        ReDim nCol(1 To n)
        For i = 1 To n
            Elem = ImportFields(i, 1)
            nCol(i) = ArrayDB_xImport_GetPosition(DataRd, Elem)
            If nCol(i) = -1 Then
                Msg1 = Msg1 & "  > " & Elem & vbCrLf
            End If
        Next i
        If Msg1 <> "" Then
            Msg1 = "[" & XlsFile & "]" & vbCrLf & vbCrLf & "Falta(n) Encabezado(s):" & vbCrLf & Msg1
            Msg = Msg1
            'MsgBox Msg1, vbInformation
        End If
    End If
    ArrayDB_xImport_GetColumns = nCol
End Function

'--------------------------------- ArrayDB_xImport_GetPosition
Public Function ArrayDB_xImport_GetPosition(Arry2D, Elem)
    Dim Ret%, j%
    If LCase(Elem) = "nodefinido" Then
        Ret = 0
    ElseIf LCase(Elem) = "autoindice" Then
        Ret = -9
    Else
        Ret = -1
        For j = LBound(Arry2D, 2) To UBound(Arry2D, 2)
            If Elem = Arry2D(1, j) Then
                Ret = j
                Exit For
            End If
        Next j
    End If
    ArrayDB_xImport_GetPosition = Ret
End Function

'--------------------------------- ArrayDB_xImport_TestHdr
Public Function ArrayDB_xImport_TestHdr(DataRd, LstFields$, XlsFile$, Optional ReportBlkHdrs As Boolean = False) As String
    Dim m&, i&, j&, x&, Repetido As Boolean, Msg$, Hdr1$, Hdr2$, k&, FieldsImport
    Msg = ""
    k = 0
    FieldsImport = Cfg.Range(LstFields & "_Import")
    m = UBound(DataRd, 2)
    For i = 1 To m - 1
        Hdr1 = DataRd(1, i)
        If Hdr1 = "" Then
            k = k + 1
        Else
            For j = i + 1 To m
                Hdr2 = DataRd(1, j)
                If Hdr1 = Hdr2 Then
                    For x = 1 To UBound(FieldsImport)
                        If Hdr1 = FieldsImport(x, 1) Then
                            Msg = Msg & "[" & Hdr1 & "]" & vbCrLf
                            Exit For
                        End If
                    Next x
                End If
            Next j
        End If
    Next i

    If Msg <> "" Then
        Msg = "Encabezados que se repiten:" & vbCrLf & Msg & vbCrLf & vbCrLf
    End If

    If k > 0 And ReportBlkHdrs = True Then
        Msg = Msg & "Encabezados en blanco [" & k & "]" & vbCrLf & vbCrLf
    End If

    If Msg <> "" Then
        Msg = "[" & XlsFile & "]" & vbCrLf & vbCrLf & Msg
        'MsgBox Msg, vbInformation
    End If

    ArrayDB_xImport_TestHdr = Msg
End Function

'--------------------------------- Collection_Exists
Public Function Collection_Exists(Key$, coll As Collection) As Boolean
    'NET>> If coll.Contains(Key) Then
    'NET>>     Return True
    'NET>> Else
    'NET>>     Return False
    'NET>> End If
    Dim Itm As Variant
    On Error GoTo EH
    Set Itm = coll.Item(Key)
    Collection_Exists = True
    Exit Function
EH:
    Collection_Exists = False
End Function

'--------------------------------- Config_Actualizar
Public Function Config_Actualizar(Optional ShtName$ = "Cfg", Optional UpdateCode As Boolean = True)
    ' ------ Las dos primeras columnas se definen como variables simples a partir de la fila 2
    Dim Wbk As Workbook, Sht As Worksheet, Arry, n%, m%
    Set Wbk = ThisWorkbook
    Set Sht = Wbk.Sheets("Cfg")
    Arry = Sht.UsedRange
    n = UBound(Arry, 1)
    m = UBound(Arry, 2)

    Dim RgNam, i%, VarName$, RgName$
    '---- Definir todos los nombres con prefijo "Cfg_"
    For Each RgNam In Wbk.Names
        If Left(RgNam.Name, 4) = "Cfg_" Then
            RgNam.Delete
        End If
    Next

    For i = 2 To n
        VarName = Arry(i, 1)
        If VarName <> "" Then
            RgName = "Cfg_" & VarName
            Wbk.Names.Add Name:=RgName, RefersToR1C1:="=Cfg!R" & i & "C2"
        End If
    Next i

    '---- Definir todos los nombres con prefijo "Lst_"
    For Each RgNam In Wbk.Names
        If Left(RgNam.Name, 4) = "Lst_" Then
            RgNam.Delete
        End If
    Next

    Dim j%, StrRango$, CodeStr$
    For j = 3 To m
        VarName = Arry(1, j)
        If VarName <> "" Then
            RgName = "Lst_" & VarName
            StrRango = "=OFFSET(Cfg!R2C" & j & ",0,0,COUNTA(" & "Cfg!C" & j & ") - 1,1)"
            'Debug.Print RgName
            Wbk.Names.Add Name:=RgName, RefersToR1C1:=StrRango
        End If
    Next j
    If UpdateCode Then
        CodeStr = Config_GenClassCode()
        Call VBA_CrearModulo(Wbk:=ThisWorkbook, VbModuleName:="clsConfig", StrContenidoModulo:=CodeStr, CodeType:=vbext_ct_ClassModule)
        MsgBox "Se actualizó el codigo de [clsConfig] y los nombres de la hoja [Cfg]"
    Else
        MsgBox "Se actualizaron los nombres de la hoja [Cfg]"
    End If
End Function

'--------------------------------- Config_FullFileVariable
Public Function Config_FullFileVariable(VarName$) As SeparaPathFile
    Dim FullFile$, dFile As SeparaPathFile
    FullFile = Config_Variable(VarName)
    dFile = OS_GetFile2(FullFile)
    Config_FullFileVariable = dFile
End Function

'--------------------------------- Config_GenClassCode (Private)
Private Function Config_GenClassCode()
    Dim Arry2, nF%, Rg As Range, Arry, strCode$
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.Sheets("Cfg")
    Arry2 = Sht.UsedRange
    nF = Sht.UsedRange.Rows.Count
    Set Rg = Sht.Cells(2, 1).Resize(nF, 2)
    Arry = Rg.Value
    strCode = "Option Explicit"
    strCode = strCode & vbCrLf & "!======= Variables de configuracion ======="

    Dim i%, CfgName$, CfgType$
    For i = 1 To nF - 1
        CfgName = Arry(i, 1)
        If CfgName <> "" Then
            CfgType = Config_TipoVariable(Arry(i, 1), Arry(i, 2))
            strCode = strCode & vbCrLf & "Public " & CfgName & " As " & CfgType
        End If
    Next i

    Dim j%, LstName$
    strCode = strCode & vbCrLf
    For j = 3 To UBound(Arry2, 2)
        If Arry2(1, j) <> "" Then
            LstName = "Lst_" & Arry2(1, j)
            strCode = strCode & vbCrLf & "Public " & LstName & " As Variant"
        End If

    Next j
    strCode = strCode & vbCrLf & "!=================================="
    strCode = strCode & vbCrLf & "Private Sub Class_Initialize()"
    strCode = strCode & vbCrLf & "    Dim Cfg As Worksheet"
    strCode = strCode & vbCrLf & "    Set Cfg = ThisWorkbook.Sheets('Cfg')"

    For i = 1 To nF - 1
        CfgName = Arry(i, 1)
        If CfgName <> "" Then
            CfgType = Config_TipoVariable(Arry(i, 1), Arry(i, 2))
            strCode = strCode & vbCrLf & "    " & CfgName & " = Cfg.Range('Cfg_" & CfgName & "').Value"
        End If
    Next i
    strCode = strCode & vbCrLf

    For j = 3 To UBound(Arry2, 2)
        If Arry2(1, j) <> "" Then
            LstName = "Lst_" & Arry2(1, j)
            strCode = strCode & vbCrLf & "    " & LstName & " = cfg.Range('" & LstName & "').Value"
        End If
    Next j
    strCode = strCode & vbCrLf & "End Sub"

    strCode = strCode & vbCrLf & "!=================================="
    strCode = strCode & vbCrLf & "Public Sub UpdateVariable(CfgName$, Valor)"
    strCode = strCode & vbCrLf & "    Dim Cfg As Worksheet"
    strCode = strCode & vbCrLf & "    Set Cfg = ThisWorkbook.Sheets('Cfg')"
    strCode = strCode & vbCrLf & "    Select Case CfgName"

    Dim n%, Sp$
    For i = 1 To nF - 1
        CfgName = Arry(i, 1)
        If CfgName <> "" Then
            n = 30 - Len(CfgName)
            Sp = String(n, " ")
            strCode = strCode & vbCrLf & "        Case '" & CfgName & "':" & Sp & "Cfg.Range('Cfg_" & CfgName & "').Value = Valor"
        End If
    Next i
    strCode = strCode & vbCrLf & "        Case Else: Msgbox 'Cfg Variable no existe'"
    strCode = strCode & vbCrLf & "    End Select"
    strCode = strCode & vbCrLf & "End Sub"
    strCode = strCode & vbCrLf & "!=================================="
    strCode = strCode & vbCrLf & "Public Sub UpdateToExcel()"
    strCode = strCode & vbCrLf & "    Dim Cfg As Worksheet"
    strCode = strCode & vbCrLf & "    Set Cfg = ThisWorkbook.Sheets('Cfg')"

    For i = 1 To nF - 1
        CfgName = Arry(i, 1)
        If CfgName <> "" Then
            n = 30 - Len(CfgName)
            Sp = String(n, " ")
            strCode = strCode & vbCrLf & "    Cfg.Range('Cfg_" & CfgName & "').Value = Me." & CfgName
        End If
    Next i



    strCode = strCode & vbCrLf & "    Cfg.UsedRange.WrapText = False"
    strCode = strCode & vbCrLf & "End Sub"

    strCode = Replace(strCode, "'", Chr(34))
    strCode = Replace(strCode, "!", "'")

    Config_GenClassCode = strCode
End Function
'--------------------------------- Config_List
Public Function Config_List(LstName$, _
                            Optional CfgName$ = "Cfg", _
                            Optional Transpuesta As Boolean = False, _
                            Optional Vector1D As Boolean = False, _
                            Optional Base0 As Boolean = False)
    Dim Var As Variant, Sht As Worksheet, k%, i%, Tmp As Variant
    Set Sht = ThisWorkbook.Sheets(CfgName)
    Var = Sht.Range(LstName)
    If Vector1D Then
        If Base0 Then
            ReDim Tmp(0 To UBound(Var) - 1)
            k = -1
        Else
            ReDim Tmp(0 To UBound(Var))
            k = 0
        End If
        For i = LBound(Var) To UBound(Var)
            k = k + 1
            Tmp(k) = Var(i, 1)
        Next i
        Var = Tmp
        Transpuesta = False
    End If
    
    If Transpuesta Then
        Var = Array_Transpose(Var)
    End If
    
    
    Config_List = Var
End Function

'--------------------------------- Config_ListBuscarEquivalente
Public Function Config_ListBuscarEquivalente(Element$, Lst1_Name$, Lst2_Name$) As String
    Dim i As Integer, Idx As Integer, Lst1, Lst2
    Lst1 = Me.Config_List(Lst1_Name)
    Lst2 = Me.Config_List(Lst2_Name)
    Idx = -1
    For i = 1 To UBound(Lst1)
        If Element = Lst1(i, 1) Then
            Idx = i
            Exit For
        End If
    Next i
    If Idx > -1 Then
        Config_ListBuscarEquivalente = Lst2(Idx, 1)
    Else
        Config_ListBuscarEquivalente = "Dnf"
    End If
End Function

'--------------------------------- Config_ListDateFields
Public Function Config_ListDateFields(LstFields$)
    Dim Fields, ImportFields, ImportFormat, Msg$, Fmt$
    Fields = ThisWorkbook.Sheets("Cfg").Range(LstFields)
    ImportFields = ThisWorkbook.Sheets("Cfg").Range(LstFields & "_Import")
    ImportFormat = ThisWorkbook.Sheets("Cfg").Range(LstFields & "_Format")
    Msg = ""

    Dim n1%, n2%, n3%, i%, Elem$, nCol As Variant, Ret$
    n1 = UBound(Fields)
    n2 = UBound(ImportFields)
    n3 = UBound(ImportFormat)

    If n1 <> n2 Or n1 <> n3 Then
        Msg = "Lista en Cfg [" & LstFields & "]" & vbCrLf & vbCrLf & "Num de encabezados a importar No corresponde con Formatos..."
        MsgBox Msg, vbInformation
    Else
        ReDim nCol(1 To n1)
        For i = 1 To n1
            Fmt = LCase(ImportFormat(i, 1))
            If Fmt = "date" Then
                Ret = Ret & "|" & Fields(i, 1)
            End If
        Next i
    End If
    Config_ListDateFields = Ret
End Function

'--------------------------------- Config_ListEquivalent
Public Function Config_ListEquivalent(Element As String, Lst1, Lst2)
    Dim i As Integer, Ret As String
    For i = 1 To UBound(Lst1)
        If Element = Lst1(i) Then
            Ret = Lst2(i)
            Exit For
        End If
    Next i
    Config_ListEquivalent = Ret
End Function

'--------------------------------- Config_SetVariable
Public Function Config_SetVariable(VarName$, Valor As Variant)
    Cfg.Range("Cfg_" & VarName).Value = Valor
End Function

'--------------------------------- Config_TipoVariable (Private)
Private Function Config_TipoVariable(Variable, Valor) As String
    Dim TipoDato As VbVarType
    If LCase(Left(Variable, 3)) = "fec" Or LCase(Left(Variable, 2)) = "f_" Then
        Config_TipoVariable = "Date"
    Else
        TipoDato = VarType(Valor)
        Select Case TipoDato
            Case vbDate:            Config_TipoVariable = "Date"
            Case vbDouble:          Config_TipoVariable = "Double"
            Case vbBoolean:         Config_TipoVariable = "Boolean"
            Case Else:              Config_TipoVariable = "String"
        End Select
    End If

End Function

'--------------------------------- Config_Variable
Public Function Config_Variable(VarName$, Optional CfgName$ = "Cfg")
    Dim Var As Variant, Sht As Worksheet
    Set Sht = ThisWorkbook.Sheets(CfgName)
    Var = Sht.Range("Cfg_" & VarName)
    Config_Variable = Var
End Function

'--------------------------------- Create_EmptyArray_FromSheet
Public Function Create_EmptyArray_FromSheet(Filas, Columnas)
    Dim Rg As Range, Dat As Variant, Sht As Worksheet
    Set Sht = ThisWorkbook.Sheets("EmptySht")
    Set Rg = Sht.Cells(1, 1).Resize(Filas, Columnas)
    Dat = Rg.Value2
    Create_EmptyArray_FromSheet = Dat
End Function

'--------------------------------- Create_IdxColl_ForDB
Public Function Create_IdxColl_ForDB(Arry2D As Variant, ColumName$, ShtName$, Optional MsgDup$ = "")
    Dim Header As Variant, NumCol%, j%, Key$, i&, Msg$
    NumCol = 0
    For j = LBound(Arry2D, 2) To UBound(Arry2D, 2)
        If Arry2D(1, j) = ColumName Then
            NumCol = j
            Exit For
        End If
    Next j
    Dim IdxColl As New Collection
    Header = Empty
    If NumCol > 0 Then
        For i = LBound(Arry2D, 1) + 1 To UBound(Arry2D, 1)
            Key = "" & Arry2D(i, NumCol)
            If Not Lib.Collection_Exists(Key, IdxColl) Then
                Call IdxColl.Add(i, Key)
            Else
                Msg = Msg & vbCrLf & "Clave duplicada en fila de excel " & i & ", Key " & Key & vbCrLf & "Registro ignorado en el indice para la DB [" & ShtName & "] / [" & ColumName & "]"
                'Call MsgBox("Clave duplicada en fila de excel " & i & ", Key " & Key & vbCrLf & "Registro ignorado en el indice para la DB [" & ShtName & "] / [" & ColumName & "]")
            End If
        Next i
    End If
    MsgDup = Msg
    Set Create_IdxColl_ForDB = IdxColl
End Function

'--------------------------------- DataSheet_ChecarHayDatos
Public Function DataSheet_ChecarHayDatos(ShtName) As Boolean
    Dim Rg As Range, Data, n As Integer
    Set Rg = ThisWorkbook.Sheets(ShtName).Range("DB" & ShtName)
    n = Rg.Rows.Count
    If n > 1 Then
        DataSheet_ChecarHayDatos = True
    Else
        DataSheet_ChecarHayDatos = False
    End If
End Function

'--------------------------------- DataSheet_DefineDBName
Public Function DataSheet_DefineDBName(Wbk As Workbook, ShtName$)
    Dim Sht As Worksheet, RgName$, RgN As Name, StrRango$
    Set Sht = Wbk.Sheets(ShtName)
    RgName = "DB" & ShtName
    For Each RgN In ThisWorkbook.Names
        If RgN.Name = RgName Then
            RgN.Delete
        End If
    Next
    StrRango = "=OFFSET(ShtName!R1C1,0,0,COUNTA(ShtName!C1),COUNTA(ShtName!R1))"
    StrRango = Replace(StrRango, "ShtName", ShtName)
    Wbk.Names.Add Name:=RgName, RefersToR1C1:=StrRango
End Function

'--------------------------------- DataSheet_DeleteData
Public Function DataSheet_DeleteData(ShtName$, Optional CreateDBname As Boolean = False, Optional FilaDatos% = 2, Optional DeleteRows As Boolean = False, Optional DeleteUsedRgRows As Boolean = True)
    Dim Sht As Worksheet, Rg As Range, nF&, nC&, F1&, c1&, RgData As Range, RgName$
    RgName = "DB" & ShtName
    Set Sht = ThisWorkbook.Worksheets(ShtName)
    If CreateDBname Then
        Call DataSheet_DefineDBName(Sht.Parent, ShtName)
    End If
    Set Rg = Sht.Range(RgName)
    nC = Rg.Columns.Count
    If DeleteUsedRgRows Then
        nF = Rg.Worksheet.UsedRange.Rows.Count
        nC = Rg.Worksheet.UsedRange.Columns.Count
    Else
        nF = Rg.Rows.Count - 1
        nC = Rg.Columns.Count
    End If
    If nF > 1 And nC > 0 Then
        F1 = Rg.Cells(1, 1).Row + 1    ' o tambien la fila de datos  (f1 = FilaDatos)
        c1 = Rg.Cells(1, 1).Column
        If DeleteUsedRgRows Then
            Set RgData = Rg.Worksheet.Cells(F1, c1).Resize(nF, nC)
        Else
            Set RgData = Rg.Offset(1, 0).Resize(nF, nC)
        End If
        Sht.Unprotect
        If DeleteRows Then
            RgData.Rows.EntireRow.Delete
        Else
            RgData.ClearContents
        End If
        Sht.Protect
    End If
End Function

'--------------------------------- DataSheet_FilterCancel
Public Function DataSheet_FilterCancel(ShtName$)
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.Sheets(ShtName)
    If Sht.FilterMode Then
        Sht.Unprotect Password:=""
        On Error Resume Next
        Sht.ShowAllData
        On Error GoTo 0
        Sht.Protect
    End If
End Function

'--------------------------------- FE_DateToExcelData
Public Function FE_DateToExcelData(Fecha As Date) As Variant
    Dim F1 As Date
    If Fecha = F1 Then
        FE_DateToExcelData = ""
    Else
        FE_DateToExcelData = Fecha
    End If
End Function

'--------------------------------- FE_DateToStr
Public Function FE_DateToStr(Fecha As Date) As String
    Dim F1 As Date
    If F1 = Fecha Then
        FE_DateToStr = ""
    Else
        FE_DateToStr = Format(Fecha, "dd/mm/yyyy")
    End If
End Function

'--------------------------------- FE_GetDate
Public Function FE_GetDate(Fecha As Variant) As Date
    Dim F1 As Date
    If IsEmpty(Fecha) Then
        FE_GetDate = F1
    End If
    If VarType(Fecha) = 7 Then
        FE_GetDate = Fecha
    Else
        FE_GetDate = F1
    End If
End Function

'--------------------------------- Find_InArrayDB
Public Function Find_InArrayDB(Id$, ArrayDB As Variant, IdxColl As Collection)
    Dim m%, RowDat As Variant, i&, j%
    m = UBound(ArrayDB, 2) - LBound(ArrayDB, 2) + 1
    RowDat = Lib.Create_EmptyArray_FromSheet(Filas:=1, Columnas:=m)
    i = Find_Index(Id, IdxColl)
    If i > 0 Then
        For j = 1 To m
            RowDat(1, j) = ArrayDB(i, j)
        Next j
    End If
    Find_InArrayDB = RowDat
End Function

'==================================
Private Function Find_Index(Clave$, IdxColl As Collection) As Long
    Dim Fnd As Boolean, i&
    i = -1
    Fnd = Lib.Collection_Exists(Clave, IdxColl)
    If Fnd Then
        i = IdxColl.Item(Clave)
    End If
    Find_Index = i
End Function
'===========================================================================================
'  Procedimiento para formatear las columnas en un rango tipo DB "DBShtName"
'  en lo que se refiere a la alineación vertical y horizontal de las celdas
'      hAlign = <xlGeneral>, xlLeft, xlRight, xlCenter
'      vAlign = <xlCenter> , xlTop, xlBottom
'  La lista de columnas que se van a formatear van en un string "sLstColumnsAlign"
'  que inicia con "|" y después cada uno de los campos separados por el caracter "|"
'  p.Ej. sLstColumnsAlign = "|NombreColumna1|NombreColumna2|NombreColumna13"
'===========================================================================================
Private Function Format_AlignColumns(ShtName$, sLstColumnsAlign, Optional hAlign As Excel.Constants = xlGeneral, Optional vAlign As Excel.Constants = xlCenter)
    Dim Rg As Range, RgName$, m&, LstColumnas2Format, j&, NombreColumna$, FormatearColumna, RgColumn As Range
    RgName = "DB" & ShtName
    Set Rg = Range(RgName)
    m = Rg.Columns.Count
    If sLstColumnsAlign <> "" Then
        LstColumnas2Format = Split(sLstColumnsAlign, "|")
        For j = 1 To m
            NombreColumna = Rg.Cells(1, j)
            FormatearColumna = Me.Vector_FindElementPosition(NombreColumna, LstColumnas2Format)
            If FormatearColumna > -1 Then
                Set RgColumn = Rg.Columns(j)
                RgColumn.HorizontalAlignment = hAlign
                RgColumn.VerticalAlignment = vAlign
                RgColumn.WrapText = False
            End If
        Next j
    End If
End Function

'===========================================================================================
'  Procedimiento para formatear las columnas en un rango tipo DB "DBShtName"
'  con el formato seleccionado que pueden ser los siguientes
'    "General", "@", "dd/mm/yyyy", "0.00", "Comma"
'  La lista de columnas que se van a formatear van en un string que inicia con "|" y después
'  cada uno de los campos separados por el caracter "|"
'  p.Ej. sLstColumnNames = "|NombreColumna1|NombreColumna2|NombreColumna13"
'===========================================================================================
Private Function Format_Columns(ShtName$, sLstColumnNames$, NumFormat$)
    Dim Sht As Worksheet, Rg As Range, RgName$, m&, LstColumnas2Format, j&, NombreColumna$, FormatearColumna, RgColumn As Range
    Set Sht = ThisWorkbook.Sheets(ShtName)
    RgName = "DB" & ShtName
    Set Rg = Sht.Range(RgName)
    m = Rg.Columns.Count
    If sLstColumnNames <> "" Then
        LstColumnas2Format = Split(sLstColumnNames, "|")
        For j = 1 To m
            NombreColumna = Rg.Cells(1, j)
            FormatearColumna = Me.Vector_FindElementPosition(NombreColumna, LstColumnas2Format)
            If FormatearColumna > -1 Then
                Set RgColumn = Rg.Columns(j)
                If NumFormat = "Comma" Then
                    RgColumn.Style = "Comma"
                Else
                    RgColumn.NumberFormat = NumFormat
                End If
            End If
        Next j
    End If
End Function
'--------------------------------- GetRowForFilter
Public Function GetRowForFilter(RgName$, ColNameToFilter$, Condition As Variant, Optional UseType As String = "String")
    '============= En el rango de la base de datos recorre para determinar si el campo "ColNameToFilter" coincide con el criterio de busqueda
    Dim ShtName$, Ret As Variant, Sht As Worksheet, Hdr As Variant, ColToFilter&, RowNum&
    Dim i&, Dat As Variant, RgData As Range, DataArrayTmp As Variant, DataArray As Variant
    Ret = -1
    ShtName = Mid(RgName, 3)
    Set Sht = ThisWorkbook.Worksheets(ShtName)
    Hdr = Sht.Range(RgName).Rows(1).Value
    ColToFilter = Array_HeaderColumnNumber(Hdr, ColNameToFilter)
    RowNum = Sht.Range(RgName).Rows.Count
    Set RgData = Sht.Range(Sht.Range(RgName).Cells(2, ColToFilter), Sht.Range(RgName).Cells(RowNum, ColToFilter))
    If RowNum = 2 Then
        'ReDim DataArrayTmp(0 To 1, 0 To 1)
        DataArrayTmp = Create_EmptyArray_FromSheet(1, 1)
        DataArrayTmp(1, 1) = RgData.Value
        DataArray = DataArrayTmp
    Else
        DataArray = RgData.Value
    End If
    Ret = -1
    For i = 1 To UBound(DataArray)
        Select Case UseType
            Case "String":  Dat = "" & DataArray(i, 1)
            Case "Numeric": Dat = Val(DataArray(i, 1))
            Case Else:      Dat = DataArray(i, 1)
        End Select
        If VarType(Dat) = VarType(Condition) Then
            If Dat = Condition Then
                Ret = i + 1
                Exit For
            End If
        End If
    Next i
    GetRowForFilter = Ret
End Function

'--------------------------------- OS_CreateDirectory
Public Function OS_CreateDirectory(RootPath As String, SubFolder As String)
    Dim Path$, Chk$
    Path = RootPath & "\" & SubFolder
    Chk = Dir(Path, vbDirectory)
    If Chk = "" Then
        Call MkDir(Path)
    End If
    OS_CreateDirectory = Path
End Function

'--------------------------------- OS_DialogoAbrirArchivo
'================================================================================================================
'  Función que permite seleccionar un archivo desde un cuadro de diálogo.
'  La ruta inicial será definida por el nombre de archivo inicial o
'  si esta no existe entonces usa la ruta del archivo de excel
'================================================================================================================
Public Function OS_DialogoAbrirArchivo(FullFileName$, TituloDialogo$, DialogFilter$) As String
    Dim Fso As New FileSystemObject
    Dim PathIni$, strFile, FileName$, Folder$, FileOK As Boolean, PathOK As Boolean
    FileOK = Fso.FileExists(FullFileName)
    FileName = Fso.GetFileName(FullFileName)
    Folder = Fso.GetParentFolderName(FullFileName)
    PathOK = Fso.FolderExists(Folder)
    PathIni = ThisWorkbook.Path
    If PathOK Then
        ChDir Folder & "\"
    Else
        ChDir PathIni & "\"
    End If
    strFile = Application.GetOpenFilename(Title:=TituloDialogo, FileFilter:=DialogFilter)
    Set Fso = Nothing
    If strFile = False Then
        OS_DialogoAbrirArchivo = ""
    Else
        OS_DialogoAbrirArchivo = strFile
    End If
End Function

'--------------------------------- OS_DialogoSelectFolder
'==================================================================================
'  Permite seleccionar un folder de destino por medio de un cuadro de diálogo
'  La ruta o path inicial se puede pasar como argumento, o si no se especifica,
'  se utiliza como ruta inicial donde se encuentra el libro de excel
'  En caso de que no exista la ruta inicial tambien utiliza el path del libro
'==================================================================================
Public Function OS_DialogoSelectFolder(Optional IniPath$ = "") As String
    Dim BrwFldr As FileDialog, sItm$
    Set BrwFldr = Application.FileDialog(msoFileDialogFolderPicker)
    BrwFldr.Title = "Select a Folder"
    BrwFldr.AllowMultiSelect = False

    If IniPath = "" Then
        BrwFldr.InitialFileName = ThisWorkbook.Path    'Application.DefaultFilePath
    Else
        If Dir(IniPath, vbDirectory) <> "" Then
            BrwFldr.InitialFileName = IniPath
        Else
            BrwFldr.InitialFileName = ThisWorkbook.Path
        End If
    End If



    If BrwFldr.Show <> -1 Then
        GoTo Salir
    End If

    sItm = BrwFldr.SelectedItems(1)

Salir:
    OS_DialogoSelectFolder = sItm
    Set BrwFldr = Nothing
End Function

'--------------------------------- OS_GetFile2
Public Function OS_GetFile2(FullFileName) As SeparaPathFile
    Dim Fso As New FileSystemObject, FileOK As Boolean, FileName$, Folder$, Ret, File As File
    FileOK = Fso.FileExists(FullFileName)
    FileName = Fso.GetFileName(FullFileName)
    Set File = Fso.GetFile(FullFileName)
    If FileOK Then
        Folder = Fso.GetParentFolderName(FullFileName)
        Ret = Array(Folder, FileName)
    Else
        Ret = Array("Dnf", FileName)
    End If
    Set Fso = Nothing
    OS_GetFile2.FullFile = FullFileName
    OS_GetFile2.Path = Ret(0)
    OS_GetFile2.Name = Ret(1)
    OS_GetFile2.ModifDate = Format(File.DateLastModified, "yyyy-mm-dd (hh:mm:ss am/pm)")
End Function

'--------------------------------- OS_StrFileList
Public Function OS_StrFileList(Path_Ini$, FileName$, Optional SearchSubFolders As Boolean = False, Optional FullPath As Boolean = True)
    Dim Fso As FileSystemObject, RootFolder As Folder, Path$, FullFile$, FindFileName$, sFileLst$, SubFolder As Folder
    Set Fso = New FileSystemObject
    If Fso.FolderExists(Path_Ini) Then
        Set RootFolder = Fso.GetFolder(Path_Ini)
        FullFile = RootFolder.Path & "\" & FileName
        FindFileName = Dir(FullFile, vbArchive)
        While FindFileName <> ""
            If FullPath Then
                sFileLst = sFileLst & "|" & RootFolder.Path & "\" & FindFileName
            Else
                sFileLst = sFileLst & "|" & FindFileName
            End If
            FindFileName = Dir
        Wend
        If SearchSubFolders Then
            For Each SubFolder In RootFolder.SubFolders
                sFileLst = sFileLst & OS_StrFileList(SubFolder.Path, FileName, True)
            Next
        End If
    End If
    OS_StrFileList = sFileLst
End Function

'--------------------------------- OS_SubFolderFileListArray
Public Function OS_SubFolderFileListArray(Path_Ini$, FileName$, Optional SearchSubFolders As Boolean = False, Optional FullPath As Boolean = True)
    Dim sLstFiles$, LstFiles
    sLstFiles = OS_StrFileList(Path_Ini, FileName, SearchSubFolders, FullPath)
    If sLstFiles <> "" Then
        LstFiles = Split(sLstFiles, "|")
    Else
        ReDim LstFiles(0 To 0)
        LstFiles(0) = ""
    End If
    OS_SubFolderFileListArray = LstFiles
End Function

'--------------------------------- ReadSheet_HeaderAndData
Public Function ReadSheet_HeaderAndData(ShtNameToRead As String, _
                                Optional TieneHeader As Boolean = True, _
                                Optional DeleteEmptyRows As Boolean = True, _
                                Optional DeleteEmptyColumns As Boolean = True, _
                                Optional FilaIni As Long = 1, _
                                Optional ErrMsg As String = "", _
                                Optional DeleteEmptyRowsByOneColumn As Boolean = False, _
                                Optional nColumn As Long = -1) As ReadXlsDatShts
        
    Dim Arry As Variant, Dat As ReadXlsDatShts, Fso As New FileSystemObject
    Arry = Me.ReadSheet_ToArray(ShtNameToRead:=ShtNameToRead, _
                                DeleteEmptyRows:=DeleteEmptyRows, DeleteEmptyColumns:=DeleteEmptyColumns, FilaIni:=FilaIni, ErrMsg:=ErrMsg, _
                                DeleteEmptyRowsByOneColumn:=DeleteEmptyRowsByOneColumn, nColumn:=nColumn)
                                
    Dat.FileName = ThisWorkbook.Name
    Dat.Folder = ThisWorkbook.Path
    Dat.ShtName = ShtNameToRead
    If TieneHeader Then
        Dat.Header = Me.Array_RowExtract(Arry, 1)
        Dat.Data = Me.Array_DeleteFirstRows(Arry, 1)
    Else
        Dat.Header = Empty
        Dat.Data = Arry
    End If
    ReadSheet_HeaderAndData = Dat
End Function

'--------------------------------- ReadSheet_ImportColumns
Public Function ReadSheet_ImportColumns(ShtNameToRead As String, LstFields As String, _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni As Long = 1, Optional ErrMsg As String = "", _
        Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn As Long = -1, Optional AutoIdColumn$ = "") As ReadXlsDatShts
    
    Dim DataRd As ReadXlsDatShts, Ret As ReadXlsDatShts, DataImport, HeaderImport, cId As Integer, i As Integer
    
    DataRd = ReadSheet_HeaderAndData(ShtNameToRead, True, DeleteEmptyRows, DeleteEmptyColumns, FilaIni, ErrMsg, DeleteEmptyRowsByOneColumn, nColumn)
    DataImport = Me.ArrayDB_ImportFields(Header2D:=DataRd.Header, Data2D:=DataRd.Data, LstFields:=LstFields, XlsFileName:=DataRd.FileName, TestHdrs:=True, MsgCampos:=ErrMsg)
    'DataImport = Me.Array_DeleteEmptyRowsByColumn(DataImport, 1)
    HeaderImport = Me.Config_List(LstFields)
    HeaderImport = Me.Array_Transpose(HeaderImport)
    If AutoIdColumn <> "" Then
        cId = Me.Array_HeaderColumnNumber(HeaderImport, AutoIdColumn)
        If cId > 0 Then
            For i = 1 To UBound(DataImport, 1)
                DataImport(i, cId) = i
            Next i
        End If
    End If
    Ret.Data = DataImport
    Ret.Header = HeaderImport
    Ret.FileName = DataRd.FileName
    Ret.Folder = DataRd.Folder
    Ret.ShtName = ShtNameToRead
    
    ReadSheet_ImportColumns = Ret

End Function

'--------------------------------- ReadSheet_ImportFields
Public Function ReadSheet_ImportFields(ShtRead$, ShtDestino$) As ReadXlsDatShts
    Dim Wbk As Workbook, Tabla As Variant, Hdr As Variant, Dat As Variant, DatN As Variant, HdrN As Variant, Ret As ReadXlsDatShts
    Set Wbk = ThisWorkbook
    Tabla = Lib.ReadSheet_UsedRangeToArray(Wbk, ShtRead)
    Tabla = Lib.Array_DeleteEmptyRowsByColumn(Tabla, 1)
    Hdr = Lib.Array_RowExtract(Tabla, 1)
    Dat = Lib.Array_DeleteFirstRows(Tabla, 1)
    DatN = Lib.ArrayDB_ImportFields(Header2D:=Hdr, Data2D:=Dat, LstFields:="Lst_" & ShtDestino)
    
    
    
    
    HdrN = Lib.Config_List("Lst_" & ShtDestino, "Cfg", True)
    Ret.Data = DatN
    Ret.Header = HdrN
    Ret.FileName = "Thisworkbook"
    Ret.ShtName = ShtRead
    Ret.Folder = Wbk.Path
    ReadSheet_ImportFields = Ret
End Function

'--------------------------------- ReadSheet_NamedRangeToArray
Public Function ReadSheet_NamedRangeToArray(Wbk As Workbook, ShtName$, Optional Msg$ = "")
    Dim Sht As Worksheet, DataRd, Msg1$
    If Me.SheetExists(Wbk, ShtName) Then
        Set Sht = Wbk.Worksheets(ShtName)
        DataRd = Sht.Range("DB" & ShtName)
    Else
        ReDim DataRd(0 To 0, 0 To 0)
        Msg1 = "No se encontró la hoja [" & ShtName & "]"
    End If

    If Msg = "" Then
        Msg = Msg1
    Else
        Msg = Msg & vbCrLf & Msg1
    End If

    ReadSheet_NamedRangeToArray = DataRd
End Function

'--------------------------------- ReadSheet_ToArray
Public Function ReadSheet_ToArray(ShtNameToRead As String, _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni As Long = 1, _
        Optional ErrMsg As String = "", Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn As Long = -1, _
        Optional TrimHeader As Boolean = False, Optional TrimData As Boolean = False _
        )
    Dim Wbk As Workbook, Sht As Worksheet, DataRd, Msg1 As String, Ret, i As Long, j As Integer
    DataRd = Me.ReadSheet_UsedRangeToArray(ThisWorkbook, ShtNameToRead, ErrMsg)
    
    If FilaIni > 1 Then
        DataRd = Array_DeleteFirstRows(DataRd, FilaIni - 1)
    End If
    
    If DeleteEmptyRows Then
        If DeleteEmptyRowsByOneColumn And nColumn > -1 Then
            DataRd = Array_DeleteEmptyRowsByColumn(DataRd, nColumn)
        Else
            DataRd = Array_DeleteEmptyRows(DataRd)
        End If
    End If
    
    If DeleteEmptyColumns Then
        DataRd = Array_DeleteEmptyColumns(DataRd)
    End If
    
    If TrimHeader Then
        Call Array_TrimRow(DataRd, 1)
    End If

    If TrimData Then
        For i = 2 To UBound(DataRd)
            For j = 1 To UBound(DataRd, 2)
                If Not IsDate(DataRd(i, j)) Then
                    DataRd(i, j) = Trim(DataRd(i, j))
                End If
            Next j
        Next i
    End If

    
    ReadSheet_ToArray = DataRd
End Function

'--------------------------------- ReadSheet_UsedRangeToArray
Public Function ReadSheet_UsedRangeToArray(Wbk As Workbook, ShtName$, Optional Msg$ = "")
    Dim Sht As Worksheet, DataRd, Msg1$
    If Me.SheetExists(Wbk, ShtName) Then
        Set Sht = Wbk.Worksheets(ShtName)
        DataRd = Sht.UsedRange
    Else
        ReDim DataRd(0 To 0, 0 To 0)
        Msg1 = "No se encontró la hoja [" & ShtName & "]"
    End If

    If Msg = "" Then
        Msg = Msg1
    Else
        Msg = Msg & vbCrLf & Msg1
    End If

    ReadSheet_UsedRangeToArray = DataRd
End Function

'--------------------------------- ReadSheet_UsedRangeToArrayDB
Public Function ReadSheet_UsedRangeToArrayDB(Wbk As Workbook, ShtName$, Optional Msg$ = "") As ReadXlsDatShts
    Dim Sht As Worksheet, UsedRg, Ret As ReadXlsDatShts, Header As Variant, Data As Variant, Msg1$, NoData

    If Me.SheetExists(Wbk, ShtName) Then
        Set Sht = Wbk.Worksheets(ShtName)
        UsedRg = Sht.UsedRange
        Header = Me.Array_RowExtract(UsedRg, 1)
        Data = Me.Array_DeleteFirstRows(UsedRg, 1)
        Ret.Header = Header
        Ret.Data = Data
        Ret.ShtName = ShtName
        Ret.FileName = ThisWorkbook.Name
        Ret.Folder = ThisWorkbook.Path
    Else
        ReDim NoData(0 To 0, 0 To 0)

        Ret.Header = NoData
        Ret.Data = NoData
        Ret.ShtName = ShtName
        Ret.FileName = ThisWorkbook.Name
        Ret.Folder = ThisWorkbook.Path

        Msg1 = "No se encontró la hoja [" & ShtName & "]"
    End If

    If Msg = "" Then
        Msg = Msg1
    Else
        Msg = Msg & vbCrLf & Msg1
    End If


    ReadSheet_UsedRangeToArrayDB = Ret
End Function

'--------------------------------- ReadXlsSheet_HeaderAndData    - Retorna ReadXlsDatShts del rango en uso (UsedRange) por medio de ReadXlsSheet
Public Function ReadXlsSheet_HeaderAndData(FullFile$, _
                                Optional ShtNameToRead$ = "", _
                                Optional TieneHeader As Boolean = True, _
                                Optional WbkPass$ = "", _
                                Optional DeleteEmptyRows As Boolean = True, _
                                Optional DeleteEmptyColumns As Boolean = True, _
                                Optional FilaIni& = 1, _
                                Optional ErrMsg$ = "", _
                                Optional DeleteEmptyRowsByOneColumn As Boolean = False, _
                                Optional nColumn& = -1) As ReadXlsDatShts

    Dim Arry As Variant, Dat As ReadXlsDatShts, Fso As New FileSystemObject
    Arry = Me.ReadXlsSheet_ToArray(XlsFullName:=FullFile, ShtNameToRead:=ShtNameToRead, WbkPass:=WbkPass, _
                                DeleteEmptyRows:=DeleteEmptyRows, DeleteEmptyColumns:=DeleteEmptyColumns, FilaIni:=FilaIni, ErrMsg:=ErrMsg, _
                                DeleteEmptyRowsByOneColumn:=DeleteEmptyRowsByOneColumn, nColumn:=nColumn)

    Dat.FileName = Fso.GetFileName(FullFile)
    Dat.Folder = Fso.GetParentFolderName(FullFile)
    Dat.ShtName = ShtNameToRead
    If TieneHeader Then
        Dat.Header = Array_RowExtract(Arry, 1)
        Dat.Data = Array_DeleteFirstRows(Arry, 1)
    Else
        Dat.Header = Empty
        Dat.Data = Arry
    End If
    ReadXlsSheet_HeaderAndData = Dat
End Function

'--------------------------------- ReadXlsSheet_ImportColumns            - Retorna ReadXlsDatShts
Public Function ReadXlsSheet_ImportColumns(XlsFullName$, ShtNameToRead$, LstFields$, Optional WbkPass$ = "", _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni& = 1, Optional ErrMsg$ = "", _
        Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn& = -1, Optional AutoIdColumn$ = "") As ReadXlsDatShts

    Dim DataRd As ReadXlsDatShts, Ret As ReadXlsDatShts, DataImport, HeaderImport

    DataRd = ReadXlsSheet_HeaderAndData(XlsFullName, ShtNameToRead, True, WbkPass, DeleteEmptyRows, DeleteEmptyColumns, FilaIni, ErrMsg, DeleteEmptyRowsByOneColumn, nColumn)
    DataImport = Me.ArrayDB_ImportFields(Header2D:=DataRd.Header, Data2D:=DataRd.Data, LstFields:=LstFields, XlsFileName:=DataRd.FileName, TestHdrs:=True, MsgCampos:=ErrMsg)
    'DataImport = Me.Array_DeleteEmptyRowsByColumn(DataImport, 1)
    HeaderImport = Me.Config_List(LstFields)
    HeaderImport = Me.Array_Transpose(HeaderImport)
    If AutoIdColumn <> "" Then
        cId = Me.Array_HeaderColumnNumber(HeaderImport, AutoIdColumn)
        If cId > 0 Then
            For i = 1 To UBound(DataImport, 1)
                DataImport(i, cId) = i
            Next i
        End If
    End If
    Ret.Data = DataImport
    Ret.Header = HeaderImport
    Ret.FileName = DataRd.FileName
    Ret.Folder = DataRd.Folder
    Ret.ShtName = ShtNameToRead

    ReadXlsSheet_ImportColumns = Ret
End Function

'--------------------------------- ReadXlsSheet_ImportColumns_MultipShts        - Retorna Array ReadXlsDatShts()
Public Function ReadXlsSheet_ImportColumns_MultipShts(XlsFullName$, VectorShtNameToRead, VectorLstFields, Optional VectorWbkPass As Variant = Empty, _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni& = 1, _
        Optional ErrMsg$ = "", Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn& = -1) As ReadXlsDatShts()

    Dim DataRd As ReadXlsDatShts, Ret() As ReadXlsDatShts, DataImport, HeaderImport, n0%, n%, i%, ShtNameToRead$, WbkPass$, LstFields$
    n0 = LBound(VectorShtNameToRead)
    n = UBound(VectorShtNameToRead)
    ReDim Ret(n0 To n)
    For i = n0 To n
        ShtNameToRead = VectorShtNameToRead(i)
        LstFields = VectorLstFields(i)
        If IsEmpty(VectorWbkPass) Then
            WbkPass = ""
        Else
            WbkPass = VectorWbkPass(i)
        End If
        DataRd = Lib.ReadXlsSheet_HeaderAndData(XlsFullName, ShtNameToRead, True, WbkPass, DeleteEmptyRows, DeleteEmptyColumns, FilaIni, ErrMsg, DeleteEmptyRowsByOneColumn, nColumn)
        Ret(i) = ReadXlsSheet_DataImportFields(ShtNameToRead:=ShtNameToRead, DataRd:=DataRd, LstFields:=LstFields, TestHdrs:=True, ErrMsg:=ErrMsg)
    Next i

    ReadXlsSheet_ImportColumns_MultipShts = Ret
End Function

'--------------------------------- ReadXlsSheet_DataImportFields  (Private)
Private Function ReadXlsSheet_DataImportFields(ShtNameToRead$, DataRd As ReadXlsDatShts, LstFields$, Optional TestHdrs As Boolean = True, Optional ErrMsg$ = "") As ReadXlsDatShts
    Dim DataImport, HeaderImport, Ret As ReadXlsDatShts
    DataImport = ArrayDB_ImportFields(Header2D:=DataRd.Header, Data2D:=DataRd.Data, LstFields:=LstFields, XlsFileName:=DataRd.FileName, TestHdrs:=True, MsgCampos:=ErrMsg)
    HeaderImport = Config_List(LstFields)
    HeaderImport = Array_Transpose(HeaderImport)
    Ret.Data = DataImport
    Ret.Header = HeaderImport
    Ret.FileName = DataRd.FileName
    Ret.Folder = DataRd.Folder
    Ret.ShtName = ShtNameToRead
    ReadXlsSheet_DataImportFields = Ret
End Function

'--------------------------------- ReadXlsSheet_ImportColumns_MultipWbks2
Public Function ReadXlsSheet_ImportColumns_MultipWbks2(VectorXlsFullName As Variant, VectorShtNameToRead As Variant, VectorLstFields As Variant, _
        Optional VectorWbkPass As Variant = Empty, _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni& = 1, _
        Optional ErrMsg$ = "", Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn& = -1, _
        Optional SinglePath$ = "") As ReadXlsDatShts()
    ' Se asume VectorXlsFullName, VectorShtNameToRead, VectorLstField,VectorWbkPass Son Base 0 formados con funcion Array()
    Dim ShowPGB As Boolean, n%, n1%, Avance%, FilePath$, Ret() As ReadXlsDatShts, i%
    Dim Fso As New FileSystemObject, File As File, FechaModifc$, FileName$, XlsFullName$, ShtNameToRead$
    Dim LstFields$, WbkPass$, DataRd As ReadXlsDatShts
    n = UBound(VectorXlsFullName)
    '------ PGB ------
    ShowPGB = True
    n1 = n + 1
    Avance = 0
    Call Pgb.MensajeAvanzar(FrmProgress1, vbCrLf & "======================================================================", 0)
    Call Pgb.MensajeAvanzar(FrmProgress1, "- Importar multiple archivos desde", 0)
    If SinglePath <> "" Then
        FilePath = Fso.GetParentFolderName(VectorXlsFullName(1))
        Call Pgb.MensajeAvanzar(FrmProgress1, "   >> Ruta  [" & FilePath & "]", 0)
    End If
    '-----------------
    ReDim Ret(0 To n)
    For i = 0 To n
        XlsFullName = VectorXlsFullName(i)
        ShtNameToRead = VectorShtNameToRead(i)
        LstFields = VectorLstFields(i)
        If IsEmpty(VectorWbkPass) Then
            WbkPass = ""
        Else
            WbkPass = VectorWbkPass(i)
        End If

        '------ PGB ------
        If ShowPGB Then
            Avance = i / n1 * 1000
            Set File = Fso.GetFile(XlsFullName)
            FechaModifc = Format(File.DateLastModified, "yyyy-mm-dd (hh:mm:ss am/pm)")
            FilePath = Fso.GetParentFolderName(XlsFullName)
            FileName = Fso.GetFileName(XlsFullName)
            If SinglePath <> "" Then
                Pgb.MensajeAvanzar FrmProgress1, "        >> " & FechaModifc & "  -  " & FileName, Avance
            Else
                Pgb.MensajeAvanzar FrmProgress1, "        >> Ruta: [" & FilePath & "]", Avance
                Pgb.MensajeAvanzar FrmProgress1, "            >> " & FechaModifc & "  -  " & FileName & "(" & ShtNameToRead & ")", Avance
            End If
        End If
        '-----------------

        DataRd = ReadXlsSheet_HeaderAndData(XlsFullName, ShtNameToRead, True, WbkPass, DeleteEmptyRows, DeleteEmptyColumns, FilaIni, ErrMsg, DeleteEmptyRowsByOneColumn, nColumn)
        Ret(i) = ReadXlsSheet_DataImportFields(ShtNameToRead:=ShtNameToRead, DataRd:=DataRd, LstFields:=LstFields, TestHdrs:=True, ErrMsg:=ErrMsg)
    Next i

    ReadXlsSheet_ImportColumns_MultipWbks2 = Ret
End Function

'--------------------------------- ReadXlsSheet_ToArray
Public Function ReadXlsSheet_ToArray(XlsFullName$, Optional ShtNameToRead$ = "", Optional WbkPass$ = "", _
        Optional DeleteEmptyRows As Boolean = True, Optional DeleteEmptyColumns As Boolean = True, Optional FilaIni& = 1, _
        Optional ErrMsg$ = "", Optional DeleteEmptyRowsByOneColumn As Boolean = False, Optional nColumn& = -1, _
        Optional TrimHeader As Boolean = False, Optional TrimData As Boolean = False _
        )
    Dim DataRd As Variant, i&, j%
    DataRd = ReadXlsSheet_UsedRangeToArray(XlsFullName, ShtNameToRead, WbkPass, ErrMsg)
    If DeleteEmptyRows Then
        If DeleteEmptyRowsByOneColumn And nColumn > -1 Then
            DataRd = Array_DeleteEmptyRowsByColumn(DataRd, nColumn)
        Else
            DataRd = Array_DeleteEmptyRows(DataRd)
        End If
    End If
    If DeleteEmptyColumns Then
        DataRd = Array_DeleteEmptyColumns(DataRd)
    End If
    If FilaIni > 1 Then
        DataRd = Array_DeleteFirstRows(DataRd, FilaIni - 1)
    End If
    If TrimHeader Then
        Call Array_TrimRow(DataRd, 1)
    End If
    If TrimData Then
        For i = 2 To UBound(DataRd)
            For j = 1 To UBound(DataRd, 2)
                If Not IsDate(DataRd(i, j)) Then
                    DataRd(i, j) = Trim(DataRd(i, j))
                End If
            Next j
        Next i
    End If
    ReadXlsSheet_ToArray = DataRd
End Function

'--------------------------------- ReadXlsSheet_UsedRangeToArray
Public Function ReadXlsSheet_UsedRangeToArray(XlsFullName, Optional ShtNameToRead$ = "", Optional WbkPass$ = "", Optional Msg$ = "")
    Dim Wbk As Workbook, Sht As Worksheet, DataRd, Msg1$
    Set Wbk = Application.Workbooks.Open(FileName:=XlsFullName, UpdateLinks:=False, Password:=WbkPass)
    Application.Windows(Wbk.Name).Visible = False
    If ShtNameToRead = "" Then
        Set Sht = Wbk.Worksheets(1)
        DataRd = Sht.UsedRange
        Wbk.Close Savechanges:=False
    Else
        If SheetExists(Wbk, ShtNameToRead) Then
            Set Sht = Wbk.Worksheets(ShtNameToRead)
            DataRd = Sht.UsedRange
            Wbk.Close Savechanges:=False
        Else
            Set Sht = Wbk.Worksheets(1)
            DataRd = Sht.UsedRange
            Wbk.Close Savechanges:=False
            'ReDim DataRd(0 To 0, 0 To 0)
            'Msg1 = "No se encontró la hoja [" & ShtNameToRead & "]"
        End If
    End If
    If Msg = "" Then
        Msg = Msg1
    Else
        Msg = Msg & vbCrLf & Msg1
    End If
    ReadXlsSheet_UsedRangeToArray = DataRd
End Function

'--------------------------------- ReadXlsSheetsAll_ToMultipleArray
Public Function ReadXlsSheetsAll_ToMultipleArray(XlsFullName, Optional WbkPass As String = "", Optional Msg As String = "") As ReadXlsDatShts()
    Dim Wbk As Workbook, Sht As Worksheet, DataRd, ShtName As String, i As Integer, n As Integer, Ret() As ReadXlsDatShts
    Set Wbk = Application.Workbooks.Open(FileName:=XlsFullName, UpdateLinks:=False, Password:=WbkPass)
    Application.Windows(Wbk.Name).Visible = False
    ReDim Ret(1 To 50)
    i = 0
    For Each Sht In Wbk.Sheets
        i = i + 1
        ShtName = Sht.Name
        DataRd = Sht.UsedRange
        Ret(i).ShtName = ShtName
        Ret(i).Data = DataRd
        Msg = Msg & vbCrLf & "         >> Hoja [" & Ret(i).ShtName & "]"
    Next
    ReDim Preserve Ret(1 To i)
    Wbk.Close Savechanges:=False
    ReadXlsSheetsAll_ToMultipleArray = Ret
End Function

'--------------------------------- SheetExists
Public Function SheetExists(Wbk As Workbook, ShtName$) As Boolean
    Dim Ret As Boolean, Sht As Worksheet
    SheetExists = False
    For Each Sht In Wbk.Sheets
        If Sht.Name = ShtName Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

'--------------------------------- STR_AnchoFijo
Public Function STR_AnchoFijo(s1, Width As Long, Optional ReplaceStr = " ", Optional Brackets As Boolean = False, Optional NumEspacios& = 0)
    Dim s As String, s2$
    s = s1 & String(Width, ReplaceStr)
    s = Left(s, Width)
    If Brackets Then
        s = "[" & s & "]"
    End If
    s2 = String(NumEspacios, " ")
    STR_AnchoFijo = s2 & s
End Function

'--------------------------------- VBA_CrearModulo (Private)
Private Function VBA_CrearModulo(Wbk As Workbook, VbModuleName$, StrContenidoModulo$, Optional CodeType As vbext_ComponentType = vbext_ct_StdModule) As VBIDE.VBComponent   '
    Dim vbPrj As VBIDE.VBProject, VbMod As VBIDE.VBComponent, ModuleStr$
    Dim FndMod As Boolean, nLin&

    '----> Buscar en los modulos del Libro, si existe el modulo VbModuleName2 (lo deja vacio si lo encuentra)
    Set vbPrj = Wbk.VBProject

    FndMod = False
    For Each VbMod In vbPrj.VBComponents
        If VbMod.Name = VbModuleName Then
            nLin = VbMod.CodeModule.CountOfLines
            VbMod.CodeModule.DeleteLines 1, nLin
            FndMod = True
            Exit For
        End If
    Next

    '----> Crea el modulo en caso de que no lo encontró
    If Not FndMod Then
        Set VbMod = vbPrj.VBComponents.Add(CodeType)
        VbMod.Name = VbModuleName
    End If

    '----> Coloca el contenido del modulo fuente en el Módulo destino
    VbMod.CodeModule.AddFromString StrContenidoModulo

    Set vbPrj = Nothing
    Set VbMod = Nothing

End Function
'--------------------------------- STR_SinAcentosMinusc
Public Function STR_SinAcentosMinusc(s1$)
    Dim s$
    s = LCase(s1)

    s = Replace(s, "á", "a")
    s = Replace(s, "à", "a")
    s = Replace(s, "â", "a")
    s = Replace(s, "ä", "a")

    s = Replace(s, "é", "e")
    s = Replace(s, "è", "e")
    s = Replace(s, "ê", "e")
    s = Replace(s, "ë", "e")

    s = Replace(s, "í", "i")
    s = Replace(s, "ì", "i")
    s = Replace(s, "î", "i")
    s = Replace(s, "ï", "i")

    s = Replace(s, "ó", "o")
    s = Replace(s, "ò", "o")
    s = Replace(s, "ô", "o")
    s = Replace(s, "ö", "o")

    s = Replace(s, "ú", "u")
    s = Replace(s, "ù", "u")
    s = Replace(s, "û", "u")
    s = Replace(s, "ü", "u")

    s = Replace(s, "ñ", "n")
    s = Replace(s, "z", "s")

    STR_SinAcentosMinusc = s
End Function

'--------------------------------- Texto_BuscarFrases
Public Function Texto_BuscarFrases(Txt$, VectorDeTxt As Variant) As Boolean
    Dim BuscarFrase&, i&
    BuscarFrase = 0
    For i = LBound(VectorDeTxt) To UBound(VectorDeTxt)
        BuscarFrase = BuscarFrase + InStr(1, Txt, VectorDeTxt(i), vbTextCompare)
    Next i
    If BuscarFrase > 0 Then
        Texto_BuscarFrases = True
    Else
        Texto_BuscarFrases = False
    End If
End Function

'--------------------------------- Vector_FindElementPosition
Public Function Vector_FindElementPosition(Elem, ArryVector) As Long
    Dim n1%, n2%, i%, RetEmpty
    n1 = LBound(ArryVector)
    n2 = UBound(ArryVector)
    For i = n1 To n2
        If ArryVector(i) = Elem Then
            Vector_FindElementPosition = i
            Exit Function
        End If
    Next i
    Vector_FindElementPosition = -1
End Function

'--------------------------------- Vector_SetDataHeaderIdx
Public Function Vector_SetDataHeaderIdx(ByRef Array2D As Variant, Header2D As Variant, FieldName, Value As Variant, Optional Fecha As Boolean = False)
    Dim nColumn%
    nColumn = Array_HeaderColumnNumber(Header2D, FieldName)
    Array2D(1, nColumn) = Value
End Function

'--------------------------------- Xls_CopySheetsToNewWbk
Public Function Xls_CopySheetsToNewWbk(ShtsArray As Variant, FileName As String, FullPath As String, _
                                        Optional Extension As String = ".xlsx", _
                                        Optional DeleteBtns As Boolean = True, _
                                        Optional Fmt As XlFileFormat = xlWorkbookDefault, _
                                        Optional EliminarNombres As Boolean = False, _
                                        Optional ExcluyeNombres As String = "", _
                                        Optional CreateTables As Boolean = False) As String
                                        
    Dim NewWbk As Workbook, Sht As Worksheet, Btn As Variant, FullFile$, Wbk As Workbook, i%, Rg As Range
    Set Wbk = ThisWorkbook
    FullFile = FullPath & "\" & FileName & Extension
    Wbk.Sheets(ShtsArray).Copy
    Set NewWbk = ActiveWorkbook
    If DeleteBtns Then
        For Each Sht In NewWbk.Sheets
            Sht.Unprotect
            For Each Btn In Sht.Shapes
                Btn.Delete
            Next
        Next
    End If
    If EliminarNombres Then
        If ExcluyeNombres = "" Then
            For i = LBound(ShtsArray) To UBound(ShtsArray)
                 ExcluyeNombres = ExcluyeNombres & "|" & "DB" & ShtsArray(i)
            Next i
        End If
        Call Xls_LimpiarNombres(NewWbk, ExcluyeNombres, False)
    End If
    If CreateTables Then
        For Each Sht In NewWbk.Sheets
            Sht.AutoFilterMode = False
            Set Rg = Sht.Range("DB" & Sht.Name)
            Sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Rg, XlListObjectHasHeaders:=xlYes).Name = "T_" & Sht.Name
        Next
    End If
    Application.DisplayAlerts = False
    NewWbk.SaveAs FileName:=FullFile, FileFormat:=Fmt
    Application.DisplayAlerts = True
    NewWbk.Close
    Xls_CopySheetsToNewWbk = FullFile
End Function

'--------------------------------- Xls_LimpiarNombres
Public Function Xls_LimpiarNombres(Wbk As Workbook, LstName As String, Optional DeleteName As Boolean = False)
    Dim LstNm As Variant, n As Integer, Fnd As Boolean, LstDel As String, LstOk As String, NomOk As Boolean, RgName, Nom_i As String, Pfx As String, FndStd As Boolean
    Dim i As Long
    LstNm = Split(LstName, "|")
    n = UBound(LstNm)
    LstOk = ""
    LstDel = ""
    On Error Resume Next
    For Each RgName In Wbk.Names
        Nom_i = RgName.Name
        FndStd = Lib.Texto_BuscarFrases(Nom_i, Array("_FilterDatabase", "_xlfn.", "Print_Area", "Print_Titles"))
        If FndStd Then
            LstOk = LstOk & "|" & RgName.Name
        Else
            Fnd = False
            For i = 1 To n
                If Nom_i = LstNm(i) Then
                   Fnd = True
                   Exit For
                End If
            Next i
            
            If DeleteName Then
                If Fnd Then
                    LstDel = LstDel & "|" & RgName.Name
                    RgName.Delete
                Else
                    LstOk = LstOk & "|" & RgName.Name
                End If
            Else
                If Not Fnd Then
                    LstDel = LstDel & "|" & RgName.Name
                    RgName.Delete
                Else
                    LstOk = LstOk & "|" & RgName.Name
                End If
            End If
        End If
    Next
    On Error GoTo 0
    
End Function

'--------------------------------- Xls_OcultarColumna
Public Function Xls_OcultarColumna(Sht As Worksheet, nCol As Long)
    On Error Resume Next
    Sht.Unprotect Password:=""
    Sht.Columns(nCol).Ungroup
    On Error GoTo 0
    Sht.Columns(nCol).EntireColumn.Hidden = False
    Sht.Columns(nCol).Group
    Sht.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    Sht.Protect Password:="", UserInterfaceOnly:=True
End Function

'--------------------------------- Xls_OrdenarDB
Public Function Xls_OrdenarDB(ShtName$, RgName$, VectorColumnNames As Variant, VectorXlSortOrder As Variant, Optional Wbk As Workbook)
    Dim Sht As Worksheet, nF&, Sht1 As Worksheet, ScrUpd, Hdr As Variant, nFilas&, n%, n0%

    If Wbk Is Nothing Then
        Set Sht = ThisWorkbook.Worksheets(ShtName)
    Else
        Set Sht = Wbk.Worksheets(ShtName)
    End If


    Set Sht1 = Application.ActiveSheet
    ScrUpd = Application.ScreenUpdating
    nFilas = Sht.Range(RgName).Rows.Count
    nF = Sht.Range(RgName).Row
    Sht.Sort.SortFields.Clear
    Hdr = Sht.Range(RgName).Rows(1).Value2
    n0 = LBound(VectorColumnNames)
    n = UBound(VectorColumnNames)

    Dim i%, NumColumn&, KeyHdr As Range, RgKey As Range, Order As XlSortOrder
    Order = xlAscending
    For i = n0 To n
        NumColumn = Me.Array_HeaderColumnNumber(Hdr, VectorColumnNames(i))
        If NumColumn = -1 Then
            MsgBox "No se encontró la columna " & VectorColumnNames(i) & " para ordenar..."
        Else
            Set KeyHdr = Sht.Range(Sht.Cells(nF, 1), Sht.Cells(nF, 1)).Offset(0, NumColumn - 1)
            Set RgKey = Sht.Cells(nF, NumColumn).Resize(nFilas, 1)
            If Not IsEmpty(VectorXlSortOrder) Then
                Order = VectorXlSortOrder(i)
            End If

            Sht.Sort.SortFields.Add Key:=RgKey, SortOn:=xlSortOnValues, Order:=Order, DataOption:=xlSortNormal
        End If
    Next i

    '>>>>>> Actualizacion para ordenar cuando estas en otra hoja
    Application.ScreenUpdating = False
    Sht.Activate
    '>>>>>> ===========================================================
    Sht.Sort.SetRange Sht.Range(RgName)
    Sht.Sort.Header = xlYes
    Sht.Sort.MatchCase = False
    Sht.Sort.Orientation = xlTopToBottom
    Sht.Sort.SortMethod = xlPinYin
    Sht.Sort.Apply
    '>>>>>> Actualizacion para ordenar cuando estas en otra hoja
    Sht1.Activate
    Application.ScreenUpdating = ScrUpd
    '>>>>>> ===========================================================

End Function

'--------------------------------- Xls_ShowSheets
Public Function Xls_ShowSheets(ShtToShw)
    Dim Sht As Worksheet, FndSht%, Lst
    Application.ScreenUpdating = False
    For Each Sht In ThisWorkbook.Sheets
        FndSht = Me.Vector_FindElementPosition(Sht.Name, ShtToShw)
        If FndSht > -1 Then
            Sht.Visible = xlSheetVisible
        Else
            Sht.Visible = xlSheetHidden
        End If
    Next Sht
    Application.ScreenUpdating = True
End Function



