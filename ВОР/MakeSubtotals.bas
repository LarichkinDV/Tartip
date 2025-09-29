Option Explicit

' ===================== КОНФИГУРАЦИЯ =====================

Private Const REQ_TYPE As String = "Type Name : String"
Private Const REQ_DEM  As String = "Phase Demolished : String"
Private Const REQ_CR   As String = "Phase Created : String"

' Деление по площади включается ТОЛЬКО для этого типа
Private Const SPECIAL_TYPE As String = "(потолок)_жилье_натяжной.отм.3м_толщ=5мм"

' Нули считать «пустыми» при расчёте итогов по суммируемым колонкам (кроме New_Count)?
Private Const TREAT_ZERO_AS_EMPTY As Boolean = True

' Эпсилон для сравнения с целым и границами
Private Const eps As Double = 5E-07

' Какие Double-колонки суммировать
Private Function SUM_COLS() As Variant
    SUM_COLS = Array( _
        "New_Count : Double", _
        "Volume : Double", _
        "Area : Double", _
        "Length : Double", _
        "Perimeter : Double", _
        "Unconnected Height : Double" _
    )
End Function

' Какие колонки показывать (остальные скрыть, кроме тех, что начинаются с new_)
Private Function KEEP_COLS() As Variant
    KEEP_COLS = Array( _
        "ID", _
        "Type Name : String", _
        "Category : String", _
        "New_Count : Double", _
        "Volume : Double", _
        "Area : Double", _
        "Length : Double", _
        "Width : Double", _
        "Phase Demolished : String", _
        "Phase Created : String", _
        "Thickness : Double", _
        "Perimeter : Double", _
        "Unconnected Height : Double", _
        "Height : Double" _
    )
End Function

' ===================== ОСНОВНОЙ МАКРОС =====================

Public Sub MakeSubtotals_Configurable()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 1) Границы и заголовки
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo SafeExit

    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column

    Dim hdr() As String
    ReDim hdr(1 To lastCol)
    For j = 1 To lastCol: hdr(j) = CStr(ws.Cells(1, j).Value): Next j

    ' 2) Поиск обязательных + Area
    Dim colType As Long, colDem As Long, colCr As Long, colArea As Long
    colType = FindHeader(hdr, REQ_TYPE)
    colDem = FindHeader(hdr, REQ_DEM)
    colCr = FindHeader(hdr, REQ_CR)
    colArea = FindHeader(hdr, "Area : Double")

    If colType = 0 Or colDem = 0 Or colCr = 0 Then
        MsgBox "Не найдены обязательные колонки: " & REQ_TYPE & ", " & REQ_DEM & ", " & REQ_CR, vbCritical
        GoTo SafeExit
    End If

    ' 3) Обеспечить New_Count : Double (вставить перед Volume : Double, либо в конец)
    Dim colNewCount As Long
    colNewCount = FindHeader(hdr, "New_Count : Double")
    If colNewCount = 0 Then
        Dim colVol As Long: colVol = FindHeader(hdr, "Volume : Double")
        If colVol = 0 Then
            ws.Cells(1, lastCol + 1).EntireColumn.Insert
            ws.Cells(1, lastCol + 1).Value = "New_Count : Double"
            colNewCount = lastCol + 1
        Else
            ws.Columns(colVol).EntireColumn.Insert
            ws.Cells(1, colVol).Value = "New_Count : Double"
            colNewCount = colVol
        End If

        lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
        ReDim hdr(1 To lastCol)
        For j = 1 To lastCol: hdr(j) = CStr(ws.Cells(1, j).Value): Next j
        colType = FindHeader(hdr, REQ_TYPE)
        colDem = FindHeader(hdr, REQ_DEM)
        colCr = FindHeader(hdr, REQ_CR)
        colArea = FindHeader(hdr, "Area : Double")
    End If

    ' 4) Инициализировать New_Count = 1 (не для строк «Итого»)
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Not IsItogoRow(ws.Cells(i, colType).Value) Then
            ws.Cells(i, colNewCount).Value = 1
        End If
    Next i

    ' 5) Предклассификация столбцов
    Dim isSum() As Boolean, isProtectedNew() As Boolean
    ReDim isSum(1 To lastCol)
    ReDim isProtectedNew(1 To lastCol)
    For j = 1 To lastCol
        isSum(j) = IsSumHeader(hdr(j))
        isProtectedNew(j) = ShouldSkipWrite(hdr(j))
    Next j

    ' 5.1) Видимость столбцов
    ApplyColumnVisibility ws, hdr, lastCol

    ' 6) Фильтрация данных (удалить неподходящее; пустые Type — удалить)
    For i = lastRow To 2 Step -1
        Dim tname As String: tname = Trim$(CStr(ws.Cells(i, colType).Value))
        If IsItogoRow(tname) Then GoTo NextRowDelete
        If Len(tname) = 0 Then
            ws.Rows(i).Delete
        Else
            Dim code As Long: code = AllowedStageCode(ws.Cells(i, colDem).Value, ws.Cells(i, colCr).Value)
            If code = 0 Then ws.Rows(i).Delete
        End If
NextRowDelete:
    Next i

    ' Подсчёт отобранных
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim selectedCount As Long: selectedCount = 0
    If lastRow >= 2 Then
        For i = 2 To lastRow
            If Not IsItogoRow(CStr(ws.Cells(i, colType).Value)) Then selectedCount = selectedCount + 1
        Next i
    End If
    If lastRow < 2 Then GoTo FinishWithReport

    ' 7) Сортировка: Stage ? Name ? Bucket ? Flag(Itogo=0/Data=1)
    PrepareForSort ws

    Dim sortStageCol As Long, sortNameCol As Long, sortBucketCol As Long, sortFlagCol As Long
    sortStageCol = lastCol + 1: ws.Cells(1, sortStageCol).Value = "__SortStage__"
    sortNameCol = lastCol + 2: ws.Cells(1, sortNameCol).Value = "__SortName__"
    sortBucketCol = lastCol + 3: ws.Cells(1, sortBucketCol).Value = "__SortBucket__"
    sortFlagCol = lastCol + 4: ws.Cells(1, sortFlagCol).Value = "__SortFlag__"

    Dim dem As String, cr As String, b As Long
    For i = 2 To lastRow
        Dim cellType As String: cellType = CStr(ws.Cells(i, colType).Value)
        If IsItogoRow(cellType) Then
            Dim nm As String, stg As Long
            ParseItogo cellType, nm, stg
            ws.Cells(i, sortStageCol).Value = IIf(stg > 0, stg, 99)
            ws.Cells(i, sortNameCol).Value = LCase$(Trim$(nm))
            ws.Cells(i, sortBucketCol).Value = ParseBucketFromCaption(cellType)
            ws.Cells(i, sortFlagCol).Value = 0
        Else
            dem = CleanText(ws.Cells(i, colDem).Value)
            cr = CleanText(ws.Cells(i, colCr).Value)
            ws.Cells(i, sortStageCol).Value = IIf(AllowedStageCode(dem, cr) > 0, AllowedStageCode(dem, cr), 99)
            ws.Cells(i, sortNameCol).Value = LCase$(Trim$(cellType))
            b = 0
            If colArea > 0 And IsSpecialType(cellType) And IsNumeric(ws.Cells(i, colArea).Value) Then
                b = AreaBucket(ws.Cells(i, colArea).Value)
            End If
            ws.Cells(i, sortBucketCol).Value = b
            ws.Cells(i, sortFlagCol).Value = 1
        End If
    Next i

    Dim rngAll As Range
    Set rngAll = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, sortFlagCol))
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Columns(sortStageCol), Order:=xlAscending
        .SortFields.Add key:=ws.Columns(sortNameCol), Order:=xlAscending
        .SortFields.Add key:=ws.Columns(sortBucketCol), Order:=xlAscending
        .SetRange rngAll
        .header = xlYes
        .Apply
    End With

    ws.Columns(sortFlagCol).Delete
    ws.Columns(sortBucketCol).Delete
    ws.Columns(sortNameCol).Delete
    ws.Columns(sortStageCol).Delete

    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 8) Агрегация по группам (Type + Stage + Bucket)
    Dim cap As Long: cap = 64
    Dim grpCount As Long: grpCount = 0

    Dim keys() As String, nameByGrp() As String, stageByGrp() As Long, bucketByGrp() As Long
    Dim sums() As Variant, hasNumArr() As Variant, texts() As Variant, nonZeroArr() As Variant

    ReDim keys(1 To cap)
    ReDim nameByGrp(1 To cap)
    ReDim stageByGrp(1 To cap)
    ReDim bucketByGrp(1 To cap)
    ReDim sums(1 To cap)
    ReDim hasNumArr(1 To cap)
    ReDim texts(1 To cap)
    ReDim nonZeroArr(1 To cap)

    For i = 2 To lastRow
        Dim nmRow As String: nmRow = CStr(ws.Cells(i, colType).Value)
        If IsItogoRow(nmRow) Then GoTo NextRowAgg
        dem = ws.Cells(i, colDem).Value
        cr = ws.Cells(i, colCr).Value
        Dim stgCode As Long: stgCode = AllowedStageCode(dem, cr)
        If stgCode = 0 Then GoTo NextRowAgg

        b = 0
        If colArea > 0 And IsSpecialType(nmRow) Then
            b = AreaBucket(ws.Cells(i, colArea).Value) ' 0/1/2/3
        End If

        Dim key As String
        key = LCase$(Trim$(nmRow)) & "||" & CStr(stgCode) & "||B" & CStr(b)

        Dim gIdx As Long: gIdx = FindGroupIndex(keys, grpCount, key)
        If gIdx = 0 Then
            If grpCount = cap Then
                cap = cap * 2
                ReDim Preserve keys(1 To cap)
                ReDim Preserve nameByGrp(1 To cap)
                ReDim Preserve stageByGrp(1 To cap)
                ReDim Preserve bucketByGrp(1 To cap)
                ReDim Preserve sums(1 To cap)
                ReDim Preserve hasNumArr(1 To cap)
                ReDim Preserve texts(1 To cap)
                ReDim Preserve nonZeroArr(1 To cap)
            End If
            grpCount = grpCount + 1
            gIdx = grpCount
            keys(gIdx) = key
            nameByGrp(gIdx) = nmRow
            stageByGrp(gIdx) = stgCode
            bucketByGrp(gIdx) = b
            sums(gIdx) = NewDoubles(lastCol)
            hasNumArr(gIdx) = NewBools(lastCol)
            texts(gIdx) = NewStrings(lastCol)
            nonZeroArr(gIdx) = NewBools(lastCol)
        End If

        Dim aSum As Variant, aHas As Variant, aTxt As Variant, aNZ As Variant, v As Variant
        aSum = sums(gIdx): aHas = hasNumArr(gIdx): aTxt = texts(gIdx): aNZ = nonZeroArr(gIdx)

        For j = 1 To lastCol
            If isProtectedNew(j) Then
                ' пропуск new_* (кроме New_Count, она не попадает сюда)
            ElseIf isSum(j) Then
                v = ws.Cells(i, j).Value
                If IsNumeric(v) Then
                    aSum(j) = aSum(j) + CDbl(v)
                    aHas(j) = True
                    If (Not TREAT_ZERO_AS_EMPTY) Then
                        aNZ(j) = True
                    ElseIf Abs(CDbl(v)) > 5E-12 Then
                        aNZ(j) = True
                    End If
                End If
            Else
                v = ws.Cells(i, j).Value
                If Len(CStr(v)) > 0 Then
                    aTxt(j) = AddUniqueText(aTxt(j), NumToTextTokenRounded(v, 7), ";")
                End If
            End If
        Next j

        sums(gIdx) = aSum: hasNumArr(gIdx) = aHas: texts(gIdx) = aTxt: nonZeroArr(gIdx) = aNZ
NextRowAgg:
    Next i

    ' 9) Вставка/обновление «Итого»
    Dim insRow As Long, k As Long
    For k = 1 To grpCount
        insRow = FindItogoRowByKey(ws, colType, nameByGrp(k), stageByGrp(k), bucketByGrp(k))
        If insRow = 0 Then
            Dim firstData As Long
            firstData = FindFirstRowOfGroup(ws, colType, colDem, colCr, colArea, _
                                            nameByGrp(k), stageByGrp(k), bucketByGrp(k))
            If firstData = 0 Then firstData = ws.Cells(ws.Rows.Count, colType).End(xlUp).Row + 1
            ws.Rows(firstData).Insert Shift:=xlDown
            insRow = firstData
        End If

        Dim aSumW As Variant, aHasW As Variant, aTxtW As Variant, aNZw As Variant
        aSumW = sums(k): aHasW = hasNumArr(k): aTxtW = texts(k): aNZw = nonZeroArr(k)

        For j = 1 To lastCol
            If isProtectedNew(j) Then
                ' не изменяем new_*
            ElseIf isSum(j) Then
                If aHasW(j) Then
                    Dim valToWrite As Double
                    If IsNewCountHeader(hdr(j)) Then
                        valToWrite = RoundN(aSumW(j), 7)
                    Else
                        If aNZw(j) Then
                            Dim origVal As Double, r7 As Double
                            origVal = aSumW(j): r7 = RoundN(origVal, 7)
                            If ShouldRound7(origVal, r7) Then
                                valToWrite = r7
                            Else
                                valToWrite = origVal
                            End If
                        Else
                            ws.Cells(insRow, j).ClearContents
                            GoTo ContinueCols
                        End If
                    End If
                    valToWrite = SnapIfNearInteger(valToWrite, eps)
                    With ws.Cells(insRow, j)
                        .Value = valToWrite
                        .NumberFormat = SumNumberFormatForValue(hdr(j), valToWrite)
                    End With
                Else
                    ws.Cells(insRow, j).ClearContents
                End If
            Else
                ' --- НЕ СУММИРУЕМЫЕ КОЛОНКИ ---
                Dim tok As String
                tok = aTxtW(j)  ' уже уникализировано и округлено как текст-токены

                With ws.Cells(insRow, j)
                    If Len(tok) = 0 Then
                        .ClearContents
                    ElseIf IsDoubleHeader(hdr(j)) Then
                        ' Для * : Double — если итог одно число, пишем как число; иначе как текст
                        If InStr(1, tok, ";", vbBinaryCompare) = 0 And IsNumeric(tok) Then
                            Dim num As Double
                            num = CDbl(tok)
                            num = SnapIfNearInteger(RoundN(num, 7), eps)
                            .Value = num
                            .NumberFormat = SumNumberFormatForValue(hdr(j), num)
                        Else
                            .NumberFormat = "@"
                            .Value = tok
                        End If
                    Else
                        ' Прочие нечисловые поля — всегда текст
                        .NumberFormat = "@"
                        .Value = tok
                    End If
                End With
            End If
ContinueCols:
        Next j

        ws.Cells(insRow, colType).Value = BuildItogoCaption(nameByGrp(k), stageByGrp(k), bucketByGrp(k))

        With ws.Rows(insRow)
            .Font.Bold = True
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
        End With
    Next k

    ' 10) Группировка (Outline)
    ws.Cells.ClearOutline
    ws.Outline.SummaryRow = xlAbove
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim itogoR As Long, rStart As Long, rEnd As Long, r As Long
    For k = 1 To grpCount
        itogoR = FindItogoRowByKey(ws, colType, nameByGrp(k), stageByGrp(k), bucketByGrp(k))
        If itogoR > 0 Then
            rStart = itogoR + 1
            If rStart <= lastRow Then
                r = rStart
                Do While r <= lastRow
                    If IsItogoRow(CStr(ws.Cells(r, colType).Value)) Then Exit Do
                    If LCase$(Trim$(CStr(ws.Cells(r, colType).Value))) <> LCase$(Trim$(nameByGrp(k))) Then Exit Do
                    If AllowedStageCode(ws.Cells(r, colDem).Value, ws.Cells(r, colCr).Value) <> stageByGrp(k) Then Exit Do
                    If bucketByGrp(k) > 0 And colArea > 0 And IsSpecialType(nameByGrp(k)) Then
                        If AreaBucket(ws.Cells(r, colArea).Value) <> bucketByGrp(k) Then Exit Do
                    End If
                    r = r + 1
                Loop
                rEnd = r - 1
                If rEnd >= rStart Then
                    On Error Resume Next
                    ws.Rows(rStart & ":" & rEnd).Group
                    On Error GoTo 0
                End If
            End If
        End If
    Next k

    ' 11) Автоподбор ширины New_Count
    colNewCount = FindHeaderRow(ws, "New_Count : Double")
    If colNewCount > 0 Then ws.Columns(colNewCount).AutoFit

    ' 12) Повторно применить видимость колонок
    RefreshHeaders ws, hdr, lastCol
    ApplyColumnVisibility ws, hdr, lastCol

FinishWithReport:
    MsgBox "Готово." & vbCrLf & _
           "Отобрано элементов: " & CStr(selectedCount) & vbCrLf & _
           "Сформировано/обновлено строк «Итого»: " & CStr(grpCount), _
           vbInformation, "Итоги обработки"

SafeExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ===================== ВСПОМОГАТЕЛЬНЫЕ =====================

' Подготовка листа к сортировке (снять фильтры/оутлайн)
Private Sub PrepareForSort(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.ShowAutoFilter = True
        lo.AutoFilter.ShowAllData
    Next lo
    ws.Cells.ClearOutline
    On Error GoTo 0
End Sub

' ----- округление / формат -----
Private Function RoundN(ByVal v As Variant, ByVal n As Long) As Double
    Dim x As Double, r As Double
    If Not IsNumeric(v) Then RoundN = 0: Exit Function
    x = CDbl(v)
    On Error Resume Next
    r = Application.WorksheetFunction.Round(x, n)
    If Err.Number <> 0 Then Err.Clear: r = VBA.Round(x, n)
    On Error GoTo 0
    If Abs(r) < 5E-12 Then r = 0
    RoundN = r
End Function

Private Function SnapIfNearInteger(ByVal v As Double, Optional ByVal eps As Double = eps) As Double
    Dim r0 As Double: r0 = Application.WorksheetFunction.Round(v, 0)
    If Abs(v - r0) <= eps Then SnapIfNearInteger = r0 Else SnapIfNearInteger = v
End Function

Private Function ShouldRound7(ByVal original As Double, ByVal r7 As Double) As Boolean
    ShouldRound7 = (Abs(original - r7) > 5E-12)
End Function

Private Function IsNewCountHeader(ByVal header As String) As Boolean
    IsNewCountHeader = (LCase$(Trim$(NormHeader(header))) = LCase$(Trim$(NormHeader("New_Count : Double"))))
End Function

' целое ли значение с учётом eps
Private Function IsIntegerish(ByVal v As Double, Optional ByVal eps As Double = eps) As Boolean
    IsIntegerish = (Abs(v - Application.WorksheetFunction.Round(v, 0)) <= eps)
End Function

' формат числа по значению: целое ? "0"; иначе ? "0.#######"
Private Function SumNumberFormatForValue(ByVal header As String, ByVal v As Double) As String
    If IsNewCountHeader(header) Or IsIntegerish(v) Then
        SumNumberFormatForValue = "0"
    Else
        SumNumberFormatForValue = "0.#######"
    End If
End Function

' Несуммируемые Double ? текст-токены с округлением и «прилипанием»
Private Function NumToTextTokenRounded(ByVal val As Variant, ByVal n As Long) As String
    Dim decSep As String, s As String, r As Double
    If Not IsNumeric(val) Then
        NumToTextTokenRounded = Trim$(CStr(val)): Exit Function
    End If
    decSep = Application.International(xlDecimalSeparator)
    r = RoundN(val, n)
    r = SnapIfNearInteger(r, eps)
    s = CStr(r)
    s = Replace$(Replace$(s, ".", decSep), ",", decSep)
    NumToTextTokenRounded = TrimTrailingZerosText(s, decSep)
End Function

Private Function TrimTrailingZerosText(ByVal s As String, ByVal decSep As String) As String
    Dim p As Long: p = InStr(1, s, decSep, vbBinaryCompare)
    If p = 0 Then TrimTrailingZerosText = s: Exit Function
    Do While Len(s) > p And Right$(s, 1) = "0": s = Left$(s, Len(s) - 1): Loop
    If Right$(s, 1) = decSep Then s = Left$(s, Len(s) - 1)
    TrimTrailingZerosText = s
End Function

' Является ли заголовок колонкой вида "* : Double"
Private Function IsDoubleHeader(ByVal header As String) As Boolean
    Dim h As String
    h = LCase$(Trim$(NormHeader(header)))
    IsDoubleHeader = (Right$(h, Len(" : double")) = " : double")
End Function

' ----- заголовки/поиск/утилиты -----
Private Function NormHeader(ByVal s As String) As String
    s = Replace(CStr(s), Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(s, " : ", ":"): s = Replace(s, " :", ":"): s = Replace(s, ": ", ":")
    s = Replace(s, ":", " : ")
    NormHeader = s
End Function

Private Function NameInList(ByVal header As String, ByVal keepCols As Variant) As Boolean
    Dim k As Long, h As String, cmp As String
    h = LCase$(NormHeader(header))
    For k = LBound(keepCols) To UBound(keepCols)
        cmp = LCase$(NormHeader(CStr(keepCols(k))))
        If h = cmp Then NameInList = True: Exit Function
    Next k
    NameInList = False
End Function

Private Sub ApplyColumnVisibility(ByVal ws As Worksheet, ByRef headers() As String, ByVal lastCol As Long)
    Dim keep As Variant: keep = KEEP_COLS()
    Dim j As Long, hn As String
    For j = 1 To lastCol
        hn = NormHeader(headers(j))
        If Left$(LCase$(hn), 4) = "new_" Then
            ws.Columns(j).Hidden = False
        ElseIf NameInList(hn, keep) Then
            ws.Columns(j).Hidden = False
        Else
            ws.Columns(j).Hidden = True
        End If
    Next j
End Sub

Private Sub RefreshHeaders(ByVal ws As Worksheet, ByRef headers() As String, ByRef lastCol As Long)
    Dim j As Long
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    ReDim headers(1 To lastCol)
    For j = 1 To lastCol: headers(j) = CStr(ws.Cells(1, j).Value): Next j
End Sub

' Разрешённые пары стадий
Private Function AllowedStageCode(ByVal demol As Variant, ByVal created As Variant) As Long
    Dim d As String: d = CleanText(demol)
    Dim c As String: c = CleanText(created)
    If d = "демонтаж" And c = "существующие" Then
        AllowedStageCode = 1
    ElseIf d = "none" And c = "новая конструкция" Then
        AllowedStageCode = 2
    Else
        AllowedStageCode = 0
    End If
End Function

' Не перезаписывать ли столбец? True для new_* (кроме системного new_count : double)
Private Function ShouldSkipWrite(ByVal header As String) As Boolean
    Dim h As String: h = LCase$(Trim$(NormHeader(header)))
    ShouldSkipWrite = (Left$(h, 4) = "new_" And h <> "new_count : double")
End Function

' Суммируемая ли колонка?
Private Function IsSumHeader(ByVal header As String) As Boolean
    Dim target As String: target = LCase$(Trim$(NormHeader(header)))
    Dim a As Variant, i As Long: a = SUM_COLS()
    For i = LBound(a) To UBound(a)
        If target = LCase$(Trim$(NormHeader(CStr(a(i))))) Then IsSumHeader = True: Exit Function
    Next i
End Function

Private Function IsItogoRow(ByVal typeCellValue As String) As Boolean
    IsItogoRow = (Left$(Trim$(typeCellValue), 6) = "Итого:")
End Function

Private Sub ParseItogo(ByVal itogoText As String, ByRef nameOut As String, ByRef stageOut As Long)
    Dim s As String, p As Long, tag As String
    s = Trim$(CStr(itogoText))
    If Left$(s, 6) = "Итого:" Then s = Trim$(Mid$(s, 7))
    p = InStr(1, s, "[", vbTextCompare)
    If p > 0 Then
        nameOut = Trim$(Left$(s, p - 1))
        tag = LCase$(Mid$(s, p + 1))
        tag = Replace$(tag, "]", "")
        If InStr(tag, "демонтаж") > 0 Then
            stageOut = 1
        ElseIf InStr(tag, "новая конструкция") > 0 Then
            stageOut = 2
        Else
            stageOut = 0
        End If
    Else
        nameOut = s
        stageOut = 0
    End If
End Sub

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    CleanText = LCase$(s)
End Function

Private Function FindHeader(ByRef headers() As String, ByVal name As String) As Long
    Dim j As Long, target As String: target = LCase$(Trim$(NormHeader(name)))
    For j = LBound(headers) To UBound(headers)
        If LCase$(Trim$(NormHeader(headers(j)))) = target Then FindHeader = j: Exit Function
    Next j
End Function

Private Function FindHeaderRow(ws As Worksheet, ByVal name As String) As Long
    Dim lastC As Long, jj As Long
    lastC = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    For jj = 1 To lastC
        If LCase$(Trim$(NormHeader(ws.Cells(1, jj).Value))) = LCase$(Trim$(NormHeader(name))) Then
            FindHeaderRow = jj: Exit Function
        End If
    Next jj
    FindHeaderRow = 0
End Function

Private Function AddUniqueText(ByVal existing As String, ByVal newVal As String, ByVal sep As String) As String
    newVal = Trim$(CStr(newVal))
    If Len(newVal) = 0 Then
        AddUniqueText = existing
    ElseIf Len(existing) = 0 Then
        AddUniqueText = newVal
    ElseIf InStr(1, sep & existing & sep, sep & newVal & sep, vbTextCompare) = 0 Then
        AddUniqueText = existing & sep & newVal
    Else
        AddUniqueText = existing
    End If
End Function

Private Function FindGroupIndex(ByRef keys() As String, ByVal cnt As Long, ByVal key As String) As Long
    Dim ii As Long
    For ii = 1 To cnt
        If keys(ii) = key Then FindGroupIndex = ii: Exit Function
    Next ii
    FindGroupIndex = 0
End Function

Private Function NewDoubles(n As Long) As Variant
    Dim a() As Double: ReDim a(1 To n): NewDoubles = a
End Function
Private Function NewBools(n As Long) As Variant
    Dim a() As Boolean: ReDim a(1 To n): NewBools = a
End Function
Private Function NewStrings(n As Long) As Variant
    Dim a() As String: ReDim a(1 To n): NewStrings = a
End Function

' ----- Деление по площади для SPECIAL_TYPE -----

Private Function IsSpecialType(ByVal typeName As String) As Boolean
    IsSpecialType = (LCase$(Trim$(typeName)) = LCase$(Trim$(SPECIAL_TYPE)))
End Function

' 0 — нет корзины; 1 — <10; 2 — >10 и <50; 3 — >50 (строгие границы)
Private Function AreaBucket(ByVal v As Variant) As Long
    If Not IsNumeric(v) Then Exit Function
    Dim x As Double: x = CDbl(v)
    If x < 10# - eps Then
        AreaBucket = 1
    ElseIf x > 10# + eps And x < 50# - eps Then
        AreaBucket = 2
    ElseIf x > 50# + eps Then
        AreaBucket = 3
    Else
        AreaBucket = 0
    End If
End Function

Private Function BucketLabel(ByVal bucket As Long, ByVal typeName As String) As String
    If Not IsSpecialType(typeName) Then Exit Function
    Select Case bucket
        Case 1: BucketLabel = " (до 10м2)"
        Case 2: BucketLabel = " (от 10 до 50м2)"
        Case 3: BucketLabel = " (от 50м2)"
        Case Else: BucketLabel = ""
    End Select
End Function

Private Function BuildItogoCaption(ByVal grpName As String, ByVal stageCode As Long, ByVal bucket As Long) As String
    Dim lbl As String
    If stageCode = 1 Then
        lbl = " [Демонтаж]"
    ElseIf stageCode = 2 Then
        lbl = " [Новая конструкция]"
    Else
        lbl = ""
    End If
    BuildItogoCaption = "Итого: " & grpName & lbl & BucketLabel(bucket, grpName)
End Function

Private Function ParseBucketFromCaption(ByVal cap As String) As Long
    Dim s As String: s = LCase$(cap)
    If InStr(s, "(до 10м2)") > 0 Then
        ParseBucketFromCaption = 1
    ElseIf InStr(s, "(от 10 до 50м2)") > 0 Then
        ParseBucketFromCaption = 2
    ElseIf InStr(s, "(от 50м2)") > 0 Then
        ParseBucketFromCaption = 3
    Else
        ParseBucketFromCaption = 0
    End If
End Function

Private Function FindItogoRowByKey(ws As Worksheet, typeCol As Long, _
                                   grpName As String, stageCode As Long, bucket As Long) As Long
    Dim want As String: want = BuildItogoCaption(grpName, stageCode, bucket)
    Dim r As Long, lastR As Long: lastR = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, typeCol).Value) = want Then
            FindItogoRowByKey = r: Exit Function
        End If
    Next r
    FindItogoRowByKey = 0
End Function

Private Function FindFirstRowOfGroup(ws As Worksheet, typeCol As Long, demCol As Long, crCol As Long, areaCol As Long, _
                                     grpName As String, stageCode As Long, bucket As Long) As Long
    Dim r As Long, lastR As Long: lastR = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row
    For r = 2 To lastR
        If Not IsItogoRow(CStr(ws.Cells(r, typeCol).Value)) Then
            If LCase$(Trim$(ws.Cells(r, typeCol).Value)) = LCase$(Trim$(grpName)) Then
                If AllowedStageCode(ws.Cells(r, demCol).Value, ws.Cells(r, crCol).Value) = stageCode Then
                    If bucket > 0 And areaCol > 0 And IsSpecialType(grpName) Then
                        If AreaBucket(ws.Cells(r, areaCol).Value) <> bucket Then GoTo NextR
                    End If
                    FindFirstRowOfGroup = r: Exit Function
                End If
            End If
        End If
NextR:
    Next r
    FindFirstRowOfGroup = 0
End Function


