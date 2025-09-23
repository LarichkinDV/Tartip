Attribute VB_Name = "MakeSubtotals"

Option Explicit

Sub MakeSubtotals_FastStable()
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
    For j = 1 To lastCol
        hdr(j) = CStr(ws.Cells(1, j).Value)
    Next j

    ' 2) Ключевые колонки
    Dim colType As Long, colDem As Long, colCr As Long, colArea As Long, colID As Long
    colType = FindHeader(hdr, "Type Name : String")
    colDem = FindHeader(hdr, "Phase Demolished : String")
    colCr = FindHeader(hdr, "Phase Created : String")
    colArea = FindHeader(hdr, "Area : Double")
    colID = FindHeader(hdr, "ID")
    If colType = 0 Or colDem = 0 Or colCr = 0 Or colArea = 0 Or colID = 0 Then
        MsgBox "Не найдены ключевые колонки (Type Name / Phase Demolished / Phase Created / Area / ID).", vbCritical
        GoTo SafeExit
    End If

    ' 3) Предклассификация столбцов
    Dim isNum() As Boolean, isNew() As Boolean
    ReDim isNum(1 To lastCol)
    ReDim isNew(1 To lastCol)
    For j = 1 To lastCol
        isNum(j) = IsNumericHeader(hdr(j))
        isNew(j) = (Left$(LCase$(Trim$(hdr(j))), 4) = "new_")
    Next j

    ' 4) Сортировка
    Dim sortCol As Long: sortCol = lastCol + 1
    ws.Cells(1, sortCol).Value = "__SortKey__"
    Dim dem As String, cr As String
    For i = 2 To lastRow
        If Len(CleanText(ws.Cells(i, colType).Value)) = 0 Then
            ws.Cells(i, sortCol).Value = 99
        Else
            dem = CleanText(ws.Cells(i, colDem).Value)
            cr = CleanText(ws.Cells(i, colCr).Value)
            If dem = "демонтаж" And cr = "существующие" Then
                ws.Cells(i, sortCol).Value = 1
            ElseIf dem = "none" And cr = "новая конструкция" Then
                ws.Cells(i, sortCol).Value = 2
            Else
                ws.Cells(i, sortCol).Value = 99
            End If
        End If
    Next i

    Dim rngAll As Range
    Set rngAll = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, sortCol))
    rngAll.Sort Key1:=ws.Cells(1, sortCol), Order1:=xlAscending, _
                Key2:=ws.Cells(1, colType), Order2:=xlAscending, Header:=xlYes
    ws.Columns(sortCol).Delete
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column

    ' 5) Агрегация
    Dim cap As Long: cap = 64
    Dim grpCount As Long: grpCount = 0

    Dim keys() As String, nameByGrp() As String, demByGrp() As String, crByGrp() As String
    Dim sums() As Variant, hasNumArr() As Variant, texts() As Variant
    Dim idSets() As Collection

    ReDim keys(1 To cap)
    ReDim nameByGrp(1 To cap)
    ReDim demByGrp(1 To cap)
    ReDim crByGrp(1 To cap)
    ReDim sums(1 To cap)
    ReDim hasNumArr(1 To cap)
    ReDim texts(1 To cap)
    ReDim idSets(1 To cap)

    For i = 2 To lastRow
        Dim tname As String: tname = Trim$(CStr(ws.Cells(i, colType).Value))
        If Len(tname) = 0 Then GoTo NextRow

        dem = CleanText(ws.Cells(i, colDem).Value)
        cr = CleanText(ws.Cells(i, colCr).Value)
        If Not ((dem = "демонтаж" And cr = "существующие") _
             Or (dem = "none" And cr = "новая конструкция")) Then GoTo NextRow

        Dim key As String
        key = CleanText(tname) & "||" & dem & "||" & cr

        Dim gIdx As Long: gIdx = FindGroupIndex(keys, grpCount, key)
        If gIdx = 0 Then
            If grpCount = cap Then
                cap = cap * 2
                ReDim Preserve keys(1 To cap)
                ReDim Preserve nameByGrp(1 To cap)
                ReDim Preserve demByGrp(1 To cap)
                ReDim Preserve crByGrp(1 To cap)
                ReDim Preserve sums(1 To cap)
                ReDim Preserve hasNumArr(1 To cap)
                ReDim Preserve texts(1 To cap)
                ReDim Preserve idSets(1 To cap)
            End If
            grpCount = grpCount + 1
            gIdx = grpCount
            keys(gIdx) = key
            nameByGrp(gIdx) = tname
            demByGrp(gIdx) = dem
            crByGrp(gIdx) = cr
            sums(gIdx) = NewDoubles(lastCol)
            hasNumArr(gIdx) = NewBools(lastCol)
            texts(gIdx) = NewStrings(lastCol)
            Set idSets(gIdx) = New Collection
        End If

        ' агрегирование
        Dim aSum As Variant, aHas As Variant, aTxt As Variant, v As Variant
        aSum = sums(gIdx)
        aHas = hasNumArr(gIdx)
        aTxt = texts(gIdx)

        For j = 1 To lastCol
            If isNew(j) Then
                ' пропускаем
            ElseIf isNum(j) Then
                v = ws.Cells(i, j).Value
                If IsNumeric(v) Then
                    aSum(j) = aSum(j) + CDbl(v)
                    aHas(j) = True
                End If
            Else
                v = ws.Cells(i, j).Value
                If Len(CStr(v)) > 0 Then
                    aTxt(j) = AddUniqueText(aTxt(j), CStr(v), ";")
                End If
            End If
        Next j

        sums(gIdx) = aSum
        hasNumArr(gIdx) = aHas
        texts(gIdx) = aTxt

        ' уникальные ID
        Dim curID As String: curID = Trim$(CStr(ws.Cells(i, colID).Value))
        If Len(curID) > 0 Then
            On Error Resume Next
            idSets(gIdx).Add curID, curID
            On Error GoTo 0
        End If

NextRow:
    Next i

    ' 6) Итоги
    If grpCount = 0 Then
        MsgBox "Подходящих групп не найдено.", vbInformation
        GoTo SafeExit
    End If

    Dim insRow As Long, lbl As String, k As Long
    For k = 1 To grpCount
        insRow = FindItogoRow(ws, colType, nameByGrp(k))
        If insRow = 0 Then
            insRow = FindFirstRowOfGroup(ws, colType, colDem, colCr, nameByGrp(k), demByGrp(k), crByGrp(k))
            If insRow = 0 Then insRow = ws.Cells(ws.Rows.Count, colType).End(xlUp).Row + 1
            ws.Rows(insRow).Insert Shift:=xlDown
        End If

        Dim aSumW As Variant, aHasW As Variant, aTxtW As Variant
        aSumW = sums(k): aHasW = hasNumArr(k): aTxtW = texts(k)

        For j = 1 To lastCol
            If isNew(j) Then
                ' new_* не изменяем
            ElseIf isNum(j) Then
                If aHasW(j) Then
                    ws.Cells(insRow, j).Value = aSumW(j)
                Else
                    ws.Cells(insRow, j).ClearContents
                End If
            Else
                ws.Cells(insRow, j).Value = aTxtW(j)
            End If
        Next j

        If demByGrp(k) = "демонтаж" And crByGrp(k) = "существующие" Then
            lbl = " [Демонтаж]"
        ElseIf demByGrp(k) = "none" And crByGrp(k) = "новая конструкция" Then
            lbl = " [Новая конструкция]"
        Else
            lbl = ""
        End If

        ws.Cells(insRow, colType).Value = "Итого: " & nameByGrp(k) & lbl & _
            ", уникальных элементов: " & idSets(k).Count

        With ws.Rows(insRow)
            .Font.Bold = True
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
        End With
    Next k

    MsgBox "Готово. Сформировано/обновлено " & grpCount & " итогов.", vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ====== Вспомогательные ======

Private Function NewDoubles(n As Long) As Variant
    Dim a() As Double: ReDim a(1 To n)
    NewDoubles = a
End Function

Private Function NewBools(n As Long) As Variant
    Dim a() As Boolean: ReDim a(1 To n)
    NewBools = a
End Function

Private Function NewStrings(n As Long) As Variant
    Dim a() As String: ReDim a(1 To n)
    NewStrings = a
End Function

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    CleanText = LCase$(s)
End Function

Private Function IsNumericHeader(ByVal header As String) As Boolean
    IsNumericHeader = (Right$(LCase$(Trim(header)), Len(" : double")) = " : double")
End Function

Private Function FindHeader(ByRef headers() As String, ByVal name As String) As Long
    Dim j As Long
    For j = LBound(headers) To UBound(headers)
        If CleanText(headers(j)) = CleanText(name) Then
            FindHeader = j: Exit Function
        End If
    Next j
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
    Dim i As Long
    For i = 1 To cnt
        If keys(i) = key Then
            FindGroupIndex = i
            Exit Function
        End If
    Next i
    FindGroupIndex = 0
End Function

Private Function FindItogoRow(ws As Worksheet, typeCol As Long, grpName As String) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row
    For r = 2 To lastRow
        If Left$(Trim$(CStr(ws.Cells(r, typeCol).Value)), 6) = "Итого:" Then
            If InStr(1, ws.Cells(r, typeCol).Value, grpName, vbTextCompare) > 0 Then
                FindItogoRow = r
                Exit Function
            End If
        End If
    Next r
    FindItogoRow = 0
End Function

Private Function FindFirstRowOfGroup(ws As Worksheet, typeCol As Long, demCol As Long, crCol As Long, _
                                     grpName As String, demNorm As String, crNorm As String) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, typeCol).Value)) = grpName Then
            If CleanText(ws.Cells(r, demCol).Value) = demNorm And _
               CleanText(ws.Cells(r, crCol).Value) = crNorm Then
                FindFirstRowOfGroup = r
                Exit Function
            End If
        End If
    Next r
    FindFirstRowOfGroup = 0
End Function
