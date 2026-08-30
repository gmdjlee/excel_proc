Attribute VB_Name = "modCSVImport"
Option Explicit

Private Const FONT_NAME As String = "맑은 고딕"
Private Const FONT_SIZE As Double = 11

Public LastError As String

' ===== Entry point (button click) =====
Public Sub RunImport()
    Dim csvPaths As Collection
    Set csvPaths = PickCSVFiles()
    If csvPaths.Count = 0 Then Exit Sub

    Dim titleText As Variant, labelText As Variant
    titleText = InputBox("A1 제목을 입력하세요 (선택한 모든 파일에 동일하게 적용됩니다)", "제목 입력", "8/28 111G")
    If titleText = "" Then Exit Sub
    labelText = InputBox("Q1 라벨을 입력하세요 (선택한 모든 파일에 동일하게 적용됩니다)", "라벨 입력", "XX")
    If labelText = "" Then Exit Sub

    Dim arr() As String
    ReDim arr(1 To csvPaths.Count)
    Dim idx As Long
    idx = 1
    Dim p As Variant
    For Each p In csvPaths
        arr(idx) = CStr(p)
        idx = idx + 1
    Next p

    MsgBox RunBatch(arr, CStr(titleText), CStr(labelText)), vbInformation
End Sub

Private Function PickCSVFiles() As Collection
    Dim result As New Collection
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "CSV 파일 선택 (여러 개 선택 가능)"
    fd.Filters.Clear
    fd.Filters.Add "CSV 파일", "*.csv"
    fd.AllowMultiSelect = True
    If fd.Show = -1 Then
        Dim item As Variant
        For Each item In fd.SelectedItems
            result.Add item
        Next item
    End If
    Set PickCSVFiles = result
End Function

' ===== Batch core: callable directly with an array of paths, for automated testing =====
' csvPaths accepts a Variant so it works both with a native VBA array (from RunImport)
' and a COM-marshaled array of variants (from external automation/testing)
Public Function RunBatch(csvPaths As Variant, titleText As String, labelText As String) As String
    Dim successCount As Long, failMsgs As String
    Dim i As Long
    For i = LBound(csvPaths) To UBound(csvPaths)
        Dim csvPath As String
        csvPath = CStr(csvPaths(i))
        Dim outputPath As String
        outputPath = DirOf(csvPath) & "\" & BaseName(csvPath) & "_result.xlsx"
        If BuildFromCSV(csvPath, titleText, labelText, outputPath) Then
            successCount = successCount + 1
        Else
            failMsgs = failMsgs & vbCrLf & BaseName(csvPath) & ": " & LastError
        End If
    Next i

    Dim total As Long
    total = UBound(csvPaths) - LBound(csvPaths) + 1
    RunBatch = successCount & " / " & total & " 개 완료" & failMsgs
End Function

Private Function BaseName(p As String) As String
    Dim n As String
    n = Mid(p, InStrRev(p, "\") + 1)
    If InStrRev(n, ".") > 0 Then n = Left(n, InStrRev(n, ".") - 1)
    BaseName = n
End Function

Private Function DirOf(p As String) As String
    DirOf = Left(p, InStrRev(p, "\") - 1)
End Function

Public Function GetLastError() As String
    GetLastError = LastError
End Function

' ===== Diagnostics: exercise ParseCSV alone without touching the UI =====
Public Function TestParseCSV(csvPath As String) As String
    On Error GoTo EH
    Dim data() As Variant
    Dim n As Long
    If Not ParseCSV(csvPath, data, n) Then
        TestParseCSV = "PARSE FAILED (row shape mismatch)"
        Exit Function
    End If
    TestParseCSV = "OK n=" & n & " row1=" & data(1, 1) & "," & data(1, 2) & " rowN_last=" & data(n, 12)
    Exit Function
EH:
    TestParseCSV = "ERROR " & Err.Number & ": " & Err.Description
End Function

' ===== Core builder: callable directly (dialogs bypassed) for automated testing =====
Public Function BuildFromCSV(csvPath As String, titleText As String, labelText As String, outputPath As String) As Boolean
    On Error GoTo EH

    Dim data() As Variant
    Dim n As Long
    If Not ParseCSV(csvPath, data, n) Then
        LastError = "CSV 형식 오류: 각 행이 정확히 12개의 값을 가져야 합니다."
        BuildFromCSV = False
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = "Sheet1"

    Application.ScreenUpdating = False
    BuildSheet ws, data, n, titleText, labelText
    Application.ScreenUpdating = True

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    BuildFromCSV = True
    Exit Function

EH:
    LastError = "오류 " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    BuildFromCSV = False
End Function

Private Function ReadFileAsUTF8(path As String) As String
    ' ADODB.Stream auto-detects/strips a UTF-8 BOM, unlike plain Line Input
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadFileAsUTF8 = stream.ReadText(-1) ' adReadAll
    stream.Close
End Function

Private Function ParseCSV(path As String, ByRef data() As Variant, ByRef n As Long) As Boolean
    Dim allText As String
    allText = Replace(ReadFileAsUTF8(path), vbCr, "")

    Dim lines() As String
    lines = Split(allText, vbLf)

    Dim rows As Collection
    Set rows = New Collection
    Dim i As Long, t As String
    For i = LBound(lines) To UBound(lines)
        t = Trim(lines(i))
        ' Tolerate trailing empty columns: a line-ending comma, or one with only spaces after it
        Do While Len(t) > 0 And (Right(t, 1) = "," Or Right(t, 1) = " " Or Right(t, 1) = vbTab)
            t = Left(t, Len(t) - 1)
        Loop
        If t <> "" Then rows.Add t
    Next i

    n = rows.Count
    If n = 0 Then
        ParseCSV = False
        Exit Function
    End If

    ReDim data(1 To n, 1 To 12)
    Dim r As Long, c As Long
    Dim parts() As String
    For r = 1 To n
        parts = Split(rows(r), ",")
        If (UBound(parts) - LBound(parts) + 1) <> 12 Then
            ParseCSV = False
            Exit Function
        End If
        For c = 1 To 12
            data(r, c) = CDbl(Trim(parts(c - 1)))
        Next c
    Next r

    ParseCSV = True
End Function

' ===== Layout builder =====
Private Sub BuildSheet(ws As Worksheet, data() As Variant, n As Long, titleText As String, labelText As String)
    ' Match the Normal style font so ColumnWidth (measured in default-font units) lines up with the template
    ws.Parent.Styles("Normal").Font.Name = FONT_NAME
    ws.Parent.Styles("Normal").Font.Size = FONT_SIZE

    Dim colLetters As Variant
    colLetters = Array("C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")

    ' Row 1: title (A1:O1) + label (Q1:R1)
    ws.Range("A1:O1").Merge
    ws.Range("A1").Value = titleText
    StyleAccent2 ws.Range("A1:O1")

    ws.Range("Q1:R1").Merge
    ws.Range("Q1").Value = labelText
    StyleGray ws.Range("Q1:R1")

    ' Row 2: A/B headers, 6 pairs across C..N
    Dim idx As Long
    For idx = 0 To 11
        Dim hc As Range
        Set hc = ws.Range(colLetters(idx) & "2")
        hc.Value = IIf(idx Mod 2 = 0, "A", "B")
        StyleAccent2 hc
    Next idx
    ws.Range("Q2").Value = "A"
    StyleAccent5 ws.Range("Q2")
    ws.Range("R2").Value = "B"
    StyleAccent4 ws.Range("R2")

    ' Main grid: each CSV row -> [index header row (1..12)] + [data row], starting row 3
    Dim i As Long, c As Long, headerRow As Long, dataRow As Long
    For i = 1 To n
        headerRow = 3 + (i - 1) * 2
        dataRow = headerRow + 1
        For c = 1 To 12
            Dim colLetter As String
            colLetter = colLetters(c - 1)
            ws.Range(colLetter & headerRow).Value = c
            StyleAccent2 ws.Range(colLetter & headerRow)
            ws.Range(colLetter & dataRow).Value = data(i, c)
            StyleData ws.Range(colLetter & dataRow)
        Next c
        ' O column: blank spacer that mirrors the row-type look
        StyleAccent2 ws.Range("O" & headerRow)
        StyleData ws.Range("O" & dataRow)
    Next i

    ' Column B: group label every 2 CSV rows (merged)
    Dim totalBGroups As Long
    totalBGroups = ((n - 1) \ 2) + 1
    Dim g As Long, bFirst As Long, bLast As Long, rFrom As Long, rTo As Long
    For g = 1 To totalBGroups
        bFirst = (g - 1) * 2 + 1
        bLast = WorksheetFunction.Min(g * 2, n)
        rFrom = 3 + (bFirst - 1) * 2
        rTo = 3 + (bLast - 1) * 2 + 1
        ws.Range("B" & rFrom & ":B" & rTo).Merge
        ws.Range("B" & rFrom).Value = g
        StyleAccent2 ws.Range("B" & rFrom & ":B" & rTo)
    Next g

    ' Column A: group label every 4 CSV rows (merged)
    Dim totalAGroups As Long
    totalAGroups = ((n - 1) \ 4) + 1
    Dim aFirst As Long, aLast As Long
    For g = 1 To totalAGroups
        aFirst = (g - 1) * 4 + 1
        aLast = WorksheetFunction.Min(g * 4, n)
        rFrom = 3 + (aFirst - 1) * 2
        rTo = 3 + (aLast - 1) * 2 + 1
        ws.Range("A" & rFrom & ":A" & rTo).Merge
        ws.Range("A" & rFrom).Value = g
        StyleAccent2 ws.Range("A" & rFrom & ":A" & rTo)
    Next g

    ' Side table P/Q/R: P = sequential 1..6n; Q = odd-column values sorted desc; R = even-column values sorted desc
    Dim totalSideRows As Long
    totalSideRows = 6 * n

    StyleGray ws.Range("P1")
    StyleGray ws.Range("P2")

    Dim k As Long, cnt As Long
    cnt = 0
    For k = 1 To 6
        For i = 1 To n
            cnt = cnt + 1
            ws.Range("T" & (2 + cnt)).Value = data(i, 2 * k - 1)
            ws.Range("U" & (2 + cnt)).Value = data(i, 2 * k)
        Next i
    Next k

    ws.Range("T3:T" & (2 + totalSideRows)).Sort Key1:=ws.Range("T3"), Order1:=xlDescending, Header:=xlNo
    ws.Range("U3:U" & (2 + totalSideRows)).Sort Key1:=ws.Range("U3"), Order1:=xlDescending, Header:=xlNo

    Dim sideRow As Long
    For i = 1 To totalSideRows
        sideRow = 2 + i
        ws.Range("P" & sideRow).Value = i
        StyleGray ws.Range("P" & sideRow)
        ws.Range("Q" & sideRow).Value = ws.Range("T" & sideRow).Value
        StyleData ws.Range("Q" & sideRow)
        ws.Range("R" & sideRow).Value = ws.Range("U" & sideRow).Value
        StyleData ws.Range("R" & sideRow)
    Next i

    ws.Range("T3:U" & (2 + totalSideRows)).Clear

    ws.Columns("P").ColumnWidth = 4.5
    ws.UsedRange.RowHeight = 16.5
End Sub

' ===== Style helpers =====
Private Sub ApplyBaseStyle(rng As Range)
    With rng
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
    End With
End Sub

Private Sub StyleAccent2(rng As Range)
    ApplyBaseStyle rng
    rng.Interior.ThemeColor = xlThemeColorAccent2
    rng.Interior.TintAndShade = 0.6
End Sub

Private Sub StyleAccent5(rng As Range)
    ApplyBaseStyle rng
    rng.Interior.ThemeColor = xlThemeColorAccent5
    rng.Interior.TintAndShade = 0.6
End Sub

Private Sub StyleAccent4(rng As Range)
    ApplyBaseStyle rng
    rng.Interior.ThemeColor = xlThemeColorAccent4
    rng.Interior.TintAndShade = 0.8
End Sub

Private Sub StyleGray(rng As Range)
    ApplyBaseStyle rng
    rng.Interior.ThemeColor = xlThemeColorDark1
    rng.Interior.TintAndShade = -0.15
End Sub

Private Sub StyleData(rng As Range)
    ApplyBaseStyle rng
    rng.Interior.Pattern = xlNone
End Sub
