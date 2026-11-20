Attribute VB_Name = "WRPerformance"
Option Explicit
Option Private Module

Private Const SCAN_REPORT As String = "Scan W-R Performance"

Public Sub BuildReport(ByVal wb As Workbook)
    Dim loc As Worksheet, idx As Worksheet, ws As Worksheet
    Dim coords As Object, excluded As Object, cats As Object
    Dim masks As Object, patterns As Collection
    Dim seenPatterns As Object, panels As Object, scanPanels As Object, scanUnavailable As Object
    Dim data As Variant, key As Variant, totals As Variant
    Dim labels As Variant, weights As Variant, color As Variant
    Dim last As Long, i As Long, j As Long, phase As Long, phases As Long
    Dim category As Long, top As Long, col As Long, columns As Long, totalDie As Long
    Dim combinationCounts(0 To 15) As Long, detailTop As Long, chartCount As Long
    Dim chip As String, scope As String, panel As String, pattern As Variant, wt As String
    Dim lot As String, wafer As String, manifestRow As Long

    Set loc = wb.Worksheets("Chip Locations"): Set idx = wb.Worksheets("Import Index")
    Set coords = CreateObject("Scripting.Dictionary")
    last = loc.Cells(loc.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then
        data = loc.Range("A2:B" & last).Value2
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 And Len(CStr(data(i, 2))) > 0 Then
                chip = CStr(CLng(data(i, 1))) & "|" & CStr(CLng(data(i, 2)))
                coords(chip) = True
            End If
        Next i
    End If
    Set cats = BLKVminEngine.BLK_Categories(wb, "WT1")
    Set excluded = CreateObject("Scripting.Dictionary")
    For Each key In cats
        Select Case CLng(cats(key))
            Case 10, 11, 14, 15, 16, 17
                If coords.Exists(CStr(key)) Then excluded(CStr(key)) = True
        End Select
    Next key
    totalDie = coords.Count - excluded.Count
    phases = BLKVminEngine.BLK_PhaseCount(wb)
    Set masks = CreateObject("Scripting.Dictionary")
    Set seenPatterns = CreateObject("Scripting.Dictionary")
    Set patterns = New Collection
    Set panels = CreateObject("Scripting.Dictionary")
    Set scanPanels = CreateObject("Scripting.Dictionary")
    Set scanUnavailable = CreateObject("Scripting.Dictionary")

    last = idx.Cells(idx.Rows.Count, 46).End(xlUp).Row
    If last >= 2 Then
        data = idx.Cells(2, 46).Resize(last - 1, 8).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = "Wafer Maps" And CStr(data(i, 2)) = "QUAD" Then
                pattern = CStr(data(i, 4))
                If Not seenPatterns.Exists(pattern) Then
                    seenPatterns.Add pattern, True: patterns.Add pattern
                End If
                wt = CStr(data(i, 3))
                panels(wt & "|" & pattern) = True
                If StrComp(Trim$(CStr(pattern)), "Scan", vbTextCompare) = 0 Then scanPanels(wt) = pattern
            End If
        Next i
    End If
    labels = Array("W", "R", "RW", "X"): weights = Array(1, 2, 4, 8)
    last = idx.Cells(idx.Rows.Count, 33).End(xlUp).Row
    If last >= 2 Then
        For top = 2 To last Step 2048
            i = last - top + 1: If i > 2048 Then i = 2048
            data = idx.Cells(top, 33).Resize(i, 12).Value2
            For i = 1 To UBound(data, 1)
                If CStr(data(i, 1)) = "Wafer Maps" Then
                    wt = CStr(data(i, 4)): pattern = CStr(data(i, 5))
                    Select Case CStr(data(i, 8))
                        Case "Unavailable"
                            If scanPanels.Exists(wt) And CStr(data(i, 12)) = "Excluded by Cat" Then
                                If CStr(scanPanels(wt)) = pattern Then
                                    chip = CStr(CLng(data(i, 6))) & "|" & CStr(CLng(data(i, 7)))
                                    scanUnavailable(wt & "|" & chip) = True
                                End If
                            End If
                        Case "Monitor Address Count"
                            If IsError(data(i, 11)) Or Not IsNumeric(data(i, 11)) Then
                                Err.Raise vbObjectError + 581, , "Invalid four-quadrant Map count."
                            End If
                            If CDbl(data(i, 11)) < 0 Then Err.Raise vbObjectError + 581, , "Negative four-quadrant Map count."
                            If CDbl(data(i, 11)) > 0 Then
                                chip = CStr(CLng(data(i, 6))) & "|" & CStr(CLng(data(i, 7)))
                                If Not coords.Exists(chip) Then Err.Raise vbObjectError + 581, , "Map Chip is absent from Raw locations."
                                If Not excluded.Exists(chip) Or StrComp(Trim$(CStr(pattern)), "Scan", vbTextCompare) = 0 Then
                                    Select Case CStr(data(i, 9))
                                        Case "Write Fail": category = 1
                                        Case "Read Fail": category = 2
                                        Case "Write&Read Fail": category = 4
                                        Case "No monitor item Fail": category = 8
                                        Case Else: Err.Raise vbObjectError + 581, , "Unknown four-quadrant category."
                                    End Select
                                    scope = wt & "|" & CStr(pattern) & "|" & chip
                                    If masks.Exists(scope) Then category = category Or CLng(masks(scope))
                                    masks(scope) = category
                                End If
                            End If
                    End Select
                End If
            Next i
        Next top
    End If

    manifestRow = 5
    lot = Trim$(CStr(idx.Cells(manifestRow, 10).Value2))
    wafer = Trim$(CStr(idx.Cells(manifestRow, 11).Value2))
    If Len(lot) = 0 Or Len(wafer) = 0 Then
        lot = Trim$(CStr(idx.Cells(2, 10).Value2))
        wafer = Trim$(CStr(idx.Cells(2, 11).Value2))
    End If
    lot = Left$(lot, 6)
    If IsNumeric(wafer) And Len(wafer) > 0 Then wafer = "#" & Format$(CLng(wafer), "00")
    Set ws = ReportSheet(wb)
    columns = 3 + 4 * phases
    If columns < 7 Then columns = 7
    With ws.Cells(1, 1).Resize(5 + 6 * patterns.Count, columns)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    ws.Columns(1).ColumnWidth = 29
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 17
    ws.Cells(1, 4).Resize(, columns - 3).EntireColumn.ColumnWidth = 11
    color = Array(RGB(255, 217, 102), vbWhite, RGB(255, 102, 0), RGB(227, 231, 235))
    top = 1
    detailTop = 5 + 5 * patterns.Count
    For Each pattern In patterns
        With ws.Cells(top, 1).Resize(3, columns)
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(194, 207, 218)
            .Borders.Weight = xlThin
        End With
        For j = 1 To 3
            ws.Cells(top, j).Resize(2, 1).Merge
        Next j
        ws.Cells(top, 1).Value2 = CStr(pattern) & vbLf & "W/R Performance"
        ws.Cells(top, 2).Value2 = "Wafer"
        ws.Cells(top, 3).Value2 = "Total Die"
        With ws.Cells(top, 1).Resize(2, 3)
            .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite
            .Font.Bold = True: .WrapText = True
        End With
        ws.Cells(top + 2, 1).Value2 = lot
        ws.Cells(top + 2, 2).Value2 = wafer
        ws.Cells(top + 2, 3).Value2 = totalDie
        For phase = 1 To phases
            wt = "WT" & phase: col = 4 + 4 * (phase - 1)
            With ws.Cells(top, col).Resize(1, 4)
                .Merge: .Value2 = wt & " Chips"
                .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
            End With
            totals = Array(0&, 0&, 0&, 0&)
            panel = wt & "|" & CStr(pattern)
            If panels.Exists(panel) Then
                Erase combinationCounts
                For Each key In coords
                    If Not excluded.Exists(key) Then
                        scope = panel & "|" & CStr(key)
                        category = 0
                        If masks.Exists(scope) Then category = CLng(masks(scope))
                        combinationCounts(category) = combinationCounts(category) + 1
                        For j = 0 To 3
                            If (category And CLng(weights(j))) <> 0 Then totals(j) = CLng(totals(j)) + 1
                        Next j
                    End If
                Next key
                WriteCountDetail ws, detailTop, CStr(pattern), wt, totals, combinationCounts, totalDie
                detailTop = detailTop + 24
                chartCount = chartCount + 1
            End If
            For j = 0 To 3
                With ws.Cells(top + 1, col + j)
                    .Value2 = labels(j): .Interior.Color = color(j): .Font.Bold = True
                    If j = 3 Then .Font.Color = RGB(102, 102, 102)
                End With
                With ws.Cells(top + 2, col + j)
                    .NumberFormat = "0.00%"
                    If totalDie > 0 And panels.Exists(panel) Then
                        .Value2 = CDbl(totals(j)) / totalDie
                    Else
                        .Value2 = "-"
                    End If
                    If j = 3 Then .Font.Color = RGB(102, 102, 102)
                End With
            Next j
        Next phase
        ws.Rows(top + 2).RowHeight = 28
        top = top + 5
    Next pattern
    If patterns.Count = 0 Then
        ws.Range("A1:G2").Merge: ws.Range("A1").Value2 = "W/R Performance - run Import and Analyze to populate Maps."
        ws.Range("A1").WrapText = True
        top = 4
    End If
    With ws.Cells(top, 1).Resize(1, columns)
        .Merge: .Value2 = "DC Pass Chips = Total Die = All Wafer Chips - excluded WT1 Cats"
    End With
    With ws.Cells(top + 1, 1).Resize(1, columns)
        .Merge: .Value2 = "Excluded Cat = C10 / C11 / C14 / C15 / C16 / C17"
    End With
    With ws.Cells(top + 2, 1).Resize(1, columns)
        .Merge: .Value2 = "W/R/RW/X count independently. Each combination requires exactly the listed Fail modes." & vbLf & _
                          "No recorded Fail can include Chips without usable panel data."
        .WrapText = True
    End With
    ws.Rows(top + 2).RowHeight = 42
    If chartCount > 0 Then
        top = detailTop - 3
        If columns < 15 Then columns = 15
    Else
        top = top + 2
    End If
    ws.PageSetup.PrintArea = ws.Cells(1, 1).Resize(top, columns).Address
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    WriteScanReport wb, ws, coords, masks, scanPanels, scanUnavailable, lot, wafer, phases
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
End Sub

Private Sub WriteScanReport(ByVal wb As Workbook, ByVal anchor As Worksheet, ByVal coords As Object, ByVal masks As Object, _
                            ByVal scanPanels As Object, ByVal scanUnavailable As Object, ByVal lot As String, _
                            ByVal wafer As String, ByVal phases As Long)
    Dim ws As Worksheet, counts(0 To 7) As Long, failCounts As Variant
    Dim phase As Long, top As Long, totalDie As Long, category As Long, j As Long
    Dim wt As String, panel As String, scope As String, key As Variant

    Set ws = ReportSheet(wb, SCAN_REPORT, anchor.Name)
    ws.Columns(1).ColumnWidth = 27
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 19
    ws.Range("D:J").ColumnWidth = 13
    top = 1
    For phase = 1 To phases
        wt = "WT" & phase
        If scanPanels.Exists(wt) Then
            Erase counts
            totalDie = 0
            panel = wt & "|" & CStr(scanPanels(wt))
            For Each key In coords
                If Not scanUnavailable.Exists(wt & "|" & CStr(key)) Then
                    totalDie = totalDie + 1
                    scope = panel & "|" & CStr(key)
                    category = 0
                    'Keep the Write, Read and Write&Read bits.
                    If masks.Exists(scope) Then category = CLng(masks(scope)) And 7
                    counts(category) = counts(category) + 1
                End If
            Next key
            failCounts = Array(counts(1) + counts(3) + counts(5) + counts(7), _
                               counts(2) + counts(3) + counts(6) + counts(7), _
                               counts(4) + counts(5) + counts(6) + counts(7), _
                               counts(3), counts(5), counts(6), counts(7))
            With ws.Cells(top, 1).Resize(5, 10)
                .Font.Name = "Calibri": .Font.Size = 11
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .RowHeight = 24
                .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(210, 223, 235)
            End With
            With ws.Cells(top, 1).Resize(1, 10)
                .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
            End With
            ws.Cells(top, 1).Resize(1, 3).Value2 = Array("Scan W/R Performance", "Wafer", "Total Die")
            With ws.Cells(top, 4).Resize(1, 7)
                .Merge: .Value2 = wt & " DC Pass Chips"
            End With
            With ws.Cells(top + 1, 1).Resize(1, 10)
                .Value2 = Array("LotID", "#Wfr", "DC Pass Count", "W", "R", "W/R", "W+R", "W+W/R", "R+W/R", "W+R+W/R")
                .Interior.Color = RGB(221, 235, 247): .Font.Bold = True
            End With
            ws.Cells(top + 2, 1).Resize(1, 3).Value2 = Array(lot, wafer, totalDie)
            ws.Cells(top + 3, 1).Value2 = "Fail Chips"
            ws.Cells(top + 3, 4).Resize(1, 7).Value2 = failCounts
            For j = 0 To 6
                With ws.Cells(top + 2, 4 + j)
                    .NumberFormat = "0.00%"
                    .Font.Bold = True
                    If totalDie > 0 Then .Value2 = CDbl(failCounts(j)) / totalDie Else .Value2 = "-"
                End With
            Next j
            With ws.Cells(top + 4, 1).Resize(1, 10)
                .Merge: .Value2 = "Raw Chips: " & coords.Count & "    |    Gray / Unavailable: " & (coords.Count - totalDie)
                .Interior.Color = RGB(241, 245, 249)
            End With
            top = top + 6
        End If
    Next phase
    If top = 1 Then
        With ws.Range("A1:J2")
            .Merge: .Value2 = "Scan W/R Performance - no Scan Wafer Map is available."
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
        top = 3
    Else
        With ws.Cells(top, 1).Resize(1, 10)
            .Merge: .Value2 = "W, R and W/R count independently. The + columns use exact combinations of these three modes."
            .HorizontalAlignment = xlCenter: .WrapText = True: .RowHeight = 24
        End With
        top = top + 1
    End If
    With ws.PageSetup
        .PrintArea = ws.Cells(1, 1).Resize(top - 1, 10).Address
        .Orientation = xlLandscape: .Zoom = False
        .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
End Sub

Private Sub WriteCountDetail(ByVal ws As Worksheet, ByVal top As Long, ByVal pattern As String, ByVal wt As String, _
                             ByVal totals As Variant, ByRef combinationCounts() As Long, ByVal totalDie As Long)
    Dim labels As Variant, order As Variant, values(1 To 20, 1 To 4) As Variant
    Dim i As Long, count As Long, first As Long
    Dim chartArea As Range, plot As ChartObject, counts As Series

    labels = Array("W", "R", "RW", "X", "W only", "R only", "RW only", "X only", _
                   "W + R", "W + RW", "W + X", "R + RW", "R + X", "RW + X", _
                   "W + R + RW", "W + R + X", "W + RW + X", "R + RW + X", _
                   "W + R + RW + X", "No recorded Fail (Chips)")
    order = Array(1, 2, 4, 8, 3, 5, 9, 6, 10, 12, 7, 11, 13, 14, 15, 0)
    For i = 1 To 20
        If i <= 4 Then
            count = CLng(totals(i - 1))
        Else
            count = combinationCounts(CLng(order(i - 5)))
        End If
        values(i, 1) = labels(i - 1)
        values(i, 2) = count
        values(i, 3) = totalDie
        values(i, 4) = "-"
        If totalDie > 0 Then values(i, 4) = CDbl(count) / totalDie
    Next i

    first = top + 2
    With ws.Cells(top, 1).Resize(22, 4)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 19
    End With
    With ws.Cells(top, 1).Resize(1, 4)
        .Merge: .Value2 = pattern & " / " & wt & " | Fail Count Detail"
        .ShrinkToFit = True
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
    End With
    With ws.Cells(top + 1, 1).Resize(1, 4)
        .Value2 = Array("Fail Mode", "Fail Count", "DC Pass Chips", "Share")
        .Interior.Color = RGB(221, 235, 247): .Font.Bold = True
    End With
    ws.Cells(first, 1).Resize(20, 4).Value2 = values
    ws.Cells(first, 4).Resize(20, 1).NumberFormat = "0.00%"
    ws.Cells(first, 1).Resize(4, 4).Interior.Color = RGB(255, 242, 204)
    ws.Cells(first + 4, 1).Resize(15, 4).Interior.Color = RGB(234, 241, 248)
    ws.Cells(first + 19, 1).Resize(1, 4).Interior.Color = RGB(227, 231, 235)

    Set chartArea = ws.Cells(top, 6).Resize(22, 10)
    Set plot = ws.ChartObjects.Add(chartArea.Left, chartArea.Top, chartArea.Width, chartArea.Height)
    With plot.Chart
        .ChartType = xlBarClustered
        Set counts = .SeriesCollection.NewSeries
        counts.Name = "Fail Count"
        counts.XValues = ws.Cells(first, 1).Resize(19, 1)
        counts.Values = ws.Cells(first, 2).Resize(19, 1)
        counts.Format.Fill.ForeColor.RGB = RGB(44, 100, 151)
        For i = 1 To 4
            counts.Points(i).Format.Fill.ForeColor.RGB = RGB(192, 137, 28)
        Next i
        counts.ApplyDataLabels xlDataLabelsShowValue
        counts.DataLabels.NumberFormat = "0;-0;;"
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = pattern & " / " & wt & " | Fail Count (Chips) / " & totalDie & " DC Pass Chips"
        With .Axes(xlCategory)
            .ReversePlotOrder = True
            .Crosses = xlMaximum
        End With
        With .Axes(xlValue)
            .MinimumScale = 0
            If .MajorUnit < 1 Then .MajorUnit = 1
            .TickLabels.NumberFormat = "0"
        End With
    End With
End Sub

Private Function ReportSheet(ByVal wb As Workbook, Optional ByVal sheetName As String = "W&R Performance", _
                             Optional ByVal anchorName As String = "Wafer Maps") As Worksheet
    Dim ws As Worksheet, anchor As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    Set anchor = wb.Worksheets(anchorName)
    On Error GoTo 0
    If anchor Is Nothing Then Set anchor = wb.Worksheets("Config")
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=anchor)
        ws.Name = sheetName
    End If
    ws.Visible = xlSheetVisible
    If ws.ChartObjects.Count > 0 Then ws.ChartObjects.Delete
    ws.Cells.UnMerge: ws.Cells.Clear
    Set ReportSheet = ws
End Function

Public Sub Invalidate(ByVal wb As Workbook)
    InvalidateSheet wb, "W&R Performance"
    InvalidateSheet wb, SCAN_REPORT
End Sub

Private Sub InvalidateSheet(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ChartObjects.Count > 0 Then ws.ChartObjects.Delete
    ws.Cells.UnMerge: ws.Cells.Clear
    ws.PageSetup.PrintArea = "$A$1:$G$2"
    With ws.Range("A1:G2")
        .Merge: .Value2 = "Source updated. Run Import and Analyze or Refresh Maps to update W/R Performance."
        .Font.Name = "Calibri": .Font.Size = 11: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub
