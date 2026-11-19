Attribute VB_Name = "WRPerformance"
Option Explicit
Option Private Module

Public Sub BuildReport(ByVal wb As Workbook)
    Dim loc As Worksheet, idx As Worksheet, ws As Worksheet
    Dim coords As Object, excluded As Object, cats As Object
    Dim masks As Object, patterns As Collection
    Dim seenPatterns As Object, phaseNames As Object
    Dim data As Variant, key As Variant, part As Variant, totals As Variant
    Dim labels As Variant, weights As Variant, color As Variant
    Dim last As Long, i As Long, j As Long, phase As Long, phases As Long
    Dim category As Long, top As Long, col As Long, columns As Long, totalDie As Long
    Dim chip As String, scope As String, pattern As Variant, wt As String
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
    Set phaseNames = CreateObject("Scripting.Dictionary")

    last = idx.Cells(idx.Rows.Count, 46).End(xlUp).Row
    If last >= 2 Then
        data = idx.Cells(2, 46).Resize(last - 1, 8).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = "Wafer Maps" And CStr(data(i, 2)) = "QUAD" Then
                pattern = CStr(data(i, 4))
                If Not seenPatterns.Exists(pattern) Then
                    seenPatterns.Add pattern, True: patterns.Add pattern
                End If
                phaseNames(CStr(data(i, 3))) = True
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
                If CStr(data(i, 1)) = "Wafer Maps" And CStr(data(i, 8)) = "Monitor Address Count" Then
                    If IsError(data(i, 11)) Or Not IsNumeric(data(i, 11)) Then
                        Err.Raise vbObjectError + 581, , "Invalid four-quadrant Map count."
                    End If
                    If CDbl(data(i, 11)) < 0 Then Err.Raise vbObjectError + 581, , "Negative four-quadrant Map count."
                    If CDbl(data(i, 11)) > 0 Then
                        chip = CStr(CLng(data(i, 6))) & "|" & CStr(CLng(data(i, 7)))
                        If Not coords.Exists(chip) Then Err.Raise vbObjectError + 581, , "Map Chip is absent from Raw locations."
                        If Not excluded.Exists(chip) Then
                            Select Case CStr(data(i, 9))
                                Case "Write Fail": category = 1
                                Case "Read Fail": category = 2
                                Case "Write&Read Fail": category = 4
                                Case "No monitor item Fail": category = 8
                                Case Else: Err.Raise vbObjectError + 581, , "Unknown four-quadrant category."
                            End Select
                            scope = CStr(data(i, 4)) & "|" & CStr(data(i, 5)) & "|" & chip
                            If masks.Exists(scope) Then category = category Or CLng(masks(scope))
                            masks(scope) = category
                        End If
                    End If
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
    ws.Columns(1).ColumnWidth = 25
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 12
    ws.Cells(1, 4).Resize(, columns - 3).EntireColumn.ColumnWidth = 11
    color = Array(RGB(255, 217, 102), vbWhite, RGB(255, 102, 0), RGB(227, 231, 235))
    top = 1
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
            If phaseNames.Exists(wt) Then
                For Each key In coords
                    scope = wt & "|" & CStr(pattern) & "|" & CStr(key)
                    If masks.Exists(scope) Then
                        For j = 0 To 3
                            If (CLng(masks(scope)) And CLng(weights(j))) <> 0 Then totals(j) = CLng(totals(j)) + 1
                        Next j
                    End If
                Next key
            End If
            For j = 0 To 3
                With ws.Cells(top + 1, col + j)
                    .Value2 = labels(j): .Interior.Color = color(j): .Font.Bold = True
                    If j = 3 Then .Font.Color = RGB(102, 102, 102)
                End With
                With ws.Cells(top + 2, col + j)
                    .NumberFormat = "0.0%"
                    If totalDie > 0 And phaseNames.Exists(wt) Then
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
        .Merge: .Value2 = "Total Die = All Wafer Chips - DC Fail Chips (WT1 Cat)"
    End With
    With ws.Cells(top + 1, 1).Resize(1, columns)
        .Merge: .Value2 = "Excluded Cat = C10 / C11 / C14 / C15 / C16 / C17"
    End With
    With ws.Cells(top + 2, 1).Resize(1, columns)
        .Merge: .Value2 = "Category (%) = Chips with category count > 0 / Total Die. A Chip may appear in multiple categories."
        .WrapText = True
    End With
    ws.Rows(top + 2).RowHeight = 32
    ws.PageSetup.PrintArea = ws.Cells(1, 1).Resize(top + 2, columns).Address
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
End Sub

Private Function ReportSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, anchor As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("W&R Performance")
    Set anchor = wb.Worksheets("Wafer Maps")
    On Error GoTo 0
    If anchor Is Nothing Then Set anchor = wb.Worksheets("Config")
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=anchor)
        ws.Name = "W&R Performance"
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.UnMerge: ws.Cells.Clear
    Set ReportSheet = ws
End Function

Public Sub Invalidate(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("W&R Performance")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Cells.UnMerge: ws.Cells.Clear
    With ws.Range("A1:G2")
        .Merge: .Value2 = "Source updated. Run Import and Analyze or Refresh Maps to update W/R Performance."
        .Font.Name = "Calibri": .Font.Size = 11: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub
