Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim detailMessage As String
    If Cancel Then Exit Sub
    'The sheet event can handle the click first. Fall back only while the clicked sheet is still active.
    If Not Application.ActiveSheet Is Sh Then Exit Sub
    If BLKVminEngine.BLK_TryOpenCatMap(Me, Sh, Target) Then
        Cancel = True
    ElseIf BLKVminEngine.BLK_IsVoltageGridCell(Sh, Target) Then
        Cancel = True
        detailMessage = BLKVminEngine.BLK_ShowBLKVoltage(Me, Sh, Target)
        If Left$(detailMessage, 6) = "Error:" Then MsgBox detailMessage, vbExclamation
    ElseIf BLKVminEngine.BLK_IsIndexedMapCell(Me, Sh, Target) Then
        Cancel = True
        BLKVminEngine.BLK_OpenIndexedMap Me, Sh, Target
    ElseIf ImportFailAddresses.IsMonitorCountCell(Sh, Target) Then
        Cancel = True
        ImportFailAddresses.OpenMonitorCountDetails Me, Sh, Target
    ElseIf ImportFailAddresses.IsMapFailCell(Sh, Target) Then
        Cancel = True
        ImportFailAddresses.OpenMapChipDetails Me, Sh, Target
    End If
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If IsSourceRowCell(Sh, Target) Then JumpToSourceRow Target
End Sub

Private Function IsSourceRowCell(ByVal sheet As Object, ByVal target As Range) As Boolean
    Dim table As ListObject, column As ListColumn, body As Range
    On Error GoTo Done
    If Not TypeOf sheet Is Worksheet Then Exit Function
    If target.CountLarge <> 1 Then Exit Function
    If sheet.Name <> "Fail Analysis" And sheet.Name <> "Chip Fail Details" Then Exit Function
    If Not IsNumeric(target.Value2) Then Exit Function
    If CDbl(target.Value2) < 2 Then Exit Function
    For Each table In sheet.ListObjects
        For Each column In table.ListColumns
            If column.Name = "Source Row" Or column.Name = "Source Address Row" Then
                Set body = column.DataBodyRange
                If Not body Is Nothing Then
                    If Not Application.Intersect(target, body) Is Nothing Then
                        IsSourceRowCell = True
                        Exit Function
                    End If
                End If
            End If
        Next column
    Next table
Done:
End Function

Private Sub JumpToSourceRow(ByVal target As Range)
    Dim eventsEnabled As Boolean, source As Worksheet, sourceRow As Long
    eventsEnabled = Application.EnableEvents
    On Error GoTo RestoreEvents
    sourceRow = CLng(target.Value2)
    Set source = Me.Worksheets("Fail Address")
    If sourceRow < 2 Or sourceRow > source.Cells(source.Rows.Count, 1).End(xlUp).Row Then Exit Sub
    Application.EnableEvents = False
    Application.Goto source.Cells(sourceRow, 1), True
RestoreEvents:
    Application.EnableEvents = eventsEnabled
End Sub
