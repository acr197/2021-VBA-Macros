' VBA Automation Suite
' This suite contains various Excel macros to enhance efficiency, formatting, and data management.

Sub RestoreExcelDefaults()
    ' Quickly reset Excel configurations to default.
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.UsedRange
End Sub

Sub ParseTaxonomyStrings()
    ' Converts taxonomy strings (KEY~Value_) into structured Excel tables.
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select cells containing taxonomy strings.", vbExclamation
        Exit Sub
    End If

    Dim entry As Range, segments As Variant, pair() As String
    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    Dim entries() As Object
    Dim idx As Long, rowCount As Long

    ReDim entries(1 To Selection.Cells.Count)
    rowCount = 0

    For Each entry In Selection.Cells
        If Trim(entry.Value) <> "" Then
            rowCount = rowCount + 1
            Set entries(rowCount) = CreateObject("Scripting.Dictionary")
            segments = Split(entry.Value, "_")
            For idx = 0 To UBound(segments)
                If InStr(segments(idx), "~") > 0 Then
                    pair = Split(segments(idx), "~", 2)
                    If Not headers.Exists(pair(0)) Then headers.Add pair(0), headers.Count + 1
                    entries(rowCount).Add pair(0), pair(1)
                End If
            Next idx
        End If
    Next entry

    If rowCount = 0 Then
        MsgBox "No taxonomy entries found.", vbExclamation
        Exit Sub
    End If

    Dim newSheetName As String: newSheetName = "Parsed_Keys"
    Dim counter As Long: counter = 1
    While WorksheetExists(newSheetName)
        newSheetName = "Parsed_Keys_" & counter
        counter = counter + 1
    Wend

    Dim outSheet As Worksheet: Set outSheet = Worksheets.Add
    outSheet.Name = newSheetName

    Dim hdr() As Variant: hdr = headers.Keys
    For idx = 0 To headers.Count - 1
        outSheet.Cells(1, idx + 1).Value = hdr(idx)
    Next idx

    Dim r As Long, c As Long
    For r = 1 To rowCount
        For c = 0 To headers.Count - 1
            If entries(r).Exists(hdr(c)) Then
                outSheet.Cells(r + 1, c + 1).Value = entries(r)(hdr(c))
            End If
        Next c
    Next r

    With outSheet.Rows(1)
        .Font.Bold = True
        .EntireColumn.AutoFit
    End With
    outSheet.Activate
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Worksheets(sheetName) Is Nothing
End Function

Sub CombineWorkbooks()
    ' Combines data from multiple selected Excel files into one new workbook.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim fileDlg As FileDialog, file As Variant, sourceWB As Workbook, masterWB As Workbook
    Dim sourceWS As Worksheet, masterWS As Worksheet, headersDone As Boolean
    Dim lastRow As Long, copyRange As Range

    Set fileDlg = Application.FileDialog(msoFileDialogFilePicker)
    fileDlg.AllowMultiSelect = True
    If fileDlg.Show = 0 Then GoTo EndMerge

    Set masterWB = Workbooks.Add
    Set masterWS = masterWB.Sheets(1)
    masterWS.Name = "Combined_Data"

    For Each file In fileDlg.SelectedItems
        Set sourceWB = Workbooks.Open(file)
        For Each sourceWS In sourceWB.Worksheets
            If LCase(sourceWS.Name) <> "help" Then
                With sourceWS.UsedRange
                    If Not headersDone Then
                        .Rows(1).Copy masterWS.Range("A1")
                        headersDone = True
                    End If
                    If .Rows.Count > 1 Then
                        Set copyRange = .Offset(1).Resize(.Rows.Count - 1)
                        lastRow = masterWS.Cells(masterWS.Rows.Count, 1).End(xlUp).Row + 1
                        copyRange.Copy masterWS.Cells(lastRow, 1)
                    End If
                End With
            End If
        Next sourceWS
        sourceWB.Close False
    Next file

    With masterWS.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlCenter
        .AutoFilter
        .EntireColumn.AutoFit
    End With

    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    Call SaveWorkbookQuickly

EndMerge:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub FormatExportedData()
    ' Formats header rows for readability after importing external CSV or SQL outputs.
    Dim headers As Range, col As Range
    Set headers = Range("A1").CurrentRegion.Rows(1)

    headers.AutoFilter

    With headers
        .Interior.Color = RGB(230, 230, 230)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

    For Each col In headers.Cells
        col.EntireColumn.AutoFit
        col.ColumnWidth = Len(col.Value) + 3
    Next col

    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    Call SaveWorkbookQuickly
End Sub

Sub SaveWorkbookQuickly()
    ' Quickly saves workbook in user's downloads, fallback to manual save if error occurs.
    On Error GoTo ManualSave
    Dim quickPath As String
    quickPath = Environ("USERPROFILE") & "\Downloads\QuickSave_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"
    ActiveWorkbook.SaveAs quickPath, FileFormat:=xlOpenXMLWorkbook
    Exit Sub

ManualSave:
    Dim userFilePath As Variant
    userFilePath = Application.GetSaveAsFilename("", "Excel Files (*.xlsx), *.xlsx")
    If userFilePath <> False Then
        ActiveWorkbook.SaveAs userFilePath, xlOpenXMLWorkbook
    End If
End Sub

Sub OpenProtectedWorkbooks()
    ' Unprotects workbooks opened in protected view automatically.
    Dim pvWindow As ProtectedViewWindow

    For Each pvWindow In Application.ProtectedViewWindows
        pvWindow.Edit
    Next pvWindow
End Sub

Sub ReplacePlaceholderSymbols()
    ' Replaces special placeholder symbols to normalize content.
    If Not TypeOf Selection Is Range Then
        MsgBox "Select cells to run replacement.", vbExclamation
        Exit Sub
    End If

    Selection.Replace "~~", "~~", xlPart
End Sub
