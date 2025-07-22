Attribute VB_Name = "Module1"
' ------------------------------------------------------------------------------
' Subroutine: ResetSettings
' Description: Restores common Excel settings to default for troubleshooting.
' ------------------------------------------------------------------------------

Sub ResetSettings()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.UsedRange ' Refresh UsedRange reference
End Sub


' ------------------------------------------------------------------------------
' Subroutine: TaxonomyKeyParse
' Description: Converts client taxonomy strings (formatted as KEY~Value_) into a
' structured table, separating keys into columns.
' ------------------------------------------------------------------------------

Sub TaxonomyKeyParse()
    If Selection.Cells.Count = 0 Then
        MsgBox "Select one or more cells with taxonomy strings.", vbExclamation
        Exit Sub
    End If

    Dim cell As Range, segs As Variant, kv() As String
    Dim dictHeaders As Object: Set dictHeaders = CreateObject("Scripting.Dictionary")
    Dim dataRows() As Object
    Dim i As Long, actualRowCount As Long

    ReDim dataRows(1 To Selection.Cells.Count)
    actualRowCount = 0

    ' Parse each selected cell
    For Each cell In Selection.Cells
        If Trim(cell.Value) <> "" Then
            actualRowCount = actualRowCount + 1
            Set dataRows(actualRowCount) = CreateObject("Scripting.Dictionary")
            segs = Split(cell.Value, "_")
            For i = 0 To UBound(segs)
                If InStr(segs(i), "~") > 0 Then
                    kv = Split(segs(i), "~", 2)
                    If Not dictHeaders.Exists(kv(0)) Then dictHeaders.Add kv(0), dictHeaders.Count + 1
                    dataRows(actualRowCount).Add kv(0), kv(1)
                End If
            Next i
        End If
    Next cell

    If actualRowCount = 0 Then
        MsgBox "No valid taxonomy strings found.", vbExclamation
        Exit Sub
    End If

    ' Generate new worksheet for parsed output
    Dim wsName As String: wsName = "Key Parse"
    Dim N As Long: N = 0
    Do While SheetExists(wsName)
        N = N + 1
        wsName = "Key Parse-" & N
    Loop
    Dim ws As Worksheet: Set ws = Worksheets.Add
    ws.Name = wsName

    ' Write headers
    Dim headerKeys() As Variant: headerKeys = dictHeaders.Keys
    For i = 0 To UBound(headerKeys)
        ws.Cells(1, i + 1).Value = headerKeys(i)
    Next i

    ' Write each parsed data row
    Dim r As Long, c As Long
    For r = 1 To actualRowCount
        For c = 0 To UBound(headerKeys)
            If dataRows(r).Exists(headerKeys(c)) Then
                ws.Cells(r + 1, c + 1).Value = dataRows(r)(headerKeys(c))
            End If
        Next c
    Next r

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    ws.Activate
End Sub

' Check if a worksheet with the given name exists
Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ------------------------------------------------------------------------------
' Subroutine: MergeExcelSheets
' Description: Merges data (excluding headers after first file) from selected
' workbooks and worksheets into one master workbook with formatting.
' ------------------------------------------------------------------------------

Sub MergeExcelSheets()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fd As FileDialog, f As Variant, wbSrc As Workbook, wbDest As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet, hdrDone As Boolean
    Dim lastDestRow As Long, rngData As Range
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = True
    If fd.Show = 0 Then GoTo CleanExit

    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Sheets(1)
    wsDest.Name = "Merged Data"
    
    For Each f In fd.SelectedItems
        Set wbSrc = Workbooks.Open(f)
        If wbSrc.ProtectStructure Then wbSrc.Windows(1).WindowState = xlNormal
        
        For Each wsSrc In wbSrc.Worksheets
            If LCase(wsSrc.Name) <> "help" Then
                With wsSrc.UsedRange
                    If Not hdrDone Then
                        .Rows(1).Copy wsDest.Range("A1")
                        hdrDone = True
                    End If
                    If .Rows.Count > 1 Then
                        Set rngData = .Offset(1).Resize(.Rows.Count - 1)
                        lastDestRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
                        rngData.Copy wsDest.Cells(lastDestRow, 1)
                    End If
                End With
            End If
        Next wsSrc
        wbSrc.Close False
    Next f

    ' Apply header formatting
    With wsDest
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(242, 242, 242)
        .Rows(1).HorizontalAlignment = xlCenter
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).AutoFilter
        .Columns.AutoFit
    End With

    With ActiveWindow
        .SplitRow = 1
        .SplitColumn = 0
        .FreezePanes = True
    End With

    Call QuickSave

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


' ------------------------------------------------------------------------------
' Subroutine: FormatCSV
' Description: Applies consistent easy-to-read formatting to header row for most SQL tables
' exported from Snowflake
' ------------------------------------------------------------------------------

Sub FormatCSV()
    Dim hdr As Range, c As Range
    Set hdr = Range("A1", Cells(1, Cells(1, Columns.Count).End(xlToLeft).Column))
    
    If Not hdr.Worksheet.AutoFilterMode Then hdr.AutoFilter

    With hdr
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Rows.RowHeight = 20
    End With

    For Each c In hdr
        c.EntireColumn.AutoFit
        c.ColumnWidth = Len(c.Value) + 2
    Next c

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    Call QuickSave
End Sub

' ------------------------------------------------------------------------------
' Subroutine: QuickSave
' Description: Instantly saves active workbook as .xlsx preserving formatting in
' Downloads folder. If fails, prompts manual save dialog.
' ------------------------------------------------------------------------------

Sub QuickSave()
    On Error GoTo eh1
    Dim savePath As String
    savePath = Environ$("USERPROFILE") & "\Downloads\"
    ActiveWorkbook.SaveAs Filename:=savePath & _
        Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".")) & _
        Format(Now, "MM.DD.YYYY") & ".xlsx", FileFormat:=51
    GoTo skipeh

eh1:
    Dim fNameAndPath As Variant
    MsgBox "Unable to auto-name. Please manually save as .xlsx.", vbInformation
    fNameAndPath = Application.GetSaveAsFilename( _
        InitialFileName:=ThisWorkbook.Path, _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Save As")
    If fNameAndPath = False Then Exit Sub
    ActiveWorkbook.SaveAs Filename:=fNameAndPath, FileFormat:=xlOpenXMLWorkbook

skipeh:
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Subroutine: UnprotectSheets
' Description: Opens any Excel files stuck in Protected View so other macros
' can operate without interruption.
' ------------------------------------------------------------------------------

Sub UnprotectSheets()
    Dim wb As Workbook
    Dim wbPV As ProtectedViewWindow

    If Application.ProtectedViewWindows.Count > 0 Then
        For Each wbPV In Application.ProtectedViewWindows
            wbPV.Activate
            Set wb = wbPV.Edit()
        Next wbPV
    End If

    Set wb = Nothing
End Sub

' ------------------------------------------------------------------------------
' Subroutine: ReplaceTildes
' Description: Replaces common instances of double tildes (~~) in the active range.
' for quick Excel lookups and find dialogues
' ------------------------------------------------------------------------------

Sub ReplaceTildes()
    ' Ensure something is selected before running replace
    If Not TypeOf Selection Is Range Then
        MsgBox "Please select a valid range to perform replace.", vbExclamation
        Exit Sub
    End If

    ' Replace double tildes with themselves (appears to be a placeholder action)
    Selection.Replace What:="~~", Replacement:="~~", LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
End Sub

