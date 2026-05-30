VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrosshairEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Application-level cursor crosshair: highlights the entire active row +
' active column in every open workbook using two Conditional Formatting
' rules. The active cell itself is excluded inline via a multiplication
' trick in the rule formula, so the selection indicator stays visible
' without needing a separate mask rule.
'
' Each rule's Formula1 carries a marker substring so we can identify and
' remove our rules reliably without holding (potentially stale)
' FormatCondition object references.
'
' All event handlers run behind a reentrancy guard (mInHandler) and
' validate mLastSheet via IsSheetAlive before touching it — the most
' common crash mode for Excel add-ins is event handlers re-entering
' themselves, or touching a worksheet reference whose workbook was just
' closed.

Private Const HIGHLIGHT_COLOR  As Long   = 13561798   ' RGB(198, 239, 206) mint green
Private Const CCH_ROW_MARKER   As String = "__CCH_ROW__"
Private Const CCH_COL_MARKER   As String = "__CCH_COL__"
Private Const CCH_MASK_MARKER  As String = "__CCH_MASK__"

Public WithEvents App As Excel.Application

Private mLastSheet As Worksheet
Private mInHandler As Boolean   ' reentrancy guard

Private Sub Class_Initialize()
    Set App = Excel.Application
End Sub

' Returns True only if the worksheet reference still points to a live
' Excel object. After a workbook closes its sheet references stay in
' module variables but accessing them can crash Excel; we use this to
' guard every touch of mLastSheet.
Private Function IsSheetAlive(ws As Worksheet) As Boolean
    On Error Resume Next
    Dim s As String
    s = ws.Name
    IsSheetAlive = (Err.Number = 0)
    Err.Clear
End Function

' Validate mLastSheet before any use; clear if dead.
Private Sub EnsureLastSheetValid()
    If mLastSheet Is Nothing Then Exit Sub
    If Not IsSheetAlive(mLastSheet) Then Set mLastSheet = Nothing
End Sub

' --- Application event handlers --------------------------------------------

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error GoTo Fail
    If Sh Is Nothing Or Target Is Nothing Then GoTo Done
    If TypeName(Sh) <> "Worksheet" Then GoTo Done

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    EnsureLastSheetValid
    If Not mLastSheet Is Nothing Then
        If Not (mLastSheet Is Sh) Then RemoveCrosshairOnSheet mLastSheet
    End If

    RemoveCrosshairOnSheet Sh
    ApplyCrosshair Sh, Target.Row, Target.Column
    Set mLastSheet = Sh

Done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInHandler = False
    Exit Sub
Fail:
    Resume Done
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Scrub leftover crosshair CFs from the active sheet only — that's the
    ' only sheet that could be showing one. Other sheets get cleaned the
    ' first time the user navigates to them (SheetActivate). Iterating
    ' every sheet on every open is too slow for multi-sheet workbooks.
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error Resume Next
    Application.EnableEvents = False
    If TypeName(Wb.ActiveSheet) = "Worksheet" Then RemoveCrosshairOnSheet Wb.ActiveSheet
    Application.EnableEvents = True
    mInHandler = False
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Paint the crosshair as soon as a workbook becomes active so the user
    ' doesn't have to click a cell to make it appear.
    PaintActiveCrosshair
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    ' Same idea when switching between sheets within a workbook: the new
    ' sheet has its own active cell that no SelectionChange has fired for.
    On Error Resume Next
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    PaintActiveCrosshair
End Sub

Private Sub PaintActiveCrosshair()
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error GoTo Fail
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then GoTo Done
    If Application.ActiveCell Is Nothing Then GoTo Done

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    EnsureLastSheetValid
    If Not mLastSheet Is Nothing Then
        If Not (mLastSheet Is ws) Then RemoveCrosshairOnSheet mLastSheet
    End If
    RemoveCrosshairOnSheet ws
    ApplyCrosshair ws, Application.ActiveCell.Row, Application.ActiveCell.Column
    Set mLastSheet = ws

Done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInHandler = False
    Exit Sub
Fail:
    Resume Done
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error Resume Next
    Application.EnableEvents = False
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        RemoveCrosshairOnSheet ws
    Next ws
    Application.EnableEvents = True
    mInHandler = False
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        ApplyCrosshair ActiveSheet, ActiveCell.Row, ActiveCell.Column
        Set mLastSheet = ActiveSheet
        Wb.Saved = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
    mInHandler = False
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If mInHandler Then Exit Sub
    mInHandler = True
    On Error Resume Next
    Application.EnableEvents = False

    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        RemoveCrosshairOnSheet ws
    Next ws

    ' If mLastSheet belongs to the closing workbook, clear it before its
    ' parent dies — otherwise the next event would touch a freed object
    ' and Excel can crash hard.
    If Not mLastSheet Is Nothing Then
        If mLastSheet.Parent Is Wb Then Set mLastSheet = Nothing
    End If

    Application.EnableEvents = True
    mInHandler = False
End Sub

' --- Crosshair add / remove -------------------------------------------------

Private Sub ApplyCrosshair(ws As Worksheet, rowNum As Long, colNum As Long)
    On Error Resume Next

    ' Preserve the workbook's Saved flag so painting the overlay never
    ' counts as a user edit (no false "save changes?" prompts on close).
    Dim wb As Workbook
    Set wb = ws.Parent
    Dim wasSaved As Boolean
    wasSaved = wb.Saved

    Dim fc As FormatCondition

    ' Row highlight on the entire active row. The active cell itself is
    ' excluded inline using a multiplication trick so we don't need a
    ' separate mask CF or SetFirstPriority call (both were expensive):
    '   (COLUMN()<>colNum)  -> 0 at active cell, 1 elsewhere in the row
    '   (LEN(marker)>0)     -> always 1; embeds our marker string for cleanup
    ' Product is 0 (FALSE / no highlight) at active cell, 1 elsewhere.
    Set fc = ws.Rows(rowNum).FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=(COLUMN()<>" & colNum & ")*(LEN(""" & CCH_ROW_MARKER & """)>0)")
    If Not fc Is Nothing Then
        fc.Interior.Color = HIGHLIGHT_COLOR
        fc.StopIfTrue = False
    End If

    ' Column highlight on the entire active column, same trick.
    Set fc = ws.Columns(colNum).FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=(ROW()<>" & rowNum & ")*(LEN(""" & CCH_COL_MARKER & """)>0)")
    If Not fc Is Nothing Then
        fc.Interior.Color = HIGHLIGHT_COLOR
        fc.StopIfTrue = False
    End If

    wb.Saved = wasSaved
End Sub

Private Sub RemoveCrosshairOnSheet(ws As Worksheet)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ws.Parent
    Dim wasSaved As Boolean
    wasSaved = wb.Saved

    Dim fcs As FormatConditions
    Set fcs = ws.Cells.FormatConditions
    Dim i As Long, f As String
    For i = fcs.Count To 1 Step -1
        f = ""
        f = fcs(i).Formula1
        If IsCrosshairFormula(f) Then fcs(i).Delete
    Next i

    wb.Saved = wasSaved
End Sub

Private Function IsCrosshairFormula(f As String) As Boolean
    If InStr(f, CCH_ROW_MARKER) > 0 _
       Or InStr(f, CCH_COL_MARKER) > 0 _
       Or InStr(f, CCH_MASK_MARKER) > 0 Then
        IsCrosshairFormula = True
        Exit Function
    End If
    ' Legacy formulas left behind by earlier broken builds.
    If Left$(f, 7) = "=ROW()=" Then
        If AllDigits(Mid$(f, 8)) Then IsCrosshairFormula = True : Exit Function
    End If
    If Left$(f, 10) = "=COLUMN()=" Then
        If AllDigits(Mid$(f, 11)) Then IsCrosshairFormula = True : Exit Function
    End If
End Function

Private Function AllDigits(s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    AllDigits = True
End Function
