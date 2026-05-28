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
' active column in every open workbook using Conditional Formatting rules.
' A third "mask" rule is placed on the active cell with StopIfTrue=True so
' the row/column highlight does not paint over the selected cell itself,
' keeping the selection indicator clearly visible.
'
' Each rule's Formula1 carries a marker substring so we can identify and
' remove our rules reliably without holding (potentially stale)
' FormatCondition object references.

Private Const HIGHLIGHT_COLOR  As Long   = 13561798   ' RGB(198, 239, 206) mint green
Private Const CCH_ROW_MARKER   As String = "__CCH_ROW__"
Private Const CCH_COL_MARKER   As String = "__CCH_COL__"
Private Const CCH_MASK_MARKER  As String = "__CCH_MASK__"

Public WithEvents App As Excel.Application

Private mLastSheet As Worksheet

Private Sub Class_Initialize()
    Set App = Excel.Application
End Sub

' --- Application event handlers --------------------------------------------

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Sh Is Nothing Or Target Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Not mLastSheet Is Nothing Then
        If Not (mLastSheet Is Sh) Then RemoveCrosshairOnSheet mLastSheet
    End If

    RemoveCrosshairOnSheet Sh
    ApplyCrosshair Sh, Target.Row, Target.Column
    Set mLastSheet = Sh

Done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Fail:
    Resume Done
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Scrub leftover crosshair CFs from the active sheet only — that's the
    ' only sheet that could be showing one. Other sheets get cleaned the
    ' first time the user navigates to them (SheetActivate). Iterating
    ' every sheet on every open is too slow for multi-sheet workbooks.
    On Error Resume Next
    Application.EnableEvents = False
    If TypeName(Wb.ActiveSheet) = "Worksheet" Then RemoveCrosshairOnSheet Wb.ActiveSheet
    Application.EnableEvents = True
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
    On Error Resume Next
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then Exit Sub
    If Application.ActiveCell Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    If Not mLastSheet Is Nothing Then
        If Not (mLastSheet Is ws) Then RemoveCrosshairOnSheet mLastSheet
    End If
    RemoveCrosshairOnSheet ws
    ApplyCrosshair ws, Application.ActiveCell.Row, Application.ActiveCell.Column
    Set mLastSheet = ws

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    Application.EnableEvents = False
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        RemoveCrosshairOnSheet ws
    Next ws
    Application.EnableEvents = True
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    On Error Resume Next
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ApplyCrosshair ActiveSheet, ActiveCell.Row, ActiveCell.Column
    Set mLastSheet = ActiveSheet
    Wb.Saved = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        RemoveCrosshairOnSheet ws
    Next ws
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
