Attribute VB_Name = "GeneralWorkbookXLModule"
'   Excel VBA
'   General Use Module
'   GeneralWorkbookXLModule
'   Leonard Sponza
'   Last Modified 15/09/2021 13:20
'   Date Time Version 00

Option Explicit

Public Sub SET_ZOOM_ALL_SHEETS(zoom_per As Integer)
    Dim ws As Worksheet

    For Each ws In Worksheets
        If (ws.Visible = xlSheetVisible) Then
            ws.Select
            ActiveWindow.Zoom = zoom_per ' change as per your requirements
        Else
            ' skip hidden pages
        End If
    Next ws
End Sub

Public Sub SET_ZOOM_ONE_HUNDRED_PAGE_BREAK_VIEW_ALL_SHEETS()
    Dim ws As Worksheet

    For Each ws In Worksheets
        If (ws.Visible = xlSheetVisible) Then
            ws.Select
            ActiveWindow.View = xlPageBreakPreview
            ActiveWindow.Zoom = 100
        Else
            ' skip hidden pages
        End If
    Next ws

'    With Worksheets
'        .Select
'        ActiveWindow.View = xlPageBreakPreview
'        ActiveWindow.Zoom = 100
'    End With

End Sub

Public Sub SET_ZOOM_ONE_HUNDRED_PAGE_LAYOUT_VIEW_ALL_SHEETS()
    Dim ws As Worksheet

    For Each ws In Worksheets
        If (ws.Visible = xlSheetVisible) Then
            ws.Select
            ActiveWindow.View = xlPageLayoutView
            ActiveWindow.Zoom = 100
        Else
            ' skip hidden pages
        End If
    Next ws
End Sub

Public Sub SET_ZOOM_ONE_HUNDRED_NORMAL_VIEW_ALL_SHEETS()
    Dim ws As Worksheet

    For Each ws In Worksheets
        If (ws.Visible = xlSheetVisible) Then
            ws.Select
            ActiveWindow.View = xlNormalView
            ActiveWindow.Zoom = 100
        Else
            ' skip hidden pages
        End If
    Next ws
End Sub

Private Sub Workbook_Open()

    '   Put code in here
    'MsgBox "test"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '   Put code in here
    '   Application.Calculation = xlCalculationManual
    ThisWorkbook.Close False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '   Put code in here
    '   Application.Calculation = xlCalculationManual
End Sub