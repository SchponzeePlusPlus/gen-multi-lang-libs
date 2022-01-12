Attribute VB_Name = "GhostDropDownFixXLModule"
Option Explicit
'   https://answers.microsoft.com/en-us/office/forum/office_365hp-excel/drop-down-menu-stuck-will-not-go-away/c588c5a0-f263-41c3-828d-c3335b3e0c90?page=2

Sub RemoveDropDownControls()
    Dim wsLog As Worksheet
    Dim strPass As String
    Dim shp As Shape
    Dim ws As Worksheet
    Dim rngDestin As Range
    Dim bolProtect As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strShtName As String
   
    strPass = "21Shirley"       'Edit "sonic" with correct password if necessary
   
    Set wsLog = Sheets.Add(After:=Sheets(Sheets.Count))
    With wsLog
        .Cells(1, "A").Value = "Worksheet"
        .Cells(1, "B").Value = "DropDown Name"
        .Cells(1, "C").Value = "Left"
        .Cells(1, "D").Value = "Top"
        .Cells(1, "E").Value = "Width"
        .Cells(1, "F").Value = "Height"
        .Cells(1, "G").Value = "Address"
        .Range(.Cells(1, "A"), .Cells(1, "G")).Font.Bold = True
    End With
   
    For Each ws In Worksheets
        strShtName = UCase(ws.Name)     'Case test is case sensitive so converts sheet name to upper case
        Select Case strShtName
            'Add additional sheets if necessary. (All in uppercase)
            '   Case "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"
            Case "Pricing Summary", "Sheet1"
                ws.Visible = xlSheetVisible
                ws.Select
                If ws.ProtectContents Then
                    ws.Unprotect Password:=strPass
                    bolProtect = True     'Save state of protection (True if Protected)
                Else
                    bolProtect = False    'Save state of protection (False if UnProtected)
                End If
               
                For Each shp In ws.Shapes
                    If Left(shp.Name, 9) = "Drop Down" Then
                        With wsLog
                            Set rngDestin = .Cells(.Rows.Count, "A").End(xlUp).Offset(1, 0)
                        End With
                        With shp
                            .Visible = msoTrue
                            rngDestin.Value = ws.Name
                            rngDestin.Offset(0, 1) = .Name
                            rngDestin.Offset(0, 2) = .Left
                            rngDestin.Offset(0, 3) = .Top
                            rngDestin.Offset(0, 4) = .Width
                            rngDestin.Offset(0, 5) = .Height
                            lngRow = rngRow(ws, .Top).Row
                            lngCol = rngCol(ws, .Left).Column
                            rngDestin.Offset(0, 6).Value = ws.Cells(lngRow, lngCol).Address
                            .Delete     'Can comment out if you want to see the list of DropwDowns first
                        End With
                    End If
                Next shp
                If bolProtect Then
                   ws.Protect Password:=strPass
                End If
        End Select
    Next ws
    wsLog.Columns("A:F").AutoFit
    wsLog.Select
    MsgBox "Finished"
End Sub

Function rngRow(wSht As Worksheet, dblTop As Double) As Range
    For Each rngRow In wSht.Rows
        If rngRow.Top >= dblTop Then Exit Function
    Next rngRow
End Function

Function rngCol(wSht As Worksheet, dblLeft As Double) As Range
    For Each rngCol In wSht.Columns
        If rngCol.Left >= dblLeft Then Exit Function
    Next rngCol
End Function
