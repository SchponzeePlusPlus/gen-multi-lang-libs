Attribute VB_Name = "GeneralXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   GeneralXLModule
'   Leonard Sponza
'   Last Modified 18/08/2021 17:10
'   Date Time Version 00

Option Explicit

Public Const CELL_VAL_CHAR_LEN_LIM As Long = 32767
Public Const CELL_VAL_CHAR_LEN_SAFE_LIM As Long = CELL_VAL_CHAR_LEN_LIM - 50

Public Function CHECK_SINGLE_CELL_IN_RANGE_V000(input_range As Range) As Boolean
    Dim result as Boolean

    If (input_range.Cells.Count = 1) Then
        result = True
    ElseIf (input_range.Cells.Count < 1) Then
        result = False
    ElseIf (input_range.Cells.Count > 1) Then
        result = False
    Else
        result = False
    End If

    CHECK_SINGLE_CELL_IN_RANGE_V000 = result
End Function

Public Function CORRECT_SINGLE_CELL_IN_RANGE_V000(ByVal input_range As Range) As Range
    Dim result As Range
    Dim single_cell_chk As Boolean

    single_cell_chk = CHECK_SINGLE_CELL_IN_RANGE_V000(input_range)

    If (single_cell_chk = True) Then
        Set result = input_range
    Else
        '   Assigns top left cell from input_range
        result = Cells(input_range.Row, input_range.Column)
    End If
    
    Set CORRECT_SINGLE_CELL_IN_RANGE_V000 = result
End Function

Public Sub Array2Range(My2DArray As Variant, aWS As Worksheet)
' Ref : https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
' Ref : http://www.ozgrid.com/VBA/sort-array.htm
' Ref : https://www.mrexcel.com/forum/excel-questions/14194-vba-arrays-examples-please-how-read-range-th.html
' Ref : https://bettersolutions.com/excel/cells-ranges/vba-working-with-arrays.htm
' Usage : Array2Range MyArray, aWS

  Dim i As Long


  For i = 1 To UBound(My2DArray) - LBound(My2DArray) + 1
      aWS.Cells(1, i).Resize(UBound(My2DArray(i))).Value = Application.Transpose(My2DArray(i))
  Next i

End Sub

Public Sub CONVERT_ONE_DIM_ARRAY_TO_RANGE(Data As Variant, Cl As Range)
    Dim local_arr() As Variant
    Dim i As Long
    ' Cl.Resize(UBound(Data, 1), UBound(Data, 2)) = Data
    
    Debug.Print "(UBound(Data, 1) - LBound(Data, 1) + 1):"
    Debug.Print (UBound(Data, 1) - LBound(Data, 1) + 1)
    
    ' one dimension - vertical only at the moment
    ReDim local_arr(LBound(Data, 1) To UBound(Data, 1), 0 To 0)
    
    For i = LBound(local_arr, 1) To UBound(local_arr, 1)
        local_arr(i, 0) = Data(i)
    Next
    
    Cl.Resize((UBound(local_arr, 1) - LBound(local_arr, 1) + 1), 1) = local_arr
End Sub

Sub PrintArray(Data As Variant, Cl As Range)
    '   https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook 
    ' Cl.Resize(UBound(Data, 1), UBound(Data, 2)) = Data
    
    Debug.Print "(UBound(Data, 1) - LBound(Data, 1) + 1):"
    Debug.Print (UBound(Data, 1) - LBound(Data, 1) + 1)
    Debug.Print "(UBound(Data, 2) - LBound(Data, 2) + 1):"
    Debug.Print (UBound(Data, 2) - LBound(Data, 2) + 1)
    
    Cl.Resize((UBound(Data, 1) - LBound(Data, 1) + 1), (UBound(Data, 2) - LBound(Data, 2) + 1)) = Data
End Sub

Public Sub CONVERT_ARRAY_TO_RANGE_VIA_PROMPT(Data() As Variant)
    Dim Cl As Range
    
    Set Cl = Application.InputBox(prompt:="Enter Destination Range for Array: ", Type:=8)
    
    Debug.Print "(UBound(Data, 1) - LBound(Data, 1) + 1):"
    Debug.Print (UBound(Data, 1) - LBound(Data, 1) + 1)
    Debug.Print "(UBound(Data, 2) - LBound(Data, 2) + 1):"
    Debug.Print (UBound(Data, 2) - LBound(Data, 2) + 1)
    
    Cl.Resize((UBound(Data, 1) - LBound(Data, 1) + 1), (UBound(Data, 2) - LBound(Data, 2) + 1)) = Data
End Sub

Public Function CONVERT_JOIN_ARR_STRING_TO_CELL_VAL_V000(input_string As String) As Variant
    Dim result As Variant

    If (Len(input_string) <= CELL_VAL_CHAR_LEN_SAFE_LIM) Then
        result = input_string
    Else
        '   result = "Error! String too long for cell"
        result = CVErr(xlErrValue)
    End If

    CONVERT_JOIN_ARR_STRING_TO_CELL_VAL_V000 = result
End Function

Public Function CONVERT_JOIN_ARR_CELL_VAL_TO_STRING_V000(input_cell_val As Variant) As String
    Dim result As String

    Dim process_cell_val As Variant

    process_cell_val = Access.Nz(input_cell_val, "(Null)")

    If (Len(CStr(process_cell_val)) <= STRING_LENGTH_SAFE_LIM) Then
        result = CStr(process_cell_val)
    Else
        result = "Error! Cell too long for String Data Type."
    End If

    CONVERT_JOIN_ARR_CELL_VAL_TO_STRING_V000 = result
End Function

' Public Sub COPY_FILTERED_TABLE_VALUES()
'     Dim input_range As Range
'     Dim output_range As Range
    '   meant to copy the filtered cells of a column to another column without pasting in between filtered results
    ' working alternative - sort by value instead
' End Sub

'   Taken from https://www.ozgrid.com/VBA/return-sheet-name.htm
Function SheetName(rCell As Range, Optional UseAsRef As Boolean) As String

    Application.Volatile

        If UseAsRef = True Then

            SheetName = "'" & rCell.Parent.Name & "'!"

        Else

            SheetName = rCell.Parent.Name

        End If

End Function

'   Based from https://www.ozgrid.com/VBA/return-sheet-name.htm
Function RETURN_SHEET_NAME_FROM_RANGE_V000(rCell As Range, Optional UseAsRef As Boolean = False, Optional ByVal volatile_enable As Boolean = False) As String
    Dim result As String
    
    If (volatile_enable = True) Then
        Application.Volatile
    Else

    End If

    If (UseAsRef = True) Then

        result = "'" & rCell.Parent.Name & "'!"

    ElseIf (UseAsRef = False) Then

        result = rCell.Parent.Name

    Else
        result = "(Error)"
    End If

    RETURN_SHEET_NAME_FROM_RANGE_V000 = result
End Function

Public Function RETURN_SHEET_NAME() As String
'    RETURN_SHEET_NAME = ActiveSheet.Name
'    RETURN_SHEET_NAME = Application.WorksheetFunction.Mid(Application.WorksheetFunction.Cell("filename", A1), Application.WorksheetFunction.Find("]", Application.WorksheetFunction.Cell("filename", A1)) + 1, 255)
    ' uncomment the below line to make it Volatile
    'Application.Volatile
    RETURN_SHEET_NAME = Application.Caller.Worksheet.Name
End Function

Public Function RETURN_SHEET_NAME_VIA_VE_V000(volatile_enable As Boolean) As String
    Dim result As String
'    RETURN_SHEET_NAME = ActiveSheet.Name
'    RETURN_SHEET_NAME = Application.WorksheetFunction.Mid(Application.WorksheetFunction.Cell("filename", A1), Application.WorksheetFunction.Find("]", Application.WorksheetFunction.Cell("filename", A1)) + 1, 255)
   ' uncomment the below line to make it Volatile
    If (volatile_enable = True) Then
        Application.Volatile
    Else

    End If
    result = Application.Caller.Worksheet.Name
    RETURN_SHEET_NAME_VIA_VE_V000 = result
End Function

Public Sub CALC_WORKSHEET(in_wk_sheet As Worksheet)
    in_wk_sheet.Calculate
End Sub

Public Sub CALC_WORKSHEETS(in_wk_sheets As Worksheets)
    in_wk_sheets.Calculate
End Sub

Public Sub CALC_FULL_WORKSHEETS(in_wk_sheets As Worksheets)
    in_wk_sheets.CalculateFull
End Sub

'   problematic in SMSP 07/12/2021
Public Sub CALC_EVERYTHING()
    Application.Calculate
End Sub

Public Sub CALC_FULL_EVERYTHING()
    Application.CalculateFull
End Sub

Public Sub CALC_FULL_EVERYTHING_NO_SU_NO_EV()
    MsgBox "MS Excel will become unresponsive until calculations are complete. Press OK to continue then please wait for the next window to appear!"
    '   https://analystcave.com/excel-improve-vba-performance/
    'Turn off automatic ScreenUpdating (on demand). Run ApplicationDoEvents procedure to manually update the screen
    Application.ScreenUpdating = False
    'Disables Excel events during the runtime of the VBA Macro
    Application.EnableEvents = False
    Application.CalculateFull
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Calculations are complete! Press OK to continue."
End Sub

Public Sub CALC_FULL_REBUILD_EVERYTHING()
    Application.CalculateFullRebuild
End Sub

Public Sub CALC_FULL_REBUILD_EVERYTHING_NO_SU_NO_EV()
    MsgBox "MS Excel will become unresponsive until calculations are complete. Press OK to continue then please wait for the next window to appear!"
    '   https://analystcave.com/excel-improve-vba-performance/
    'Turn off automatic ScreenUpdating (on demand). Run ApplicationDoEvents procedure to manually update the screen
    Application.ScreenUpdating = False
    'Disables Excel events during the runtime of the VBA Macro
    Application.EnableEvents = False
    Application.CalculateFullRebuild
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Calculations are complete! Press OK to continue."
End Sub

Public Sub REFRESH_ALL_THIS_WORKBOOK()
    ThisWorkbook.RefreshAll
End Sub

Public Sub CALC_WORKSHEET_VIA_RANGE(in_sheet_rng_add As Range, Optional ByVal volatile_enable As Boolean = False)
    Dim wk_sheet_name_str As String

    wk_sheet_name_str = RETURN_SHEET_NAME_FROM_RANGE_V000(in_sheet_rng_add, False, volatile_enable)
    Worksheets(wk_sheet_name_str).Calculate
    
End Sub

'Public Sub SET_PAGE_SETUP_NARROW_MARGINS_A4()
'
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = ""
'        .PrintTitleColumns = ""
'    End With
'    Application.PrintCommunication = True
'    ' ActiveSheet.PageSetup.PrintArea =
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
''        .LeftHeader = ""
''        .CenterHeader = ""
''        .RightHeader = ""
''        .LeftFooter = ""
''        .CenterFooter = ""
''        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        ' .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = False
'        .CenterVertically = False
'        '.Orientation = xlPortrait
'        .Draft = False
'        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        '.BlackAndWhite = False
'        '.Zoom = False
'        .FitToPagesWide = 1
'        '.FitToPagesTall = False
'        .PrintErrors = xlPrintErrorsDisplayed
'        '.OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = False
'        .EvenPage.LeftHeader.Text = ""
'        .EvenPage.CenterHeader.Text = ""
'        .EvenPage.RightHeader.Text = ""
'        .EvenPage.LeftFooter.Text = ""
'        .EvenPage.CenterFooter.Text = ""
'        .EvenPage.RightFooter.Text = ""
'        .FirstPage.LeftHeader.Text = ""
'        .FirstPage.CenterHeader.Text = ""
'        .FirstPage.RightHeader.Text = ""
'        .FirstPage.LeftFooter.Text = ""
'        .FirstPage.CenterFooter.Text = ""
'        .FirstPage.RightFooter.Text = ""
'    End With
'    Application.PrintCommunication = True
'End Sub