Attribute VB_Name = "CustomXLLibrary"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   CustomXLLibrary
'   Leonard Sponza
'   Last Modified 15/09/2021 12:35
'   Date Time Version 00

Option Explicit

Public Const DEFAULT_SHEET_ROW_ADDRESS As String = "1"
Public Const DEFAULT_SHEET_COL_ADDRESS As String = "A"
Public Const DEFAULT_SHEET_CELL_ADDRESS As String = DEFAULT_SHEET_COL_ADDRESS & DEFAULT_SHEET_COL_ADDRESS

' CUSTOM GENERAL EXCEL FUNCTIONS

Public Function WRAP_CELL_VALUE_TO_VARIANT(input_cell As Variant) As Variant
    WRAP_CELL_VALUE_TO_VARIANT = input_cell.Value2
End Function

Public Function WRAP_RANGE_CELL_VALUE_TO_VARIANT(input_cell As Range) As Variant
    Dim result As Variant
    If ((input_cell.Rows.Count = 1) And (input_cell.Columns.Count = 1)) Then
        result = input_cell.Value2
    Else
        result = "(Error)"
    End If
    WRAP_RANGE_CELL_VALUE_TO_VARIANT = result
End Function

Public Function PRINT_CELL_VALUE_TYPENAME(input_cell As Variant) As String
    Dim input_var As Variant
    
    Dim test_var As Variant
    test_var = "String"
    
    Dim result As String
    
'    PRINT_CELL_VALUE_TYPENAME = TypeName(input_cell.Value2)
'    PRINT_CELL_VALUE_TYPENAME = TypeName(test_var)
'    result = TypeName(test_var)
    result = TypeName(input_cell.Value2)
    PRINT_CELL_VALUE_TYPENAME = result
End Function

Public Function CAST_CELL_VALUE_TO_DOUBLE(input_cell As Variant) As Double
    Dim val_var As Variant
    
    val_var = WRAP_CELL_VALUE_TO_VARIANT(input_cell)
    
    CAST_CELL_VALUE_TO_DOUBLE = CAST_VARIANT_TO_DOUBLE(val_var)
End Function

Public Function CAST_CELL_VALUE_TO_DOUBLE_NO_ENUM_CODE_VIA_IC_V000(input_cell As Variant) As Double
    Dim result As Double
    Dim val_var As Variant
    
    val_var = WRAP_CELL_VALUE_TO_VARIANT(input_cell)
    
    result = CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_V000(val_var)

    CAST_CELL_VALUE_TO_DOUBLE_NO_ENUM_CODE_VIA_IC_V000 = result
End Function

Public Function CAST_CELL_VALUE_TO_STRING(input_cell As Variant) As String
    Dim val_var As Variant
    
    val_var = WRAP_CELL_VALUE_TO_VARIANT(input_cell)
    
    CAST_CELL_VALUE_TO_STRING = CAST_VARIANT_TO_STRING(val_var)
End Function

Private Function IsValidColorIndex(ColorIndex As Long) As Boolean
    Select Case ColorIndex
        Case 1 To 56
            IsValidColorIndex = True
        Case xlColorIndexAutomatic, xlColorIndexNone
            IsValidColorIndex = True
        Case Else
            IsValidColorIndex = False
    End Select
End Function

Public Function ColorIndexOfOneCell(Cell As Range, OfText As Boolean, _
    DefaultColorIndex As Long) As Long
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ColorIndexOfOneCell
    ' This returns the ColorIndex of the cell referenced by Cell.
    ' If Cell refers to more than one cell, only Cell(1,1) is
    ' tested. If OfText True, the ColorIndex of the Font property is
    ' returned. If OfText is False, the ColorIndex of the Interior
    ' property is returned. If DefaultColorIndex is >= 0, this
    ' value is returned if the ColorIndex is either xlColorIndexNone
    ' or xlColorIndexAutomatic.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim CI As Long
    
    Application.Volatile True
    If OfText = True Then
        CI = Cell(1, 1).Font.ColorIndex
    Else
        CI = Cell(1, 1).Interior.ColorIndex
    End If
    If CI < 0 Then
        If IsValidColorIndex(ColorIndex:=DefaultColorIndex) = True Then
            CI = DefaultColorIndex
        Else
            CI = -1
        End If
    End If
    
    ColorIndexOfOneCell = CI
    
End Function

Public Sub INSERT_STATIC_HYPERLINK_TO_CELL(target_hl_cell As Range, hyperlink_address As String, display_text As String)
    ' If ((target_hl_cell.Cells = 0) Or (target_hl_cell.Cells = 1)) Then
    If (target_hl_cell.Cells.Count = 1) Then
        target_hl_cell.Hyperlinks.Add target_hl_cell, hyperlink_address, , , display_text
    Else
        MsgBox "Error! Too Many cells in argument!"
    End If
End Sub

Public Function CONVERT_RANGE_TO_ONE_DIM_ARRAY(inputRange As Range, element_arrangement As String) As Variant()
   
    ' Dim out() As Double
    ' ReDim out(inputRange.Columns.Count - 1)
    Dim out() As Variant
    ReDim out(0 To ((inputRange.Rows.Count * inputRange.Columns.Count) - 1))

    Dim i As Long, j As Long, elem_cntr As Long
    elem_cntr = 0
    
    Select Case element_arrangement
        ' goes rightwards before downwards
        Case "RIGHTWARDS_ALONG_ROWS_FIRST"
            For i = 1 To inputRange.Rows.Count
                For j = 1 To inputRange.Columns.Count
                    out(elem_cntr) = inputRange(i, j)
                    elem_cntr = elem_cntr + 1
                Next
            Next
        Case "DOWNWARDS_ALONG_COLS_FIRST"
            For j = 1 To inputRange.Columns.Count
                For i = 1 To inputRange.Rows.Count
                    out(elem_cntr) = inputRange(i, j)
                    elem_cntr = elem_cntr + 1
                Next
            Next
        Case Else
            For i = 1 To inputRange.Rows.Count
                For j = 1 To inputRange.Columns.Count
                    out(elem_cntr) = ERROR_DOUBLE_VAL
                    elem_cntr = elem_cntr + 1
                Next
            Next
    End Select

    CONVERT_RANGE_TO_ONE_DIM_ARRAY = out
End Function

Public Function CONVERT_RANGE_TO_TWO_DIM_ARRAY(inputRange As Range, element_arrangement As String) As Variant()
   
    ' Dim out() As Double
    ' ReDim out(inputRange.Columns.Count - 1)
    Dim out() As Variant
    Dim output_var As Variant
    Dim i As Long, j As Long
    Dim input_row_cnt As Long, input_col_cnt As Long
    
    input_row_cnt = inputRange.Rows.Count
    input_col_cnt = inputRange.Columns.Count
    
    Select Case element_arrangement
        Case "DEFAULT"
            output_var = inputRange.Value2
            out = output_var
        Case "DIM_ONE_EACH_ROW"
            ' same operation as default just starts elements at 0
            ReDim out((inputRange.Rows.Count - 1), (inputRange.Columns.Count - 1))
        
            For i = 0 To inputRange.Rows.Count - 1
                For j = 0 To inputRange.Columns.Count - 1
                    out(i, j) = inputRange(i + 1, j + 1)
                Next
            Next
        Case "DIM_ONE_EACH_COL"
            'ReDim out((inputRange.Columns.Count - 1), (inputRange.Rows.Count - 1))
            ReDim out(0 To (input_col_cnt - 1), 0 To (input_row_cnt - 1))
        
            For j = 0 To input_col_cnt - 1
                For i = 0 To input_row_cnt - 1
                    out(j, i) = inputRange(i + 1, j + 1)
                Next
            Next
        Case Else
            ReDim out((inputRange.Rows.Count - 1), (inputRange.Columns.Count - 1))
        
            For i = 0 To inputRange.Rows.Count - 1
                For j = 0 To inputRange.Columns.Count - 1
                    out(i, j) = NULL_DOUBLE_VAL
                Next
            Next
    End Select

    CONVERT_RANGE_TO_TWO_DIM_ARRAY = out
End Function

Public Function CONVERT_ONE_DIM_VARIANT_ARR_TO_INTEGER_ARRAY(input_variant_array() As Variant) As Integer()
    Dim result() As Integer
    
    Dim input_arr_low_bound As Long, input_arr_up_bound As Long
    Dim input_arr_length As Long
    
    input_arr_low_bound = LBound(input_variant_array, 1)
    input_arr_up_bound = UBound(input_variant_array, 1)
    
    input_arr_length = input_arr_up_bound - input_arr_low_bound + 1
    
    ReDim result(input_arr_low_bound To input_arr_up_bound)
    
    Dim i As Long
    For i = input_arr_low_bound To input_arr_up_bound
        result(i) = CAST_VARIANT_TO_INTEGER(input_variant_array(i))
    Next
    
    CONVERT_ONE_DIM_VARIANT_ARR_TO_INTEGER_ARRAY = result
    
End Function

Public Function CONVERT_ONE_DIM_VARIANT_ARR_TO_LONG_ARRAY(input_variant_array() As Variant) As Long()
    Dim result() As Long
    
    Dim input_arr_low_bound As Long, input_arr_up_bound As Long
    Dim input_arr_length As Long
    
    input_arr_low_bound = LBound(input_variant_array, 1)
    input_arr_up_bound = UBound(input_variant_array, 1)
    
    input_arr_length = input_arr_up_bound - input_arr_low_bound + 1
    
    ReDim result(input_arr_low_bound To input_arr_up_bound)
    
    Dim i As Long
    For i = input_arr_low_bound To input_arr_up_bound
        result(i) = CAST_VARIANT_TO_LONG(input_variant_array(i))
    Next
    
    CONVERT_ONE_DIM_VARIANT_ARR_TO_LONG_ARRAY = result
    
End Function

Public Function CONVERT_ONE_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY(input_variant_array() As Variant) As Double()
    Dim result() As Double
    
    Dim input_arr_low_bound As Long, input_arr_up_bound As Long
    Dim input_arr_length As Long
    
    input_arr_low_bound = LBound(input_variant_array, 1)
    input_arr_up_bound = UBound(input_variant_array, 1)
    
    input_arr_length = input_arr_up_bound - input_arr_low_bound + 1
    
    ReDim result(input_arr_low_bound To input_arr_up_bound)
    
    Dim i As Long
    For i = input_arr_low_bound To input_arr_up_bound
        result(i) = CAST_VARIANT_TO_DOUBLE(input_variant_array(i))
    Next
    
    CONVERT_ONE_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY = result
    
End Function

Public Function CONVERT_ONE_DIM_VARIANT_ARR_TO_STRING_ARRAY(input_variant_array() As Variant) As String()
    Dim result() As String
    
    Dim input_arr_low_bound As Long, input_arr_up_bound As Long
    Dim input_arr_length As Long
    
    input_arr_low_bound = LBound(input_variant_array, 1)
    input_arr_up_bound = UBound(input_variant_array, 1)
    
    input_arr_length = input_arr_up_bound - input_arr_low_bound + 1
    
    ReDim result(input_arr_low_bound To input_arr_up_bound)
    
    Dim i As Long
    For i = input_arr_low_bound To input_arr_up_bound
        result(i) = CStr(input_variant_array(i))
    Next
    
    CONVERT_ONE_DIM_VARIANT_ARR_TO_STRING_ARRAY = result
    
End Function

Public Function CONVERT_TWO_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY(input_variant_array() As Variant) As Double()
    Dim result() As Double
    
    Dim input_arr_dim_one_low_bound As Long, input_arr_dim_one_up_bound As Long, input_arr_dim_two_low_bound As Long, input_arr_dim_two_up_bound As Long
    Dim input_arr_dim_one_length As Long, input_arr_dim_two_length As Long
    
    input_arr_dim_one_low_bound = LBound(input_variant_array, 1)
    input_arr_dim_one_up_bound = UBound(input_variant_array, 1)
    
    input_arr_dim_two_low_bound = LBound(input_variant_array, 2)
    input_arr_dim_two_up_bound = UBound(input_variant_array, 2)
    
    input_arr_dim_one_length = input_arr_dim_one_up_bound - input_arr_dim_one_low_bound + 1
    input_arr_dim_two_length = input_arr_dim_two_up_bound - input_arr_dim_two_low_bound + 1
    
    ReDim result(input_arr_dim_one_low_bound To input_arr_dim_one_up_bound, input_arr_dim_two_low_bound To input_arr_dim_two_up_bound)
    
    Dim i As Long, j As Long
    For i = input_arr_dim_one_low_bound To input_arr_dim_one_up_bound
        For j = input_arr_dim_two_low_bound To input_arr_dim_two_up_bound
            result(i, j) = CAST_VARIANT_TO_DOUBLE(input_variant_array(i, j))
        Next
    Next
    
    CONVERT_TWO_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY = result
    
End Function

Public Function CONVERT_RANGE_TO_ONE_DIM_INTEGER_ARRAY(input_range As Range, element_arrangement As String) As Integer()
    Dim local_variant_arr() As Variant
    
    local_variant_arr = CONVERT_RANGE_TO_ONE_DIM_ARRAY(input_range, element_arrangement)
    
    CONVERT_RANGE_TO_ONE_DIM_INTEGER_ARRAY = CONVERT_ONE_DIM_VARIANT_ARR_TO_INTEGER_ARRAY(local_variant_arr)
    
End Function

Public Function CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_range As Range, element_arrangement As String) As Long()
    Dim local_variant_arr() As Variant
    
    local_variant_arr = CONVERT_RANGE_TO_ONE_DIM_ARRAY(input_range, element_arrangement)
    
    CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY = CONVERT_ONE_DIM_VARIANT_ARR_TO_LONG_ARRAY(local_variant_arr)
    
End Function

Public Function CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_range As Range, element_arrangement As String) As Double()
    Dim local_variant_arr() As Variant
    
    local_variant_arr = CONVERT_RANGE_TO_ONE_DIM_ARRAY(input_range, element_arrangement)
    
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY = CONVERT_ONE_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY(local_variant_arr)
    
End Function

Public Function CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_range As Range, element_arrangement As String) As String()
    Dim local_variant_arr() As Variant
    
    local_variant_arr = CONVERT_RANGE_TO_ONE_DIM_ARRAY(input_range, element_arrangement)
    
    CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY = CONVERT_ONE_DIM_VARIANT_ARR_TO_STRING_ARRAY(local_variant_arr)
    
End Function

Public Function CONVERT_RANGE_TO_TWO_DIM_DOUBLE_ARRAY(input_range As Range, element_arrangement As String) As Double()
    Dim local_variant_arr() As Variant
    
    local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, element_arrangement)
    
    CONVERT_RANGE_TO_TWO_DIM_DOUBLE_ARRAY = CONVERT_ONE_DIM_VARIANT_ARR_TO_DOUBLE_ARRAY(local_variant_arr)
    
End Function

' Public Function CONVERT_RANGE_TO_TWO_DIM_STRING_ARRAY(input_range As Range, element_arrangement As String) As String()
'     Dim local_variant_arr() As Variant
    
'     local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, element_arrangement)
    
'     CONVERT_RANGE_TO_TWO_DIM_STRING_ARRAY = CONVERT_TWO_DIM_VARIANT_ARR_TO_STRING_ARRAY(local_variant_arr)
    
' End Function

Public Function SELECT_DIM_TWO_FROM_TWO_DIM_VAR_ARRAY_MAKE_ONE_DIM(input_arr() As Variant, col_num As Long) As Variant()
    Dim result() As Variant
    
    Dim i As Long
    
    Dim input_arr_dim_one_low_bound As Long, input_arr_dim_one_up_bound As Long
    Dim input_arr_dim_one_length As Long
    
    input_arr_dim_one_low_bound = LBound(input_arr, 1)
    input_arr_dim_one_up_bound = UBound(input_arr, 1)
    
    input_arr_dim_one_length = input_arr_dim_one_up_bound - input_arr_dim_one_low_bound + 1
    
    ReDim result(input_arr_dim_one_low_bound To input_arr_dim_one_up_bound)
    
    For i = input_arr_dim_one_low_bound To input_arr_dim_one_up_bound
        result(i) = input_arr(i, (col_num - 1))
    Next
    
    'result = input_variant_arr
    SELECT_DIM_TWO_FROM_TWO_DIM_VAR_ARRAY_MAKE_ONE_DIM = result
End Function

Public Function SELECT_DIM_TWO_FROM_TWO_DIM_VAR_ARRAY_KEEP_TWO_DIM(input_arr() As Variant, col_num As Long) As Variant()
    Dim result() As Variant
    
    Dim i As Long, j As Long
    
    Dim input_arr_dim_one_low_bound As Long, input_arr_dim_one_up_bound As Long, input_arr_dim_two_low_bound As Long, input_arr_dim_two_up_bound As Long
    Dim input_arr_dim_one_length As Long, input_arr_dim_two_length As Long
    
    input_arr_dim_one_low_bound = LBound(input_arr, 1)
    input_arr_dim_one_up_bound = UBound(input_arr, 1)
    
    input_arr_dim_two_low_bound = LBound(input_arr, 2)
    input_arr_dim_two_up_bound = UBound(input_arr, 2)
    
    input_arr_dim_one_length = input_arr_dim_one_up_bound - input_arr_dim_one_low_bound + 1
    input_arr_dim_two_length = input_arr_dim_two_up_bound - input_arr_dim_two_low_bound + 1
    
    ReDim result(input_arr_dim_one_low_bound To input_arr_dim_one_up_bound, 0 To 0)
    
    j = 0
    
    For i = input_arr_dim_one_low_bound To input_arr_dim_one_up_bound
        result(i, 0) = input_arr(i, col_num)
    Next
    
    'result = input_variant_arr
    SELECT_DIM_TWO_FROM_TWO_DIM_VAR_ARRAY_KEEP_TWO_DIM = result
End Function

Public Function SELECT_ROW_FROM_RANGE(input_range As Range, col_num As Long) As Variant()
    Dim result() As Variant
    
    'Dim sel_col_range As Range
    Dim local_variant_arr() As Variant
    Dim i As Long, j As Long
    
    Dim local_arr_dim_one_low_bound As Long, local_arr_dim_one_up_bound As Long, local_arr_dim_two_low_bound As Long, local_arr_dim_two_up_bound As Long
    Dim local_arr_dim_one_length As Long, local_arr_dim_two_length As Long
    
    'sel_col_range = input_range
    'sel_col_range = input_range.Columns(col_num).Item
    'Set sel_col_range = input_range.Columns.Item(col_num)
    'Set sel_col_range = input_range.Columns(col_num)
    
    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, element_arrangement)
    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(sel_col_range, "DIM_ONE_EACH_COL")
    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(sel_col_range, "DIM_ONE_EACH_ROW")
    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range.Columns(col_num), "DIM_ONE_EACH_ROW")
    local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, "DIM_ONE_EACH_ROW")
    
    local_arr_dim_one_low_bound = LBound(local_variant_arr, 1)
    local_arr_dim_one_up_bound = UBound(local_variant_arr, 1)
    
    local_arr_dim_two_low_bound = LBound(local_variant_arr, 2)
    local_arr_dim_two_up_bound = UBound(local_variant_arr, 2)
    
    local_arr_dim_one_length = local_arr_dim_one_up_bound - local_arr_dim_one_low_bound + 1
    local_arr_dim_two_length = local_arr_dim_two_up_bound - local_arr_dim_two_low_bound + 1
    
    ReDim result(0 To 0, local_arr_dim_two_low_bound To local_arr_dim_two_up_bound)
    
    i = 0
    
    For j = local_arr_dim_two_low_bound To local_arr_dim_two_up_bound
        result(0, j) = local_variant_arr((col_num - 1), j)
    Next
    
    'result = local_variant_arr
    SELECT_ROW_FROM_RANGE = result
End Function

Public Function SELECT_COLUMN_FROM_RANGE(input_range As Range, col_num As Long) As Variant()
    Dim result() As Variant
    
    'Dim sel_col_range As Range
    Dim local_variant_arr() As Variant
'    Dim i As Long, j As Long
'
'    Dim local_arr_dim_one_low_bound As Long, local_arr_dim_one_up_bound As Long, local_arr_dim_two_low_bound As Long, local_arr_dim_two_up_bound As Long
'    Dim local_arr_dim_one_length As Long, local_arr_dim_two_length As Long
'
'    'sel_col_range = input_range
'    'sel_col_range = input_range.Columns(col_num).Item
'    'Set sel_col_range = input_range.Columns.Item(col_num)
'    'Set sel_col_range = input_range.Columns(col_num)
'
'    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, element_arrangement)
'    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(sel_col_range, "DIM_ONE_EACH_COL")
'    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(sel_col_range, "DIM_ONE_EACH_ROW")
'    'local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range.Columns(col_num), "DIM_ONE_EACH_ROW")
'    local_variant_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY(input_range, "DIM_ONE_EACH_ROW")
'
'    local_arr_dim_one_low_bound = LBound(local_variant_arr, 1)
'    local_arr_dim_one_up_bound = UBound(local_variant_arr, 1)
'
'    local_arr_dim_two_low_bound = LBound(local_variant_arr, 2)
'    local_arr_dim_two_up_bound = UBound(local_variant_arr, 2)
'
'    local_arr_dim_one_length = local_arr_dim_one_up_bound - local_arr_dim_one_low_bound + 1
'    local_arr_dim_two_length = local_arr_dim_two_up_bound - local_arr_dim_two_low_bound + 1
'
'    ReDim result(local_arr_dim_one_low_bound To local_arr_dim_one_up_bound, 0 To 0)
'
'    j = 0
'
'    For i = local_arr_dim_one_low_bound To local_arr_dim_one_up_bound
'        result(i, 0) = local_variant_arr(i, (col_num - 1))
'    Next
'
'    'result = local_variant_arr
    
    result = SELECT_DIM_TWO_FROM_TWO_DIM_VAR_ARRAY_KEEP_TWO_DIM(local_variant_arr, col_num)
    
    SELECT_COLUMN_FROM_RANGE = result
End Function

Public Sub DELETE_RANGE_VALUES(target_range As Range)
    target_range.ClearContents
End Sub

Public Sub CLEAR_RANGE_FONT(target_range As Range)
    With target_range.Font
        .Background = xlBackgroundTransparent
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Bold = False
        .Size = 10
        .Name = "Arial"
    End With
End Sub

Public Sub SET_RANGE_FONT_COLOUR_AUTOMATIC(target_range As Range)
    target_range.Font.ColorIndex = xlAutomatic
End Sub

Public Sub SET_RANGE_FONT_NO_UNDERLINE(target_range As Range)
    target_range.Font.Underline = xlUnderlineStyleNone
End Sub

Public Sub SET_RANGE_FONT_BOLD(target_range As Range)
    target_range.Font.Bold = True
End Sub

Public Sub CLEAR_RANGE_ALIGNMENT(target_range As Range)
    With target_range
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub HORIZ_VERTI_CENTRE_RANGE_VALUES(target_range As Range)
    With target_range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
    End With
End Sub

Public Sub AUTOFIT_COLUMNS(target_range As Range)
    target_range.EntireColumn.AutoFit
End Sub

Public Sub CLEAR_RANGE_CELL_STYLE(target_range As Range)
    target_range.Style = "Normal"
End Sub

' Public Function QUICKSORT_ONE_DIM_VAR_ARRAY(in_var_arr As Variant, in_var_arr_low_bound As Long, in_var_arr_high_bound As Long)
'     Dim result As Variant
    
'     Dim pivot   As Variant
'     Dim tmpSwap As Variant
'     Dim tmpLow  As Long
'     Dim tmpHi   As Long

'     tmpLow = in_var_arr_low_bound
'     tmpHi = in_var_arr_high_bound

'     pivot = in_var_arr((in_var_arr_low_bound + in_var_arr_high_bound) \ 2)

'     While (tmpLow <= tmpHi)
'         While ((in_var_arr(tmpLow) < pivot) And (tmpLow < in_var_arr_high_bound))
'             tmpLow = tmpLow + 1
'         Wend

'         While (pivot < in_var_arr(tmpHi) And tmpHi > in_var_arr_low_bound)
'             tmpHi = tmpHi - 1
'         Wend

'         If (tmpLow <= tmpHi) Then
'             tmpSwap = in_var_arr(tmpLow)
'             in_var_arr(tmpLow) = in_var_arr(tmpHi)
'             in_var_arr(tmpHi) = tmpSwap
'             tmpLow = tmpLow + 1
'             tmpHi = tmpHi - 1
'         End If
'     Wend

'     If (in_var_arr_low_bound < tmpHi) Then QuickSort in_var_arr, in_var_arr_low_bound, tmpHi
'     End If
'     If (tmpLow < in_var_arr_high_bound) Then QuickSort in_var_arr, tmpLow, in_var_arr_high_bound
'     End If
' End Function

'Public Function ASSIGN_ACTIVE_INPUT(input_data_type_validity As Boolean, input_null_chk As Boolean) As Boolean
'    If ((input_data_type_validity = True) And (input_null_chk = False)) Then
'        ASSIGN_ACTIVE_INPUT = True
'    Else
'        ASSIGN_ACTIVE_INPUT = False
'    End If
'End Function

'Public Function ACTIVATE_RAW_STRING_INPUT(input_str As String, expected_data_type As String) As Boolean
'
'End Function

' Public Function CHECK_NULL_INTEGER()
' End Function

' Public Function RETURN_RESULTANT_VALID_INTEGER()
' End Function

Public Function RETURN_RESULTANT_VALID_VARIANT_FIELD(default_var As Variant, overwrite_var As Variant) As Variant
    If (Not (IsNull(overwrite_var))) Then
        RETURN_RESULTANT_VALID_VARIANT_FIELD = overwrite_var
'        RETURN_RESULTANT_VALID_INT_FIELD = 1
    Else
        RETURN_RESULTANT_VALID_VARIANT_FIELD = default_var
'        RETURN_RESULTANT_VALID_INT_FIELD = 2
    End If
End Function

Public Function RETURN_RESULTANT_VALID_RANGE_CELL_VALUE(default_cell As Range, overwrite_cell As Range) As Variant
    Dim result As Variant
    Dim default_var As Variant, overwrite_var As Variant
    Dim valid_overwrite_val As Boolean
    default_var = WRAP_RANGE_CELL_VALUE_TO_VARIANT(default_cell)
    overwrite_var = WRAP_RANGE_CELL_VALUE_TO_VARIANT(overwrite_cell)

    If ((IsEmpty(overwrite_var)) Or (IsNull(overwrite_var)) Or (overwrite_var = vbNullString)) Then
        valid_overwrite_val = False
    Else
        valid_overwrite_val = True
    End If

'     If ((Not (IsNull(overwrite_var))) Or (Not (IsEmpty(overwrite_var)))) Then
'     '   If (Not (IsEmpty(overwrite_var))) Then
'         result = overwrite_var
' '        RETURN_RESULTANT_VALID_INT_FIELD = 1
'     Else
'         result = default_var
' '        RETURN_RESULTANT_VALID_INT_FIELD = 2

'     If ((IsEmpty(overwrite_var)) Or (IsNull(overwrite_var))) Then
'         result = default_var
' '        RETURN_RESULTANT_VALID_INT_FIELD = 2
'     Else
'         result = overwrite_var
' '        RETURN_RESULTANT_VALID_INT_FIELD = 1

    If (valid_overwrite_val = True) Then
        result = overwrite_var
        '   RETURN_RESULTANT_VALID_INT_FIELD = 1
    Else
        result = default_var
        '   RETURN_RESULTANT_VALID_INT_FIELD = 2

    End If
    RETURN_RESULTANT_VALID_RANGE_CELL_VALUE = result
End Function

Public Function RETURN_RESULTANT_VARIANT_FIELD_VIA_OVERWRITE_BOOL(default_var As Variant, overwrite_var As Variant, overwrite_bool_chk As Boolean) As Variant
    If (overwrite_bool_chk = True) Then
        RETURN_RESULTANT_VARIANT_FIELD_VIA_OVERWRITE_BOOL = overwrite_var
    Else
        RETURN_RESULTANT_VARIANT_FIELD_VIA_OVERWRITE_BOOL = default_var
    End If
End Function

