Attribute VB_Name = "GeneralCalcXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   GeneralCalcXLModule
'   Leonard Sponza
'   Last Modified 16/09/2021 17:35
'   Date Time Version 00

Option Explicit

Public Function CUSTOM_VLOOKUP(lookup_value As Variant, table_array As Range, col_index_num As Variant, range_lookup As Variant) As Variant
	If Application.WorksheetFunction.VLookup(lookup_value, table_array, col_index_num, range_lookup) = "" Then
		CUSTOM_VLOOKUP = ""
	Else
		CUSTOM_VLOOKUP = Application.WorksheetFunction.VLookup(lookup_value, table_array, col_index_num, range_lookup)
	End If
End Function

' Public Function XLOOKUP_STRING_IO(lookup_str As String, lookup_str_arr() As String, return_array() As Variant, if_not_found As Variant, match_mode As Variant, search_mode As Variant) As String
'     Dim result As String
'     result = Application.WorksheetFunction.XLOOKUP(lookup_str, lookup_str_arr, return_array, if_not_found, match_mode, search_mode)
'     XLOOKUP_STRING_IO = result
' End Function
'
' Public Function XLOOKUP_CAST_LOOKUP_VAL_TO_STRING(lookup_value As Variant, lookup_array As Range, return_array As Range, if_not_found As Variant, match_mode As Variant, search_mode As Variant) As String
'     Dim result As String

'     Dim lookup_str As String
'     Dim lookup_str_arr() As String
'     Dim return_array() As Variant

'     lookup_str = CAST_VARIANT_TO_STRING(lookup_value)
'     lookup_str_arr = CONVERT_RANGE_TO_TWO_DIM_ARRAY 


'     result = XLOOKUP_STRING_IO(lookup_str, lookup_str_arr, return_array, if_not_found, match_mode, search_mode)
'     XLOOKUP_CAST_LOOKUP_VAL_TO_STRING = result
' End Function

Function CUSTOM_XLOOKUP_ONE(searchVal As Variant, searchArray As Range, returnArray As Variant, Optional notFound As Variant, Optional arg1 As Variant, Optional arg2 As Variant) As Variant 'v1.1
  'https://www.reddit.com/u/excelevator
  'https://old.reddit.com/r/excelevator
  'https://www.reddit.com/r/excel - for all your Spreadsheet questions!
	If IsMissing(arg1) Then arg1 = 0
	If IsMissing(arg2) Then arg2 = 0
	Dim rsult As Variant 'take the final result array
	Dim r2width As Integer: r2width = searchArray.Columns.Count
	Dim r3width As Integer: r3width = returnArray.Columns.Count
	Dim rtnHeaderColumn As Boolean: rtnHeaderColumn = r2width > 1
	If r2width > 1 And r2width <> r3width Then
	CUSTOM_XLOOKUP_ONE = CVErr(xlErrRef)
	Exit Function
	End If
	Dim srchVal As Variant: srchVal = searchVal 'THE SEARCH VALUE
	Dim sIndex As Double: sIndex = searchArray.Row - 1 'the absolute return range address
	Dim n As Long 'for array loop
	'format the search value for wildcards or not
	If (arg1 <> 2 And VarType(searchVal) = vbString) Then srchVal = Replace(Replace(Replace(srchVal, "*", "~*"), "?", "~?"), "#", "~#") 'for wildcard switch, escape if not
	'-----------------------
	Dim srchType As String
	Dim matchArg As Integer
	Dim lDirection As String
	Dim nextSize As String
	On Error GoTo error_control
	Select Case arg1 'work out the return mechanism from parameters, index match or array loop
		Case 0, 2
			If arg2 = 0 Or arg2 = 1 Then
				srchType = "im"
				matchArg = 0
			End If
		Case 1, -1
			nextSize = IIf(arg1 = -1, "s", "l") 'next smaller or larger
			If arg2 = 0 Or arg2 = 1 Then
				srchType = "lp"
				lDirection = "forward"
			End If
	End Select
	Select Case arg2 'get second parameter processing option
		Case -1
			srchType = "lp": lDirection = "reverse"
		Case 2
			srchType = "im": matchArg = 1
		Case -2
			srchType = "im": matchArg = -1
	End Select
	If srchType = "im" Then ' for index match return
		If rtnHeaderColumn Then
			Set CUSTOM_XLOOKUP_ONE = returnArray.Columns(WorksheetFunction.Match(srchVal, searchArray, matchArg))
		Else
			Set CUSTOM_XLOOKUP_ONE = returnArray.Rows(WorksheetFunction.Match(srchVal, searchArray, matchArg))
		End If
		Exit Function
	Else  'load search range into array for loop search
		Dim vArr As Variant: vArr = IIf(rtnHeaderColumn, WorksheetFunction.Transpose(searchArray), searchArray) 'assign the lookup range to an array
		Dim nsml As Variant: ' nsmal - next smallest value
		Dim nlrg As Variant: ' nlrg - next largest value
		Dim nStart As Double: nStart = IIf(lDirection = "forward", 1, UBound(vArr))
		Dim nEnd As Double: nEnd = IIf(lDirection = "forward", UBound(vArr), 1)
		Dim nStep As Integer: nStep = IIf(lDirection = "forward", 1, -1)
			For n = nStart To nEnd Step nStep
				If vArr(n, 1) Like srchVal Then Set CUSTOM_XLOOKUP_ONE = IIf(rtnHeaderColumn, returnArray.Columns(n), returnArray.Rows(n)): Exit Function 'exact match found
				If nsml < vArr(n, 1) And vArr(n, 1) < srchVal Then 'get next smallest
					Set nsml = searchArray.Rows(n)
				End If
				If vArr(n, 1) > srchVal And (IsEmpty(nlrg) Or nlrg > vArr(n, 1)) Then 'get next largest
					Set nlrg = IIf(rtnHeaderColumn, searchArray.Columns(n), searchArray.Rows(n))
				End If
			Next
	End If
	If arg1 = -1 Then 'next smallest
		Set CUSTOM_XLOOKUP_ONE = returnArray.Rows(nsml.Row - sIndex)
	ElseIf arg1 = 1 Then 'next largest
		Set CUSTOM_XLOOKUP_ONE = returnArray.Rows(nlrg.Row - sIndex)
	End If
	If Not IsEmpty(CUSTOM_XLOOKUP_ONE) Then Exit Function
	error_control:
	If IsMissing(notFound) Then
		CUSTOM_XLOOKUP_ONE = CVErr(xlErrNA)
	Else
		CUSTOM_XLOOKUP_ONE = [notFound]
	End If
End Function

Public Function BOOL_SWITCHABLE_XLOOKUP(ByVal lookup_value As Variant, ByVal lookup_array As Variant, ByVal return_array As Variant, Optional ByVal if_not_found As Variant = CVErr(xlErrNA), Optional ByVal match_mode As Variant = 0, Optional ByVal search_mode As Variant = 1, Optional ByVal xlookup_calc_switch As Boolean = True) As Variant
	Dim result As Variant

	If (xlookup_calc_switch = True) Then
		result = Application.WorksheetFunction.XLookup(lookup_value, lookup_array, return_array, if_not_found, match_mode, search_mode)
	Else
		result = "(Null)"
	End If

	BOOL_SWITCHABLE_XLOOKUP = result
End Function

'   @return File, Sheet, and Cell address of resultant
Public Function LOOKUP_TABLE_CELL_ADDRESS_V000(ByVal in_col_header As Range, ByVal lookup_value As Variant, ByVal lookup_col As Range, Optional ByVal if_not_found As Variant = CVErr(xlErrNA), Optional ByVal match_mode As Variant = 0, Optional ByVal search_mode As Variant = 1, Optional ByVal volatile_enable As Boolean = False) As String
    Dim result As String
    Dim col_header As Range
    Dim return_cell_row_position As Long
    Dim return_cell As Range
	Dim return_cell_sheet_name As String

    Set col_header = CORRECT_SINGLE_CELL_IN_RANGE_V000(in_col_header)
    '   result = Application.WorksheetFunction.Cell("address", Application.WorksheetFunction.XLookup(lookup_value, lookup_col, lookup_col, if_not_found, match_mode, search_mode))
    return_cell_row_position = Application.WorksheetFunction.Match(lookup_value, lookup_col, match_mode)
    Set return_cell = col_header.Offset(return_cell_row_position, 0)
    
	'	Debug.Print "Return Cell: " & return_cell

	return_cell_sheet_name = RETURN_SHEET_NAME_FROM_RANGE_V000(return_cell, True, volatile_enable)

	'result = Application.WorksheetFunction.Cell("address", return_cell)
	result = return_cell_sheet_name & return_cell.Address

    LOOKUP_TABLE_CELL_ADDRESS_V000 = result
End Function

Public Function BOOLSW_LOOKUP_TABLE_CELL_ADDR_V000(ByVal in_col_header As Range, ByVal lookup_value As Variant, ByVal lookup_col As Range, Optional ByVal if_not_found As Variant = CVErr(xlErrNA), Optional ByVal match_mode As Variant = 0, Optional ByVal search_mode As Variant = 1, Optional ByVal volatile_enable As Boolean = False, Optional ByVal calc_switch As Boolean = True) As String
    Dim result As String
    
    If (calc_switch = True) Then
		result = LOOKUP_TABLE_CELL_ADDRESS_V000(in_col_header, lookup_value, lookup_col, if_not_found, match_mode, search_mode, volatile_enable)
	Else
		result = "(Null)"
	End If

	BOOLSW_LOOKUP_TABLE_CELL_ADDR_V000 = result
End Function

'	could make a higher level function to perform the offset subtraction for this
Public Function RETURN_OFF_COL_VARIANT_FROM_ADDRESS_V000(ref_addr As String, col_offset_num As Long) As Variant
    Dim result As Variant

    If (Range(ref_addr).Cells.Count = 1) Then
        result = Range(ref_addr).Offset(0, col_offset_num)
    Else
       '     result = "#VALUE!"
        '   result = xlErrValue ...?
        result = CVErr(xlErrValue)
    End If

    RETURN_OFF_COL_VARIANT_FROM_ADDRESS_V000 = result
End Function

Public Function BOOLSW_RETURN_OFF_COL_VARIANT_FROM_ADDRESS_V000(ref_addr As String, col_offset_num As Long, Optional ByVal calc_switch As Boolean = True) As Variant
    Dim result As Variant
    If (calc_switch = True) Then
		result = RETURN_OFF_COL_VARIANT_FROM_ADDRESS_V000(ref_addr, col_offset_num)
	Else
		result = "(Null)"
	End If
    BOOLSW_RETURN_OFF_COL_VARIANT_FROM_ADDRESS_V000 = result
End Function

'   https://www.youtube.com/watch?v=nV_oDWJccu8&list=PLmHVyfmcRKyzmbDy6QoBuUDrU5D-jD-Se&index=11
'   https://docs.microsoft.com/en-us/office/vba/api/excel.range.find
'   https://excelmacromastery.com/excel-vba-find/
'   https://docs.microsoft.com/en-us/office/vba/api/excel.range.findnext
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/join-function
Public Function CONCAT_FIND_RESULTS_000_V000(find_range As Range, lookup_value As Variant, find_type As Variant, cell_match_type As Variant, search_order As Variant, search_direction As Variant, case_sensitive As Variant, match_byte As Variant, search_format As Variant) As String
	Dim result As String
	
	Dim found_cell As Range
	Dim first_cell_address As String
	Dim first_found_value As Variant
	Dim found_value As Variant
	Dim found_string As String
	Dim found_string_length As Long
	Dim found_ctr As Long
	Dim found_array() As Variant
	Dim result_ctr As Long
	Dim string_limit As Long
	Dim concat_string_length As Long
	Dim concat_test As String
	Dim expected_string_length As Long
	Dim string_limit_reached As Boolean

	'   safe limit
	'   string_limit = 2^31 - 1
    '   safer limit
    string_limit = 1 * 10^9
	string_limit_reached = False

	concat_string_length = 0
	
	found_ctr = 0
	result_ctr = 0
	ReDim found_array(result_ctr)

	' '   Set found_cell = find_range.Find(lookup_value, , find_type, cell_match_type, search_order, search_direction, case_sensitive, match_byte, search_format)
	' Set found_cell = find_range.Find(CAST_VARIANT_TO_STRING(lookup_value), , xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)

	' If Not found_cell Is Nothing Then
	'     first_cell_address = found_cell.Address
	'     found_array(result_ctr) = found_cell.Value2
	'     Set found_cell = find_range.FindNext(found_cell)
	' Else
	'     '   ...
	' End If

	' Debug.Print found_cell.Address
	' Debug.Print found_cell.Value2

	' While (Not found_cell Is Nothing) And (Len(result) < string_limit)
	'     result_ctr = result_ctr + 1
	'     ReDim found_array(result_ctr)
	'     found_array(result_ctr) = found_cell.Value2
	'     result = Join(found_array, " , ")
	'     Set found_cell = find_range.FindNext(found_cell)
	'     Debug.Print found_cell.Address
	'     Debug.Print found_cell.Value2
	' Wend

	' If (Not found_cell Is Nothing) And (Len(result) = string_limit) Then
	'     Set found_cell = find_range.FindNext(found_cell)
	'     If (Not found_cell Is Nothing) Then
	'         result = "Error: result overflow!"
	'     End If
	' End If

	With find_range
		'   Set found_cell = .Find(what:="o", lookin:=xlValues)
		Set found_cell = .Find(CAST_VARIANT_TO_STRING(lookup_value), .Cells(.Rows.Count,1), xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
		'   Set found_cell = .Find("o", Worksheets(1).Range("A1:A1") , xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
		'	Debug.Print found_cell.Address
		'	Debug.Print found_cell.Value2
		If Not found_cell Is Nothing Then 
			first_cell_address = found_cell.Address
			Do 
				'	Debug.Print found_cell.Address
				'	Debug.Print found_cell.Value2
				found_ctr = found_ctr + 1
				result_ctr = result_ctr + 1

				'   Set found_cell = .FindNext(found_cell)
				Set found_cell = .Find(CAST_VARIANT_TO_STRING(lookup_value), found_cell, xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
				
				' If (Not found_cell Is Nothing) And (found_cell.Address <> first_cell_address) Then
				'     result_ctr = result_ctr + 1
				' Else
				' End If
			Loop While (Not found_cell Is Nothing) And (found_cell.Address <> first_cell_address)
		End If 
	End With

	ReDim found_array(0 To (result_ctr - 1))

	result_ctr = 0

	With find_range
		'   Set found_cell = .Find(what:="o", lookin:=xlValues)
		Set found_cell = .Find(CAST_VARIANT_TO_STRING(lookup_value), .Cells(.Rows.Count,1), xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
		'   Set found_cell = .Find("o", Worksheets(1).Range("A1:A1") , xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
		'   Debug.Print found_cell.Address
		'   Debug.Print found_cell.Value2
		If Not found_cell Is Nothing Then 
			first_cell_address = found_cell.Address
			Do 
				'   Debug.Print found_cell.Address
				'   Debug.Print found_cell.Value2
				found_ctr = found_ctr + 1
				result_ctr = result_ctr + 1

				' concat_test = Join(found_array, " , ")
				' concat_string_length = Len(concat_test)
				
				found_value = found_cell.Value2
				found_string = CAST_VARIANT_TO_STRING(found_value)
				found_string_length = Len(found_string)

				concat_string_length = concat_string_length + found_string_length + Len(" , ")

				'	expected_string_length = concat_string_length + found_string_length

				If (concat_string_length >= string_limit) Then
					string_limit_reached = True
				Else
					string_limit_reached = False
				End If

				found_array(result_ctr - 1) = found_cell.Value2
				'   Set found_cell = .FindNext(found_cell)
				Set found_cell = .Find(CAST_VARIANT_TO_STRING(lookup_value), found_cell, xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
				
				' If (Not found_cell Is Nothing) And (found_cell.Address <> first_cell_address) Then
				'     result_ctr = result_ctr + 1
				' Else
				' End If

			Loop While (Not found_cell Is Nothing) And (found_cell.Address <> first_cell_address)
		End If 
	End With

	If (string_limit_reached = False) Then
		result = Join(found_array, " , ")
	Else
		result = "Error: result overflow!"
	End If

	CONCAT_FIND_RESULTS_000_V000 = result
End Function

Public Function CONVERT_VALUE_ANGLE_DEG_TO_RAD(input_angle_deg As Double) As Double
	CONVERT_VALUE_ANGLE_DEG_TO_RAD = input_angle_deg * ((WorksheetFunction.Pi()) / 180)
End Function

Public Function CONVERT_VALUE_ANGLE_RAD_TO_DEG(input_angle_rad As Double) As Double
	CONVERT_VALUE_ANGLE_RAD_TO_DEG = input_angle_rad * (180 / (WorksheetFunction.Pi()))
End Function