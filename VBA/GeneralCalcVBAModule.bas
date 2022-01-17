Attribute VB_Name = "GeneralCalcVBAModule"
'   MS Office VBA
'   General Use Calculation Module (for all MS applications)
'   GeneralCalcVBAModule
'   Leonard Sponza
'   Last Modified 15/09/2021 13:00
'   Date Time Version 00

'   Requires:
'       - GeneralXLModule

Option Explicit

'   https://www.mrexcel.com/board/threads/max-min-vba.132404/
Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, Value As Variant
   minValue = values(0)
   For Each Value In values
	   If Value < minValue Then minValue = Value
   Next
   Min = minValue
End Function

Public Function RETURN_MIN_VAL_FROM_DOUBLE_OD_ARR_V00(input_arr() As Double) As Double
	Dim result As Double
	Dim input_arr_odadt As ArrayDimensionsType
	Dim min_val As Double
	Dim i As Long

	input_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_arr)
	min_val = input_arr(0)
	For i = (input_arr_odadt.l_bnd) To (input_arr_odadt.u_bnd)
		If (input_arr(i) < min_val) Then
			min_val = input_arr(i)
		End If
	Next
	result = min_val
	RETURN_MIN_VAL_FROM_DOUBLE_OD_ARR_V00 = result
End Function

'   https://www.mrexcel.com/board/threads/max-min-vba.132404/
Function Max(ParamArray values() As Variant) As Variant
	Dim maxValue, Value As Variant
	maxValue = values(0)
	For Each Value In values
		If Value > maxValue Then maxValue = Value
	Next
	Max = maxValue
End Function

Public Function RETURN_MAX_VAL_FROM_DOUBLE_OD_ARR_V00(input_arr() As Double) As Double
	Dim result As Double
	Dim input_arr_odadt As ArrayDimensionsType
	Dim max_val As Double
	Dim i As Long

	input_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_arr)
	max_val = input_arr(0)
	For i = (input_arr_odadt.l_bnd) To (input_arr_odadt.u_bnd)
		If (input_arr(i) > max_val) Then
			max_val = input_arr(i)
		End If
	Next
	result = max_val
	RETURN_MAX_VAL_FROM_DOUBLE_OD_ARR_V00 = result
End Function