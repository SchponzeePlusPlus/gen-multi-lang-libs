Attribute VB_Name = "GeneralDataTypeXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   GeneralDataTypeXLModule
'   Leonard Sponza
'   Last Modified 12/06/2021 15:20
'   Date Time Version 00

Option Explicit

Public Enum CustomBoolean
    TRUE_CUST_BOOL = 1
    FALSE_CUST_BOOL = 2
    UNASSIGNED_CUST_BOOL = UNASSIGNED_LONG_VAL
    NULL_CUST_BOOL = NULL_LONG_VAL
    NOT_APPLICABLE_CUST_BOOL = NOT_APPLICABLE_LONG_VAL
    ERROR_CUST_BOOL = ERROR_LONG_VAL
    TEST_CUST_BOOL = TEST_LONG_VAL
End Enum

Public Type ArrayDimensionsType
    l_bnd As Long
    u_bnd As Long
    length As Long
End Type

Public Function CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_long_od_arr() As Long) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

    result.l_bnd = LBound(input_long_od_arr, 1)
    result.u_bnd = UBound(input_long_od_arr, 1)

    result.length = result.u_bnd + result.l_bnd + 1

	CREATE_LONG_ONE_DIM_ARRAY_DIM_TYPE = result
End Function

Public Function CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_double_od_arr() As Double) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

    result.l_bnd = LBound(input_double_od_arr, 1)
    result.u_bnd = UBound(input_double_od_arr, 1)

    result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DOUBLE_ONE_DIM_ARRAY_DIM_TYPE = result
End Function