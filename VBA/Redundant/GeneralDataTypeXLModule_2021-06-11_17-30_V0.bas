Attribute VB_Name = "GeneralDataTypeXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   GeneralDataTypeXLModule
'   Leonard Sponza
'   Last Modified 11/06/2021 12:30
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
    arr_dim_l_bnd As Long
    arr_dim_u_bnd As Long
    arr_dim_length As Long
End Type

'	I feel like these are bad

Public Type CustomLongOneDimArray
    long_od_arr() As Long
    long_od_arr_l_bnd As Long
    long_od_arr_u_bnd As Long
    long_od_arr_length As Long
End Type

Public Type CustomDoubleOneDimArray
    doub_od_arr() As Double
    doub_od_arr_l_bnd As Long
    doub_od_arr_u_bnd As Long
    doub_od_arr_length As Long
End Type

Public Function CREATE_CUSTOM_LONG_ONE_DIM_ARRAY(input_long_od_arr() As Long) As CustomLongOneDimArray
	Dim result As CustomLongOneDimArray

	CREATE_CUSTOM_LONG_ONE_DIM_ARRAY = result
End Function