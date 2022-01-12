Attribute VB_Name = "GeneralMathVBAModule"
'
'	@file GeneralMathVBAModule.bas
'
'	Write description of source file here for dOxygen.
'
'	System:         Independant Tool
'	Component Name: Suspended Magnet Selection Program Rev0 (Excel), Module Tramp Height Magnet (THM)
'  
'	Language: VBA (General Use for ~all MS Office Applications)
'
'	@brief Can use "brief" tag to explicitly generate comments for file documentation.
'	@author Leonard Sponza
'	@version 0.41.0

'	License: MIT? 
'	Licensed Material - N/A
'	Open-Source Code
'	Address:
'		
'	Author E-Mail: 
'
'	Description / Abstract: Module file for Suspended Magnet Selection Program Rev0 (Excel)
'		This file contains the defined types for Project X:
'		Notes:
'
'	Limitations: _
'	Function:
'		1) _
'	
'	Database tables used: VBA does not refer directly to DB
'	Thread Safe: No
'	Extendable: No
'	Platform Dependencies: Microsoft Excel 32-bit
'	Compiler Options: N/A
'	Change History / Revisions:
'
'	Date			Time		Author       		Description
'	----------------------------------------------------------------------------
'	05/11/2021		16:20		Leonard Sponza		GeneralMathModule created
'                                                   functions moved to here
'   08/11/2021      15:50                           Renamed to GeneralMathVBAModule
'
'
'	----------------------------------------------------------------------------
'
'	Requires:
'		- GeneralCalcVBAModule
'		- GeneralXLModule

Option Explicit

'   Pi accuracy source: https://groups.google.com/g/microsoft.public.excel.programming/c/LqH10-6QfbM?pli=1
Public Const PI_VBA_DOUBLE As Double = 3.14159265358979

'   Useful for interpolation / extrapolation checking
Public Enum RefDataLinearIntervalChk
    WITHIN_REF_DATA = 1
    ABOVE_REF_DATA = 2
    BELOW_REF_DATA = 3
    EQUAL_TO_REF_DATA = 4
    '   Multiple occurences where lookup value equals a ref data val; ref data has duplicate values in field / dimension
    MULTI_MATCH_TO_REF_DATA = 5
    UNASSIGNED_RDLIC = UNASSIGNED_LONG_VAL
    NULL_INTERVAL_CHK = NULL_LONG_VAL
    ERROR_INTERVAL_CHK = ERROR_LONG_VAL
End Enum

'   Data Analysis Function Condition
Public Enum DataAnalysisFuncCondEnum
    DIRECT_ASSIGNMENT = 1
    '   LINEAR_GRADIENT_VIA_CLOSE_DATA ? for both inter and extra
    LINEAR_INTERPOLATION_VIA_CLOSE_DATA = 2
    LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA = 3
    EXPONENTIAL = 4
    DO_NOT_CALCULATE = 5
    UNASSIGNED_DAFC = UNASSIGNED_LONG_VAL
    NULL_DAFC = NULL_LONG_VAL
    ERROR_DAFC = ERROR_LONG_VAL
End Enum

Public Enum EntityAnalysisMethodEnum
    SPECIFIC_EAM = 1
    INDEPENDANT_EAM = 2
    UNASSIGNED_EAM = UNASSIGNED_LONG_VAL
    NULL_EAM = NULL_LONG_VAL
    ERROR_EAM = ERROR_LONG_VAL
End Enum

Public Enum EntitySelectionMethodEnum
    SINGLE_ESM = 1
    MULTI_ESM = 2
    RANGE_ESM = 4
    ALL_ESM = 3
    UNASSIGNED_ESM = UNASSIGNED_LONG_VAL
    NULL_ESM = NULL_LONG_VAL
    ERROR_ESM = ERROR_LONG_VAL
End Enum

'	Routine: NameOfRoutine()
'
'	Inputs:
'		@param
'		Externals:
'		Others:
'
'	Outputs:
'		Arguments:
'		Externals:
'		@return
'		@bug
'		Errors:
'
'	Routines Called:

Public Type PointTwoDimDoubleType
    x As Double
    y As Double
End Type

Public Type PointThrDimDoubleType
    xy As PointTwoDimDoubleType
    z As Double
End Type

Public Type FuncDataFXDoubleType
    in_x As Double
    out_f As Double
End Type

Public Type FuncDataFYDoubleType
    in_y As Double
    out_f As Double
End Type

Public Type FuncDataFXYDoubleType
    in_xy As PointTwoDimDoubleType
    out_f As Double
End Type

Public Type DBFuncDataFYDoubleType
    id As Long
    sort_num As Long
    filter_code As String
    data As FuncDataFYDoubleType
End Type

'   i and i + 1
Public Type DBFuncDataFYIPOneDoubleType
    id As Long
    sort_num As Long
    filter_code As String
    data_i As FuncDataFYDoubleType
    data_i_p_one As FuncDataFYDoubleType
End Type

'   i Plus/Minus Two
Public Type DBFuncDataFYIPMTwoDoubleType
    id As Long
    sort_num As Long
    filter_code As String
    data_i As FuncDataFYDoubleType
    data_i_p_one As FuncDataFYDoubleType
    data_i_p_two As FuncDataFYDoubleType
    data_i_m_one As FuncDataFYDoubleType
    data_i_m_two As FuncDataFYDoubleType
End Type

Public Type EntityDataAnalysisRulesType
    within_rd_dafc As DataAnalysisFuncCondEnum
    above_rd_dafc As DataAnalysisFuncCondEnum
    below_rd_dafc As DataAnalysisFuncCondEnum
    equal_to_rd_dafc As DataAnalysisFuncCondEnum
    multi_match_to_rd_dafc As DataAnalysisFuncCondEnum
End Type

'   Ref Data Handling Functions

Public Function ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING(input_str As String) As Long
    Select Case input_str
        Case "TRUE"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.WITHIN_REF_DATA
        Case "True"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.WITHIN_REF_DATA
        Case "WITHIN_REF_DATA"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.WITHIN_REF_DATA
        Case "ABOVE_REF_DATA"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.ABOVE_REF_DATA
        Case "BELOW_REF_DATA"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.BELOW_REF_DATA
        Case "NULL_INTERVAL_CHK"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.NULL_INTERVAL_CHK
        Case "ERROR_INTERVAL_CHK"
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.ERROR_INTERVAL_CHK
        Case Else
            ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING = RefDataLinearIntervalChk.ERROR_INTERVAL_CHK
    End Select
End Function

Public Function ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING_V000(input_str As String) As RefDataLinearIntervalChk
    Dim result As RefDataLinearIntervalChk

    Select Case input_str
        Case "TRUE"
            result = WITHIN_REF_DATA
        Case "True"
            result = WITHIN_REF_DATA
        Case "WITHIN_REF_DATA"
            result = WITHIN_REF_DATA
        Case "ABOVE_REF_DATA"
            result = ABOVE_REF_DATA
        Case "BELOW_REF_DATA"
            result = BELOW_REF_DATA
        Case "EQUAL_TO_REF_DATA"
            result = EQUAL_TO_REF_DATA
        Case "MULTI_MATCH_TO_REF_DATA"
            result = MULTI_MATCH_TO_REF_DATA
        Case "NULL_INTERVAL_CHK"
            result = NULL_INTERVAL_CHK
        Case "ERROR_INTERVAL_CHK"
            result = ERROR_INTERVAL_CHK
        Case Else
            result = ERROR_INTERVAL_CHK
    End Select

    ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING_V000 = result
End Function

Public Function ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM(input_ref_data_linear_interval_chk As RefDataLinearIntervalChk) As String
    Select Case input_ref_data_linear_interval_chk
        Case WITHIN_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "WITHIN_REF_DATA"
        Case ABOVE_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ABOVE_REF_DATA"
        Case BELOW_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "BELOW_REF_DATA"
        Case EQUAL_TO_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "EQUAL_TO_REF_DATA"
        Case MULTI_MATCH_TO_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "MULTI_MATCH_TO_REF_DATA"
        Case NULL_INTERVAL_CHK
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "NULL_INTERVAL_CHK"
        Case ERROR_INTERVAL_CHK
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ERROR_INTERVAL_CHK"
        Case Else
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ERROR_INTERVAL_CHK"
    End Select
End Function

Public Function ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(input_str As String) As DataAnalysisFuncCondEnum
    Dim result As DataAnalysisFuncCondEnum
    Select Case input_str
        Case "DIRECT_ASSIGNMENT"
            result = DIRECT_ASSIGNMENT
        Case "LINEAR_INTERPOLATION_VIA_CLOSE_DATA"
            result = LINEAR_INTERPOLATION_VIA_CLOSE_DATA
        Case "LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA"
            result = LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA
        Case "EXPONENTIAL"
            result = EXPONENTIAL
        Case "DO_NOT_CALCULATE"
            result = DO_NOT_CALCULATE
        Case "UNASSIGNED_DAFC"
            result = UNASSIGNED_DAFC
        Case "NULL_DAFC"
            result = NULL_DAFC
        Case "ERROR_DAFC"
            result = ERROR_DAFC
        Case Else
            result = ERROR_DAFC
    End Select
    ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING = result
End Function

Public Function ASSIGN_ENTITYANALYSISMETHOD_ENUM_FROM_STRING(input_str As String) As EntityAnalysisMethodEnum
    Dim result As EntityAnalysisMethodEnum

    result = UNASSIGNED_EAM
    Select Case input_str
        Case "SPECIFIC_EAM"
            result = SPECIFIC_EAM
        Case "INDEPENDANT_EAM"
            result = INDEPENDANT_EAM
        Case "UNASSIGNED_EAM"
            result = UNASSIGNED_EAM
        Case "NULL_EAM"
            result = NULL_EAM
        Case "ERROR_EAM"
            result = ERROR_EAM
        Case Else
            result = ERROR_EAM
    End Select

    ASSIGN_ENTITYANALYSISMETHOD_ENUM_FROM_STRING = result
End Function

Public Function ASSIGN_ENTITYSELECTIONMETHOD_ENUM_FROM_STRING(input_str As String) As EntitySelectionMethodEnum
    Dim result As EntitySelectionMethodEnum

    result = UNASSIGNED_EAM
    Select Case input_str
        Case "SINGLE_ESM"
            result = SINGLE_ESM
        Case "MULTI_ESM"
            result = MULTI_ESM
        Case "RANGE_ESM"
            result = RANGE_ESM
        Case "ALL_ESM"
            result = ALL_ESM
        Case "UNASSIGNED_ESM"
            result = UNASSIGNED_ESM
        Case "NULL_ESM"
            result = NULL_ESM
        Case "ERROR_ESM"
            result = ERROR_ESM
        Case Else
            result = ERROR_ESM
    End Select

    ASSIGN_ENTITYSELECTIONMETHOD_ENUM_FROM_STRING = result
End Function

Public Function CREATE_FUNCDATAFYDOUBLETYPE(input_in_y As Double, input_out_f As Double) As FuncDataFYDoubleType
    Dim result As FuncDataFYDoubleType

    result.in_y = input_in_y
    result.out_f = input_out_f

    CREATE_FUNCDATAFYDOUBLETYPE = result
End Function

Public Function NULLIFY_FUNCDATAFYDOUBLETYPE() As FuncDataFYDoubleType
    Dim result As FuncDataFYDoubleType

    result.in_y = NULL_DOUBLE_VAL
    result.out_f = NULL_DOUBLE_VAL

    NULLIFY_FUNCDATAFYDOUBLETYPE = result
End Function

'   erroneously
Public Function ERRORIFY_FUNCDATAFYDOUBLETYPE() As FuncDataFYDoubleType
    Dim result As FuncDataFYDoubleType

    result.in_y = ERROR_DOUBLE_VAL
    result.out_f = ERROR_DOUBLE_VAL

    ERRORIFY_FUNCDATAFYDOUBLETYPE = result
End Function

Public Function JOIN_FUNCDATAFYDOUBLETYPE_TO_STRING_V000(input_type As FuncDataFYDoubleType) As String
    Dim result As String

    '   ... function incomplete

    JOIN_FUNCDATAFYDOUBLETYPE_TO_STRING_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE(input_id As Long, input_sort_num As Long, input_in_y As Double, input_out_f As Double) As DBFuncDataFYDoubleType
    Dim result As DBFuncDataFYDoubleType

    result.id = input_id
    result.sort_num = input_sort_num
    result.data = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y, input_out_f)

    CREATE_DBFUNCDATAFYDOUBLETYPE = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE_V001(input_id As Long, input_sort_num As Long, input_filter_code As String, input_in_y As Double, input_out_f As Double) As DBFuncDataFYDoubleType
    Dim result As DBFuncDataFYDoubleType

    result.id = input_id
    result.sort_num = input_sort_num
    result.filter_code = input_filter_code
    result.data = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y, input_out_f)

    CREATE_DBFUNCDATAFYDOUBLETYPE_V001 = result
End Function

Public Function ERRORIFY_DBFUNCDATAFYDOUBLETYPE_V000() As DBFuncDataFYDoubleType
    Dim result As DBFuncDataFYDoubleType

    result.id = ERROR_LONG_VAL
    result.sort_num = ERROR_LONG_VAL
    result.filter_code = ERROR_STRING_VAL
    result.data = ERRORIFY_FUNCDATAFYDOUBLETYPE()

    ERRORIFY_DBFUNCDATAFYDOUBLETYPE_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE_ARR_V000(input_id_arr() As Long, input_sort_num_arr() As Long, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double) As DBFuncDataFYDoubleType()
    Dim result() As DBFuncDataFYDoubleType
    Dim result_arr_odadt As ArrayDimensionsType
    Dim input_id_arr_odadt As ArrayDimensionsType
    Dim input_sort_num_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_arr_odadt As ArrayDimensionsType
    Dim equal_arr_lengths As Boolean
    Dim i As Long

    equal_arr_lengths = False
    input_id_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
    input_sort_num_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_sort_num_arr)
    input_in_y_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_arr)
    input_out_f_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_arr)

    If ((input_id_arr_odadt.length = input_sort_num_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_arr_odadt.length)) Then
        equal_arr_lengths = True
    Else
        equal_arr_lengths = False
    End If

    If (equal_arr_lengths = True) Then
        result_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
        ReDim result(result_arr_odadt.l_bnd To result_arr_odadt.u_bnd)

        For i = result_arr_odadt.l_bnd To result_arr_odadt.u_bnd
		    result(i) = CREATE_DBFUNCDATAFYDOUBLETYPE(input_id_arr(i), input_sort_num_arr(i), input_in_y_i_arr(i), input_out_f_i_arr(i))
	    Next
    Else
        ' nullify type
    End If

    CREATE_DBFUNCDATAFYDOUBLETYPE_ARR_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE_ARR_V001(input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double) As DBFuncDataFYDoubleType()
    Dim result() As DBFuncDataFYDoubleType
    Dim result_arr_odadt As ArrayDimensionsType
    Dim input_id_arr_odadt As ArrayDimensionsType
    Dim input_sort_num_arr_odadt As ArrayDimensionsType
    Dim input_filter_code_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_arr_odadt As ArrayDimensionsType
    Dim equal_arr_lengths As Boolean
    Dim i As Long

    equal_arr_lengths = False
    input_id_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
    input_sort_num_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_sort_num_arr)
    input_filter_code_arr_odadt = CREATE_STRING_ONE_DIM_ARRAYDIMSTYPE(input_filter_code_arr)
    input_in_y_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_arr)
    input_out_f_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_arr)

    If ((input_id_arr_odadt.length = input_sort_num_arr_odadt.length) And (input_id_arr_odadt.length = input_filter_code_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_arr_odadt.length)) Then
        equal_arr_lengths = True
    Else
        equal_arr_lengths = False
    End If

    If (equal_arr_lengths = True) Then
        result_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
        ReDim result(result_arr_odadt.l_bnd To result_arr_odadt.u_bnd)

        For i = result_arr_odadt.l_bnd To result_arr_odadt.u_bnd
		    result(i) = CREATE_DBFUNCDATAFYDOUBLETYPE_V001(input_id_arr(i), input_sort_num_arr(i), input_filter_code_arr(i), input_in_y_i_arr(i), input_out_f_i_arr(i))
	    Next
    Else
        ' nullify type
    End If

    CREATE_DBFUNCDATAFYDOUBLETYPE_ARR_V001 = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_type_arr() As DBFuncDataFYDoubleType) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_type_arr, 1)
	result.u_bnd = UBound(input_type_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DBFUNCDATAFYDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE(input_id As Long, input_sort_num As Long, input_in_y_i As Double, input_out_f_i As Double, input_in_y_i_p_one As Double, input_out_f_i_p_one As Double) As DBFuncDataFYIPOneDoubleType
    Dim result As DBFuncDataFYIPOneDoubleType

    result.id = input_id
    result.sort_num = input_sort_num
    result.data_i = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i, input_out_f_i)
    result.data_i_p_one = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_p_one, input_out_f_i_p_one)

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_V001(input_id As Long, input_sort_num As Long, input_filter_code As String, input_in_y_i As Double, input_out_f_i As Double, input_in_y_i_p_one As Double, input_out_f_i_p_one As Double) As DBFuncDataFYIPOneDoubleType
    Dim result As DBFuncDataFYIPOneDoubleType

    result.id = input_id
    result.sort_num = input_sort_num
    result.filter_code = input_filter_code
    result.data_i = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i, input_out_f_i)
    result.data_i_p_one = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_p_one, input_out_f_i_p_one)

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_V001 = result
End Function

Public Function NULLIFY_DBFUNCDATAFYIPONEDOUBLETYPE_V000() As DBFuncDataFYIPOneDoubleType
    Dim result As DBFuncDataFYIPOneDoubleType

    result.id = NULL_LONG_VAL
    result.sort_num = NULL_LONG_VAL
    result.filter_code = NULL_STRING_VAL
    result.data_i = NULLIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_p_one = NULLIFY_FUNCDATAFYDOUBLETYPE()

    NULLIFY_DBFUNCDATAFYIPONEDOUBLETYPE_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR(input_id_arr() As Long, input_sort_num_arr() As Long, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim result_arr_odadt As ArrayDimensionsType
    Dim input_id_arr_odadt As ArrayDimensionsType
    Dim input_sort_num_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_p_one_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_p_one_arr_odadt As ArrayDimensionsType
    Dim equal_arr_lengths As Boolean
    Dim i As Long

    equal_arr_lengths = False
    input_id_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
    input_sort_num_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_sort_num_arr)
    input_in_y_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_arr)
    input_out_f_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_arr)
    input_in_y_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_p_one_arr)
    input_out_f_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_p_one_arr)

    If ((input_id_arr_odadt.length = input_sort_num_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_p_one_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_p_one_arr_odadt.length)) Then
        equal_arr_lengths = True
    Else
        equal_arr_lengths = False
    End If

    If (equal_arr_lengths = True) Then
        result_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
        ReDim result(result_arr_odadt.l_bnd To result_arr_odadt.u_bnd)

        For i = result_arr_odadt.l_bnd To result_arr_odadt.u_bnd
		    result(i) = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE(input_id_arr(i), input_sort_num_arr(i), input_in_y_i_arr(i), input_out_f_i_arr(i), input_in_y_i_p_one_arr(i), input_out_f_i_p_one_arr(i))
	    Next
    Else
        ' nullify type
    End If

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V001(input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim result_arr_odadt As ArrayDimensionsType
    Dim input_id_arr_odadt As ArrayDimensionsType
    Dim input_sort_num_arr_odadt As ArrayDimensionsType
    Dim input_filter_code_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_p_one_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_p_one_arr_odadt As ArrayDimensionsType
    Dim equal_arr_lengths As Boolean
    Dim i As Long

    equal_arr_lengths = False
    input_id_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
    input_sort_num_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_sort_num_arr)
    input_filter_code_arr_odadt = CREATE_STRING_ONE_DIM_ARRAYDIMSTYPE(input_filter_code_arr)
    input_in_y_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_arr)
    input_out_f_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_arr)
    input_in_y_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_p_one_arr)
    input_out_f_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_p_one_arr)

    If ((input_id_arr_odadt.length = input_sort_num_arr_odadt.length) And (input_id_arr_odadt.length = input_filter_code_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_p_one_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_p_one_arr_odadt.length)) Then
        equal_arr_lengths = True
    Else
        equal_arr_lengths = False
    End If

    If (equal_arr_lengths = True) Then
        result_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
        ReDim result(result_arr_odadt.l_bnd To result_arr_odadt.u_bnd)

        For i = result_arr_odadt.l_bnd To result_arr_odadt.u_bnd
		    result(i) = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_V001(input_id_arr(i), input_sort_num_arr(i), input_filter_code_arr(i), input_in_y_i_arr(i), input_out_f_i_arr(i), input_in_y_i_p_one_arr(i), input_out_f_i_p_one_arr(i))
	    Next
    Else
        ' nullify type
    End If

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V001 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_type_arr() As DBFuncDataFYIPOneDoubleType) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_type_arr, 1)
	result.u_bnd = UBound(input_type_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE = result
End Function

Public Function FILTER_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000(lookup_filter_val As String, input_type_arr() As DBFuncDataFYIPOneDoubleType) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim input_type_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim filtered_arr_ctr As Long

    input_type_arr_odadt = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_type_arr)

    i = UNASSIGNED_LONG_VAL
    filtered_arr_ctr = 0

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        If (lookup_filter_val = input_type_arr(i).filter_code) Then
            filtered_arr_ctr = filtered_arr_ctr + 1
        Else
            '   do nothing
        End If
    Next

    ReDim result(0 To filtered_arr_ctr)

    filtered_arr_ctr = 0

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        If (lookup_filter_val = input_type_arr(i).filter_code) Then
            result(filtered_arr_ctr) = input_type_arr(i)
            filtered_arr_ctr = filtered_arr_ctr + 1
        Else
            '   do nothing
        End If
    Next

    FILTER_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000 = result
End Function

'   function not developed, only copied from filter func
Public Function SORT_BY_SN_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000(lookup_filter_val As String, input_type_arr() As DBFuncDataFYIPOneDoubleType) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim input_type_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim filtered_arr_ctr As Long

    input_type_arr_odadt = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_type_arr)

    i = UNASSIGNED_LONG_VAL
    filtered_arr_ctr = 0

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        If (lookup_filter_val = input_type_arr(i).filter_code) Then
            filtered_arr_ctr = filtered_arr_ctr + 1
        Else
            '   do nothing
        End If
    Next

    ReDim result(0 To filtered_arr_ctr)

    filtered_arr_ctr = 0

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        If (lookup_filter_val = input_type_arr(i).filter_code) Then
            result(filtered_arr_ctr) = input_type_arr(i)
            filtered_arr_ctr = filtered_arr_ctr + 1
        Else
            '   do nothing
        End If
    Next

    SORT_BY_SN_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_V000(input_id As Long, input_sort_num As Long, input_filter_code As String, input_in_y_i As Double, input_out_f_i As Double, input_in_y_i_p_one As Double, input_out_f_i_p_one As Double, input_in_y_i_p_two As Double, input_out_f_i_p_two As Double, input_in_y_i_m_one As Double, input_out_f_i_m_one As Double, input_in_y_i_m_two As Double, input_out_f_i_m_two As Double) As DBFuncDataFYIPMTwoDoubleType
    Dim result As DBFuncDataFYIPMTwoDoubleType

    result.id = input_id
    result.sort_num = input_sort_num
    result.filter_code = input_filter_code
    result.data_i = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i, input_out_f_i)
    result.data_i_p_one = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_p_one, input_out_f_i_p_one)
    result.data_i_p_two = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_p_two, input_out_f_i_p_two)
    result.data_i_m_one = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_m_one, input_out_f_i_m_one)
    result.data_i_m_two = CREATE_FUNCDATAFYDOUBLETYPE(input_in_y_i_m_two, input_out_f_i_m_two)

    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_V000 = result
End Function

Public Function NULLIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000() As DBFuncDataFYIPMTwoDoubleType
    Dim result As DBFuncDataFYIPMTwoDoubleType

    result.id = NULL_LONG_VAL
    result.sort_num = NULL_LONG_VAL
    result.filter_code = NULL_STRING_VAL
    result.data_i = NULLIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_p_one = NULLIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_p_two = NULLIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_m_one = NULLIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_m_two = NULLIFY_FUNCDATAFYDOUBLETYPE()

    NULLIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000 = result
End Function

Public Function ERRORIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000() As DBFuncDataFYIPMTwoDoubleType
    Dim result As DBFuncDataFYIPMTwoDoubleType

    result.id = NULL_LONG_VAL
    result.sort_num = NULL_LONG_VAL
    result.filter_code = NULL_STRING_VAL
    result.data_i = ERRORIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_p_one = ERRORIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_p_two = ERRORIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_m_one = ERRORIFY_FUNCDATAFYDOUBLETYPE()
    result.data_i_m_two = ERRORIFY_FUNCDATAFYDOUBLETYPE()

    ERRORIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double, _
    input_in_y_i_p_two_arr() As Double, input_out_f_i_p_two_arr() As Double, input_in_y_i_m_one_arr() As Double, input_out_f_i_m_one_arr() As Double, input_in_y_i_m_two_arr() As Double, input_out_f_i_m_two_arr() As Double) As DBFuncDataFYIPMTwoDoubleType()
    
    Dim result() As DBFuncDataFYIPMTwoDoubleType
    Dim result_arr_odadt As ArrayDimensionsType
    Dim input_id_arr_odadt As ArrayDimensionsType
    Dim input_sort_num_arr_odadt As ArrayDimensionsType
    Dim input_filter_code_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_p_one_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_p_one_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_p_two_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_p_two_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_m_one_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_m_one_arr_odadt As ArrayDimensionsType
    Dim input_in_y_i_m_two_arr_odadt As ArrayDimensionsType
    Dim input_out_f_i_m_two_arr_odadt As ArrayDimensionsType
    Dim equal_arr_lengths As Boolean
    Dim i As Long

    equal_arr_lengths = False
    input_id_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
    input_sort_num_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_sort_num_arr)
    input_filter_code_arr_odadt = CREATE_STRING_ONE_DIM_ARRAYDIMSTYPE(input_filter_code_arr)
    input_in_y_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_arr)
    input_out_f_i_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_arr)
    input_in_y_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_p_one_arr)
    input_out_f_i_p_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_p_one_arr)
    input_in_y_i_p_two_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_p_two_arr)
    input_out_f_i_p_two_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_p_two_arr)
    input_in_y_i_m_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_m_one_arr)
    input_out_f_i_m_one_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_m_one_arr)
    input_in_y_i_m_two_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_in_y_i_m_two_arr)
    input_out_f_i_m_two_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_out_f_i_m_two_arr)

    If ((input_id_arr_odadt.length = input_sort_num_arr_odadt.length) And (input_id_arr_odadt.length = input_filter_code_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_arr_odadt.length) And _
    (input_id_arr_odadt.length = input_in_y_i_p_one_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_p_one_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_p_two_arr_odadt.length) And _
    (input_id_arr_odadt.length = input_out_f_i_p_two_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_m_one_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_m_one_arr_odadt.length) And (input_id_arr_odadt.length = input_in_y_i_m_two_arr_odadt.length) And (input_id_arr_odadt.length = input_out_f_i_m_two_arr_odadt.length)) Then
        
        equal_arr_lengths = True
    Else
        equal_arr_lengths = False
    End If

    If (equal_arr_lengths = True) Then
        result_arr_odadt = CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_id_arr)
        ReDim result(result_arr_odadt.l_bnd To result_arr_odadt.u_bnd)

        For i = result_arr_odadt.l_bnd To result_arr_odadt.u_bnd
		    result(i) = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_V000(input_id_arr(i), input_sort_num_arr(i), input_filter_code_arr(i), input_in_y_i_arr(i), input_out_f_i_arr(i), input_in_y_i_p_one_arr(i), input_out_f_i_p_one_arr(i), input_in_y_i_p_two_arr(i), _
    input_out_f_i_p_two_arr(i), input_in_y_i_m_one_arr(i), input_out_f_i_m_one_arr(i), input_in_y_i_m_two_arr(i), input_out_f_i_m_two_arr(i))
	    Next
    Else
        ' nullify type
    End If

    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_type_arr() As DBFuncDataFYIPMTwoDoubleType) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_type_arr, 1)
	result.u_bnd = UBound(input_type_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000 = result
End Function

Public Function FILTER_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_filter_val As String, input_type_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType()
    Dim result() As DBFuncDataFYIPMTwoDoubleType
    Dim input_type_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim filtered_arr_ctr As Long

    input_type_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_type_arr)

    i = UNASSIGNED_LONG_VAL
    filtered_arr_ctr = -1

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        If (lookup_filter_val = input_type_arr(i).filter_code) Then
            filtered_arr_ctr = filtered_arr_ctr + 1
        Else
            '   do nothing
        End If
    Next

    If (filtered_arr_ctr >= 0) Then

        ReDim result(0 To filtered_arr_ctr)

        filtered_arr_ctr = 0

        For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
            If (lookup_filter_val = input_type_arr(i).filter_code) Then
                result(filtered_arr_ctr) = input_type_arr(i)
                filtered_arr_ctr = filtered_arr_ctr + 1
            Else
                '   do nothing
            End If
        Next
    Else
        ReDim result(0 To 0)
        result(0) = NULLIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000()
    End If

    FILTER_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

'   data must be filtered before this function is used!
'   Sorts the array elements via the sort number value in ascending order
Public Function SORT_BY_SN_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(input_type_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType()
    Dim result() As DBFuncDataFYIPMTwoDoubleType
    Dim input_type_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim sort_num_lookup As Long
    Dim sort_num_match_ctr As Long

    input_type_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_type_arr)

    ReDim result(0 To (input_type_arr_odadt.length - 1))
    i = UNASSIGNED_LONG_VAL
    sort_num_lookup = UNASSIGNED_LONG_VAL
    sort_num_match_ctr = 0

    For sort_num_lookup = (1) To (input_type_arr_odadt.length)
        For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
            If ((sort_num_lookup = input_type_arr(i).sort_num) And (sort_num_match_ctr = 0)) Then
                result(sort_num_lookup - 1) = input_type_arr(i)
                sort_num_match_ctr = sort_num_match_ctr + 1
            Else
                '   do nothing
            End If
        Next
        sort_num_match_ctr = 0
    Next

    SORT_BY_SN_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function SF_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_filter_val As String, input_type_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType()
    '   Dim result() As DBFuncDataFYIPOneDoubleType

    '   SF_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
    SF_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = SORT_BY_SN_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(FILTER_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_filter_val, input_type_arr))
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_V000(lookup_filter_val As String, input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double, _
    input_in_y_i_p_two_arr() As Double, input_out_f_i_p_two_arr() As Double, input_in_y_i_m_one_arr() As Double, input_out_f_i_m_one_arr() As Double, input_in_y_i_m_two_arr() As Double, input_out_f_i_m_two_arr() As Double)  As DBFuncDataFYIPMTwoDoubleType()
    '    raw_arr() As DBFuncDataFYIPMTwoDoubleType
    '   Dim result() As DBFuncDataFYIPOneDoubleType

    '   SF_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_V000 = SORT_BY_SN_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(FILTER_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_filter_val, CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(input_id_arr, input_sort_num_arr, input_filter_code_arr, input_in_y_i_arr, input_out_f_i_arr, input_in_y_i_p_one_arr, _
        input_out_f_i_p_one_arr, input_in_y_i_p_two_arr, input_out_f_i_p_two_arr, input_in_y_i_m_one_arr, input_out_f_i_m_one_arr, input_in_y_i_m_two_arr, input_out_f_i_m_two_arr)))
End Function

Public Function RETURN_MIN_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(input_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType
	Dim result As DBFuncDataFYIPMTwoDoubleType
	Dim input_arr_odadt As ArrayDimensionsType
	Dim min_rec As DBFuncDataFYIPMTwoDoubleType
	Dim i As Long

	input_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_arr)
	min_rec = input_arr(input_arr_odadt.l_bnd)
	For i = (input_arr_odadt.l_bnd + 1) To (input_arr_odadt.u_bnd)
		If (input_arr(i).data_i.in_y < min_rec.data_i.in_y) Then
			min_rec = input_arr(i)
		End If
	Next
	result = min_rec
	RETURN_MIN_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00 = result
End Function

Public Function RETURN_MAX_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(input_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType
	Dim result As DBFuncDataFYIPMTwoDoubleType
	Dim input_arr_odadt As ArrayDimensionsType
	Dim max_rec As DBFuncDataFYIPMTwoDoubleType
	Dim i As Long

	input_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_arr)
	max_rec = input_arr(input_arr_odadt.l_bnd)
	For i = (input_arr_odadt.l_bnd + 1) To (input_arr_odadt.u_bnd)
		If (input_arr(i).data_i.in_y > max_rec.data_i.in_y) Then
			max_rec = input_arr(i)
		End If
	Next
	result = max_rec
	RETURN_MAX_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00 = result
End Function

'   returns the last match found when going through the record array
Public Function RETURN_LAST_MATCH_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(lookup_val As Double, input_arr() As DBFuncDataFYIPMTwoDoubleType) As DBFuncDataFYIPMTwoDoubleType
	Dim result As DBFuncDataFYIPMTwoDoubleType
	Dim input_arr_odadt As ArrayDimensionsType
	Dim match_rec As DBFuncDataFYIPMTwoDoubleType
	Dim i As Long

	input_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_arr)
	match_rec = NULLIFY_DBFUNCDATAFYIPMTWODOUBLETYPE_V000()
	For i = (input_arr_odadt.l_bnd) To (input_arr_odadt.u_bnd)
		If (input_arr(i).data_i.in_y = lookup_val) Then
			match_rec = input_arr(i)
		End If
	Next
	result = match_rec
	RETURN_LAST_MATCH_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00 = result
End Function

Public Function EXTRACT_IN_Y_I_ARR_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(input_type_arr() As DBFuncDataFYIPMTwoDoubleType) As Double()
    Dim result() As Double
    Dim input_type_arr_odadt As ArrayDimensionsType
    Dim i As Long

    input_type_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(input_type_arr)

    ReDim result((input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd))

    i = UNASSIGNED_LONG_VAL

    For i = (input_type_arr_odadt.l_bnd) To (input_type_arr_odadt.u_bnd)
        result(i) = input_type_arr(i).data_i.in_y
    Next

    EXTRACT_IN_Y_I_ARR_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function CREATE_EDARTYPE_V000(in_within_rd_dafc As DataAnalysisFuncCondEnum, in_above_rd_dafc As DataAnalysisFuncCondEnum, in_below_rd_dafc As DataAnalysisFuncCondEnum, in_equal_to_rd_dafc As DataAnalysisFuncCondEnum, in_multi_match_to_rd_dafc As DataAnalysisFuncCondEnum) As EntityDataAnalysisRulesType
    Dim result As EntityDataAnalysisRulesType
    result.within_rd_dafc = in_within_rd_dafc
    result.above_rd_dafc = in_above_rd_dafc
    result.below_rd_dafc = in_below_rd_dafc
    result.equal_to_rd_dafc = in_equal_to_rd_dafc
    result.multi_match_to_rd_dafc =  in_multi_match_to_rd_dafc

    CREATE_EDARTYPE_V000 = result
End Function

Public Function CREATE_EDARTYPE_VIA_DAFC_STRINGS_V000(in_within_rd_dafc_str As String, in_above_rd_str As String, in_below_rd_str As String, in_equal_to_rd_str As String, in_multi_match_to_rd_str As String) As EntityDataAnalysisRulesType
    CREATE_EDARTYPE_VIA_DAFC_STRINGS_V000 = CREATE_EDARTYPE_V000(ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(in_within_rd_dafc_str), ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(in_above_rd_str), ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(in_below_rd_str), ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(in_equal_to_rd_str), ASSIGN_DATAANALYSISFUNCCOND_ENUM_FROM_STRING(in_multi_match_to_rd_str))
End Function

'   General Calc

Public Function RETURN_Y_LINEAR_FUNCTION(gradient As Double, x_coordinate As Double, y_intercept As Double) As Double
    ' y = m * x + c
    RETURN_Y_LINEAR_FUNCTION = gradient * x_coordinate + y_intercept
End Function

Public Function CALC_F_X_GRADIENT_V000(f_i_p_one As Double, f_i As Double, x_i_p_one As Double, x_i As Double) As Double
    Dim result As Double
    result = ((f_i_p_one - f_i) / (x_i_p_one - x_i))
    CALC_F_X_GRADIENT_V000 = result
End Function

Public Function CUSTOM_LINEAR_INTERPOLATION(y_0 As Double, y_2 As Double, x_0 As Double, x_1 As Double, x_2 As Double)
    ' returns y_1
    ' CUSTM_LNR_INTERPOLATION = ((y_2 - y_0) / (x_2 - x_0)) * (x_1 - x_0) + y_0
    CUSTOM_LINEAR_INTERPOLATION = RETURN_Y_LINEAR_FUNCTION(((y_2 - y_0) / (x_2 - x_0)), (x_1 - x_0), y_0)
End Function

Public Function CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION(y_data_point_i_plus_one As Double, y_data_point_i_minus_one As Double, x_data_point_i_plus_one As Double, x_data_point_i As Double, x_data_point_i_minus_one As Double) As Double
'    CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION = (((y_data_point_i_plus_one - y_data_point_i_minus_one) / (x_data_point_i_plus_one - x_data_point_i_minus_one)) * (x_data_point_i - x_data_point_i_minus_one)) + y_data_point_i_minus_one
    CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION = RETURN_Y_LINEAR_FUNCTION(((y_data_point_i_plus_one - y_data_point_i_minus_one) / (x_data_point_i_plus_one - x_data_point_i_minus_one)), (x_data_point_i - x_data_point_i_minus_one), y_data_point_i_minus_one)
End Function

Public Function CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_BELOW_X_Y_DATA_SET(y_data_point_i_plus_two As Double, y_data_point_i_plus_one As Double, x_data_point_i_plus_two As Double, x_data_point_i_plus_one As Double, x_data_point_i As Double) As Double
    Dim result As Double
    Dim local_gradient As Double, local_x_coordinate As Double, local_y_intercept As Double
    
    local_gradient = ((y_data_point_i_plus_two - y_data_point_i_plus_one) / (x_data_point_i_plus_two - x_data_point_i_plus_one))
    local_x_coordinate = x_data_point_i - x_data_point_i_plus_one
    local_y_intercept = y_data_point_i_plus_one
    
    result = RETURN_Y_LINEAR_FUNCTION(local_gradient, local_x_coordinate, local_y_intercept)
    CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_BELOW_X_Y_DATA_SET = result
End Function

Public Function CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_ABOVE_X_Y_DATA_SET(y_data_point_i_minus_one As Double, y_data_point_i_minus_two As Double, x_data_point_i As Double, x_data_point_i_minus_one As Double, x_data_point_i_minus_two As Double) As Double
    Dim result As Double
    Dim local_gradient As Double, local_x_coordinate As Double, local_y_intercept As Double
    
    local_gradient = ((y_data_point_i_minus_one - y_data_point_i_minus_two) / (x_data_point_i_minus_one - x_data_point_i_minus_two))
    local_x_coordinate = x_data_point_i - x_data_point_i_minus_two
    local_y_intercept = y_data_point_i_minus_two
    
    result = RETURN_Y_LINEAR_FUNCTION(local_gradient, local_x_coordinate, local_y_intercept)
    CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_ABOVE_X_Y_DATA_SET = result
End Function

'   only checks if lookup value against min and max values in ref data
Public Function ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000(lookup_val As Double, rd_min_val As Double, rd_max_val As Double) As RefDataLinearIntervalChk
    Dim result As RefDataLinearIntervalChk

    If ((lookup_val < rd_min_val) And (rd_min_val < rd_max_val)) Then
        result = BELOW_REF_DATA
    ElseIf ((lookup_val > rd_max_val) And (rd_min_val < rd_max_val)) Then
        result = ABOVE_REF_DATA
    ElseIf ((lookup_val > rd_min_val) And (lookup_val < rd_max_val) And (rd_min_val < rd_max_val)) Then
        result = WITHIN_REF_DATA
    Else
        result = ERROR_INTERVAL_CHK
    End If

    ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000 = result
End Function

Public Function ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_DOUBLE_ARR_V000(lookup_val As Double, ref_data_val_arr() As Double) As RefDataLinearIntervalChk
    Dim process As RefDataLinearIntervalChk
    Dim result As RefDataLinearIntervalChk
    Dim ref_data_val_arr_odadt As ArrayDimensionsType
    Dim min_ref_data_val As Double
    Dim max_ref_data_val As Double
    Dim close_ref_data_below As Double
    Dim close_ref_data_above As Double
    Dim i As Long
    Dim rd_match_ctr As Long
    
    process = UNASSIGNED_RDLIC
    result = UNASSIGNED_RDLIC

    close_ref_data_below = UNASSIGNED_DOUBLE_VAL
    close_ref_data_above = UNASSIGNED_DOUBLE_VAL

    i = UNASSIGNED_LONG_VAL
    rd_match_ctr = 0

    ref_data_val_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(ref_data_val_arr)

    min_ref_data_val = RETURN_MIN_VAL_FROM_DOUBLE_OD_ARR_V00(ref_data_val_arr)

    max_ref_data_val = RETURN_MAX_VAL_FROM_DOUBLE_OD_ARR_V00(ref_data_val_arr)

    process = ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000(lookup_val, min_ref_data_val, max_ref_data_val)

    If (process = WITHIN_REF_DATA) Then
        For i = (ref_data_val_arr_odadt.l_bnd) To (ref_data_val_arr_odadt.u_bnd)
            If (lookup_val = ref_data_val_arr(i)) Then
                rd_match_ctr = rd_match_ctr + 1
            Else
                '   do nothing
            End If
        Next

        If (rd_match_ctr = 1) Then
            result = EQUAL_TO_REF_DATA
        ElseIf (rd_match_ctr = 0) Then
            '   no matches were found, lookup val is still within ref data
            result = process
        ElseIf (rd_match_ctr > 1) Then
            result = MULTI_MATCH_TO_REF_DATA
        Else
            result = ERROR_INTERVAL_CHK
        End If
    Else
        '   no need to check for matches, assign process value to result
        result = process
    End If

    ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_DOUBLE_ARR_V000 = result
End Function

Public Function ASSIGN_REFDATALINEARINTERVALCHK_ENUM_VIA_DOUBLE_ARR_BOUNDARY_VALS_V000(lookup_val As Double, rd_val_arr() As Double, min_rd_val As Double, max_rd_val As Double) As RefDataLinearIntervalChk
    Dim process As RefDataLinearIntervalChk
    Dim result As RefDataLinearIntervalChk
    Dim ref_data_val_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim rd_match_ctr As Long
    
    process = UNASSIGNED_RDLIC
    result = UNASSIGNED_RDLIC

    i = UNASSIGNED_LONG_VAL
    rd_match_ctr = 0

    ref_data_val_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(rd_val_arr)

    process = ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000(lookup_val, min_rd_val, max_rd_val)

    If (process = WITHIN_REF_DATA) Then
        For i = (ref_data_val_arr_odadt.l_bnd) To (ref_data_val_arr_odadt.u_bnd)
            If (lookup_val = rd_val_arr(i)) Then
                rd_match_ctr = rd_match_ctr + 1
            Else
                '   do nothing
            End If
        Next

        If (rd_match_ctr = 1) Then
            result = EQUAL_TO_REF_DATA
        ElseIf (rd_match_ctr = 0) Then
            '   no matches were found, lookup val is still within ref data
            result = process
        ElseIf (rd_match_ctr > 1) Then
            result = MULTI_MATCH_TO_REF_DATA
        Else
            result = ERROR_INTERVAL_CHK
        End If
    Else
        '   no need to check for matches, assign process value to result
        result = process
    End If

    ASSIGN_REFDATALINEARINTERVALCHK_ENUM_VIA_DOUBLE_ARR_BOUNDARY_VALS_V000 = result
End Function

' Public Function ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType) As RefDataLinearIntervalChk
'     Dim process As RefDataLinearIntervalChk
'     Dim result As RefDataLinearIntervalChk
'     Dim ref_data_val_arr_odadt As ArrayDimensionsType
'     Dim min_ref_data_val As Double
'     Dim max_ref_data_val As Double
'     Dim close_ref_data_below As Double
'     Dim close_ref_data_above As Double
'     Dim i As Long
'     Dim rd_match_ctr As Long
    
'     process = UNASSIGNED_RDLIC
'     result = UNASSIGNED_RDLIC

'     close_ref_data_below = UNASSIGNED_DOUBLE_VAL
'     close_ref_data_above = UNASSIGNED_DOUBLE_VAL

'     i = UNASSIGNED_LONG_VAL
'     rd_match_ctr = 0

'     ref_data_val_arr_odadt = CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(in_rd_arr)

'     min_ref_data_val = RETURN_MIN_VAL_FROM_DOUBLE_OD_ARR_V00(in_rd_arr)

'     max_ref_data_val = RETURN_MAX_VAL_FROM_DOUBLE_OD_ARR_V00(in_rd_arr)

'     process = ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000(lookup_val, min_ref_data_val, max_ref_data_val)

'     If (process = WITHIN_REF_DATA) Then
'         For i = (ref_data_val_arr_odadt.l_bnd) To (ref_data_val_arr_odadt.u_bnd)
'             If (lookup_val = in_rd_arr.data_i.in_y(i)) Then....
'                 rd_match_ctr = rd_match_ctr + 1
'             Else
'                 '   do nothing
'             End If
'         Next

'         If (rd_match_ctr = 1) Then
'             result = EQUAL_TO_REF_DATA
'         ElseIf (rd_match_ctr = 0) Then
'             '   no matches were found, lookup val is still within ref data
'             result = process
'         ElseIf (rd_match_ctr > 1) Then
'             result = MULTI_MATCH_TO_REF_DATA
'         Else
'             result = ERROR_INTERVAL_CHK
'         End If
'     Else
'         '   no need to check for matches, assign process value to result
'         result = process
'     End If

'     ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
' End Function

Public Function CALC_Y_I_VIA_LINEAR_ESTIMATION_OF_X_Y_DATA_SET(y_data_point_i_plus_two As Double, y_data_point_i_plus_one As Double, y_data_point_i_minus_one As Double, y_data_point_i_minus_two As Double, x_data_point_i_plus_two As Double, x_data_point_i_plus_one As Double, x_data_point_i As Double, x_data_point_i_minus_one As Double, x_data_point_i_minus_two As Double, input_ref_data_linear_interval_chk As RefDataLinearIntervalChk) As Double
    Dim result As Double
    Select Case input_ref_data_linear_interval_chk
        Case WITHIN_REF_DATA
            result = CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION(y_data_point_i_plus_one, y_data_point_i_minus_one, x_data_point_i_plus_one, x_data_point_i, x_data_point_i_minus_one)
        Case ABOVE_REF_DATA
            result = CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_ABOVE_X_Y_DATA_SET(y_data_point_i_minus_one, y_data_point_i_minus_two, x_data_point_i, x_data_point_i_minus_one, x_data_point_i_minus_two)
        Case BELOW_REF_DATA
            result = CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_BELOW_X_Y_DATA_SET(y_data_point_i_plus_two, y_data_point_i_plus_one, x_data_point_i_plus_two, x_data_point_i_plus_one, x_data_point_i)
        Case EQUAL_TO_REF_DATA
            ' need to double check this is the right value being passed in
            result = y_data_point_i_minus_one
        Case MULTI_MATCH_TO_REF_DATA
            result = ERROR_DOUBLE_VAL
        Case NULL_INTERVAL_CHK
            result = NULL_DOUBLE_VAL
        Case ERROR_INTERVAL_CHK
            result = ERROR_DOUBLE_VAL
        Case Else
            result = ERROR_DOUBLE_VAL
    End Select
    CALC_Y_I_VIA_LINEAR_ESTIMATION_OF_X_Y_DATA_SET = result
End Function

Public Function RETURN_DATA_ANALYSIS_ID_FROM_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000(lookup_val As Double, input_data() As DBFuncDataFYIPOneDoubleType, rdlic As RefDataLinearIntervalChk) As Long
    Dim result As Long
    Dim input_data_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim in_y_i_p_one_found As Double, sort_num_found As Double
    Dim first_rec_found As Boolean
    Dim in_y_arr() As Double
    Dim min_max_in_y_found As Double

    input_data_arr_odadt = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_data)

    Select Case rdlic
        Case WITHIN_REF_DATA
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
                If ((input_data(i).data_i_p_one.in_y > lookup_val) And (first_rec_found = False)) Then
                    in_y_i_p_one_found = input_data(i).data_i_p_one.in_y
                    sort_num_found = input_data(i).sort_num
                    result = input_data(i).id
                    first_rec_found = True
                ElseIf ((input_data(i).data_i_p_one.in_y > lookup_val) And (first_rec_found = True) And (input_data(i).data_i_p_one.in_y <= in_y_i_p_one_found) And (input_data(i).sort_num < sort_num_found)) Then
                    in_y_i_p_one_found = input_data(i).data_i_p_one.in_y
                    sort_num_found = input_data(i).sort_num
                    result = input_data(i).id
                End If
            Next
        Case ABOVE_REF_DATA
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
               in_y_arr(i) = input_data(i).data_i_p_one.in_y
            Next
            min_max_in_y_found = RETURN_MAX_VAL_FROM_DOUBLE_OD_ARR_V00(in_y_arr)
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
                If (min_max_in_y_found = input_data(i).data_i.in_y) Then
                    result = input_data(i).id
                Else
                    '   do nothing
                End If
            Next
        Case BELOW_REF_DATA
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
               in_y_arr(i) = input_data(i).data_i_p_one.in_y
            Next
            min_max_in_y_found = RETURN_MIN_VAL_FROM_DOUBLE_OD_ARR_V00(in_y_arr)
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
                If (min_max_in_y_found = input_data(i).data_i.in_y) Then
                    result = input_data(i).id
                Else
                    '   do nothing
                End If
            Next
        Case EQUAL_TO_REF_DATA
            For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
                If (lookup_val = input_data(i).data_i.in_y) Then
                    result = input_data(i).id
                Else
                    '   do nothing
                End If
            Next
        Case MULTI_MATCH_TO_REF_DATA
            result = ERROR_LONG_VAL
        Case NULL_INTERVAL_CHK
            result = NULL_LONG_VAL
        Case ERROR_INTERVAL_CHK
            result = ERROR_LONG_VAL
        Case Else
            result = ERROR_LONG_VAL
    End Select
    
    RETURN_DATA_ANALYSIS_ID_FROM_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000 = result
End Function

Public Function HANDLE_WITHIN_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType, within_rd_dafc As DataAnalysisFuncCondEnum) As Double
    Dim result As Double
    Dim result_rec As DBFuncDataFYIPMTwoDoubleType
    Dim input_data_arr_odadt As ArrayDimensionsType
    Dim i As Long
    Dim in_y_i_p_one_found As Double, sort_num_found As Double
    Dim first_rec_found As Boolean
    
    input_data_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(in_rd_arr)

    If (within_rd_dafc = LINEAR_INTERPOLATION_VIA_CLOSE_DATA) Then
        For i = (input_data_arr_odadt.l_bnd) To (input_data_arr_odadt.u_bnd)
            If ((in_rd_arr(i).data_i_p_one.in_y > lookup_val) And (first_rec_found = False)) Then
                in_y_i_p_one_found = in_rd_arr(i).data_i_p_one.in_y
                sort_num_found = in_rd_arr(i).sort_num
                result_rec = in_rd_arr(i)
                first_rec_found = True
            ElseIf ((in_rd_arr(i).data_i_p_one.in_y > lookup_val) And (first_rec_found = True) And (in_rd_arr(i).data_i_p_one.in_y <= in_y_i_p_one_found) And (in_rd_arr(i).sort_num < sort_num_found)) Then
                in_y_i_p_one_found = in_rd_arr(i).data_i_p_one.in_y
                sort_num_found = in_rd_arr(i).sort_num
                result_rec = in_rd_arr(i)
            End If
        Next

        result = CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION(result_rec.data_i_p_one.out_f, result_rec.data_i.out_f, result_rec.data_i_p_one.in_y, lookup_val, result_rec.data_i.in_y)

    Else
        ' do nothing yet?
        result = NULL_DOUBLE_VAL
    End If

    HANDLE_WITHIN_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function HANDLE_BELOW_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType, min_rd_rec As DBFuncDataFYIPMTwoDoubleType, below_rd_dafc As DataAnalysisFuncCondEnum) As Double
    Dim result As Double
    Dim input_data_arr_odadt As ArrayDimensionsType
    
    input_data_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(in_rd_arr)

    If (below_rd_dafc = LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA) Then
        result = CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_BELOW_X_Y_DATA_SET(min_rd_rec.data_i_p_two.out_f, min_rd_rec.data_i_p_one.out_f, min_rd_rec.data_i_p_two.in_y, min_rd_rec.data_i_p_one.in_y, lookup_val)
    Else
        ' do nothing yet?
        result = NULL_DOUBLE_VAL
    End If

    HANDLE_BELOW_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function HANDLE_ABOVE_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType, max_rd_rec As DBFuncDataFYIPMTwoDoubleType, above_rd_dafc As DataAnalysisFuncCondEnum) As Double
    Dim result As Double
    Dim input_data_arr_odadt As ArrayDimensionsType
    
    input_data_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(in_rd_arr)

    If (above_rd_dafc = LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA) Then
        result = CALC_Y_I_VIA_LINEAR_EXTRAPOLATION_ABOVE_X_Y_DATA_SET(max_rd_rec.data_i_m_one.out_f, max_rd_rec.data_i_m_two.out_f, lookup_val, max_rd_rec.data_i_m_one.in_y, max_rd_rec.data_i_m_two.in_y)
    Else
        ' do nothing yet?
        result = NULL_DOUBLE_VAL
    End If

    HANDLE_ABOVE_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function HANDLE_EQUAL_TO_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType, equal_to_rd_dafc As DataAnalysisFuncCondEnum) As Double
    Dim result As Double
    Dim result_rec As DBFuncDataFYIPMTwoDoubleType
    Dim input_data_arr_odadt As ArrayDimensionsType
    
    input_data_arr_odadt = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE_V000(in_rd_arr)

    If (equal_to_rd_dafc = DIRECT_ASSIGNMENT) Then
        result_rec = RETURN_LAST_MATCH_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(lookup_val, in_rd_arr)
        result = result_rec.data_i.out_f
    Else
        ' do nothing yet?
        result = NULL_DOUBLE_VAL
    End If

    HANDLE_EQUAL_TO_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val As Double, in_rd_arr() As DBFuncDataFYIPMTwoDoubleType, edar As EntityDataAnalysisRulesType) As Double
    Dim result As Double
    Dim in_y_i_arr() As Double
    Dim min_rd_rec As DBFuncDataFYIPMTwoDoubleType
    Dim max_rd_rec As DBFuncDataFYIPMTwoDoubleType
    Dim rdlic As RefDataLinearIntervalChk

    min_rd_rec = RETURN_MIN_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(in_rd_arr)
    max_rd_rec = RETURN_MAX_IN_Y_I_REC_VIA_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V00(in_rd_arr)
    in_y_i_arr = EXTRACT_IN_Y_I_ARR_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(in_rd_arr)
    rdlic = ASSIGN_REFDATALINEARINTERVALCHK_ENUM_VIA_DOUBLE_ARR_BOUNDARY_VALS_V000(lookup_val, in_y_i_arr, min_rd_rec.data_i.in_y, max_rd_rec.data_i.in_y)

    Erase in_y_i_arr

    Select Case rdlic
        Case WITHIN_REF_DATA
            result = HANDLE_WITHIN_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val, in_rd_arr, edar.within_rd_dafc)
        Case ABOVE_REF_DATA
            result = HANDLE_ABOVE_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val, in_rd_arr, max_rd_rec, edar.above_rd_dafc)
        Case BELOW_REF_DATA
            result = HANDLE_BELOW_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val, in_rd_arr, min_rd_rec, edar.below_rd_dafc)
        Case EQUAL_TO_REF_DATA
            result = HANDLE_EQUAL_TO_RD_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(lookup_val, in_rd_arr, edar.equal_to_rd_dafc)
        Case MULTI_MATCH_TO_REF_DATA
            result = ERROR_DOUBLE_VAL
        Case NULL_INTERVAL_CHK
            result = ERROR_DOUBLE_VAL
        Case ERROR_INTERVAL_CHK
            result = ERROR_DOUBLE_VAL
        Case Else
            result = ERROR_DOUBLE_VAL
    End Select
    
    RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000 = result
End Function

Public Function SOLVE_SIMUL_FUNC_FY_DOUB_FOR_Y_V000() As Double
    Dim result As Double

    SOLVE_SIMUL_FUNC_FY_DOUB_FOR_Y_V000 = result
End Function

Public Function CONVERT_VALUE_ANGLE_DEG_TO_RAD_V001(input_angle_deg As Double) As Double
    CONVERT_VALUE_ANGLE_DEG_TO_RAD_V001 = input_angle_deg * (PI_VBA_DOUBLE / 180)
End Function

Public Function CONVERT_VALUE_ANGLE_RAD_TO_DEG_V001(input_angle_rad As Double) As Double
    CONVERT_VALUE_ANGLE_RAD_TO_DEG_V001 = input_angle_rad * (180 / PI_VBA_DOUBLE)
End Function