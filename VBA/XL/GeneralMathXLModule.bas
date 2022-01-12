Attribute VB_Name = "GeneralMathXLModule"
'
'	@file GeneralMathXLModule.bas
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
'		Eriez Magnetics Pty Ltd., Sales & R&D Department
'		21 Shirley Way
'		Epping, Victoria, Australia 3076
'	Author E-Mail: lsponza@eriez.com
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
'	
'
'
'
'	----------------------------------------------------------------------------
'
'	Requires:
'		- GeneralCalcVBAModule
'		- GeneralXLModule


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

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_FROM_RANGES_V000(input_id_col As Range, input_sort_num_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim input_id_arr() As Long, input_sort_num_arr() As Long, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double

    input_id_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_sort_num_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")

    result = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR(input_id_arr, input_sort_num_arr, input_in_y_i_arr, input_out_f_i_arr, input_in_y_i_p_one_arr, input_out_f_i_p_one_arr)

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_FROM_RANGES_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_FROM_RANGES_V001(input_id_col As Range, input_sort_num_col As Range, input_filter_code_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range) As DBFuncDataFYIPOneDoubleType()
    Dim result() As DBFuncDataFYIPOneDoubleType
    Dim input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double

    input_id_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_sort_num_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_filter_code_arr = CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_filter_code_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")

    result = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V001(input_id_arr, input_sort_num_arr, input_filter_code_arr, input_in_y_i_arr, input_out_f_i_arr, input_in_y_i_p_one_arr, input_out_f_i_p_one_arr)

    CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_FROM_RANGES_V001 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_FROM_RANGES_V000(input_id_col As Range, input_sort_num_col As Range, input_filter_code_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range, _
    input_in_y_i_p_two_col As Range, input_out_f_i_p_two_col As Range, input_in_y_i_m_one_col As Range, input_out_f_i_m_one_col As Range, input_in_y_i_m_two_col As Range, input_out_f_i_m_two_col As Range) As DBFuncDataFYIPMTwoDoubleType()
    
    Dim result() As DBFuncDataFYIPMTwoDoubleType
    Dim input_id_arr() As Long, input_sort_num_arr() As Long, input_filter_code_arr() As String, input_in_y_i_arr() As Double, input_out_f_i_arr() As Double, input_in_y_i_p_one_arr() As Double, input_out_f_i_p_one_arr() As Double, input_in_y_i_p_two_arr() As Double, _
        input_out_f_i_p_two_arr() As Double, input_in_y_i_m_one_arr() As Double, input_out_f_i_m_one_arr() As Double, input_in_y_i_m_two_arr() As Double, input_out_f_i_m_two_arr() As Double

    input_id_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_sort_num_arr = CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_filter_code_arr = CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_filter_code_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_p_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_p_two_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_p_two_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_m_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_m_one_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_in_y_i_m_two_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST")
    input_out_f_i_m_two_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST")

    result = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(input_id_arr, input_sort_num_arr, input_filter_code_arr, input_in_y_i_arr, input_out_f_i_arr, input_in_y_i_p_one_arr, input_out_f_i_p_one_arr, input_in_y_i_p_two_arr, _
        input_out_f_i_p_two_arr, input_in_y_i_m_one_arr, input_out_f_i_m_one_arr, input_in_y_i_m_two_arr, input_out_f_i_m_two_arr)

    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_FROM_RANGES_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_FROM_RANGES_V001(input_id_col As Range, input_sort_num_col As Range, input_filter_code_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range, _
    input_in_y_i_p_two_col As Range, input_out_f_i_p_two_col As Range, input_in_y_i_m_one_col As Range, input_out_f_i_m_one_col As Range, input_in_y_i_m_two_col As Range, input_out_f_i_m_two_col As Range) As DBFuncDataFYIPMTwoDoubleType()

    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_FROM_RANGES_V001 = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_filter_code_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"))
End Function

Public Function CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_FROM_RANGES_V000(lookup_filter_val As Variant, input_id_col As Range, input_sort_num_col As Range, input_filter_code_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, _
    input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range, input_in_y_i_p_two_col As Range, input_out_f_i_p_two_col As Range, input_in_y_i_m_one_col As Range, input_out_f_i_m_one_col As Range, input_in_y_i_m_two_col As Range, _
    input_out_f_i_m_two_col As Range) As DBFuncDataFYIPMTwoDoubleType()
    
    CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_FROM_RANGES_V000 = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_V000(CAST_CELL_VALUE_TO_STRING(lookup_filter_val), CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_filter_code_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"))
End Function

Public Function ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_STRING_FROM_BOUNDARY_VALS_V000(lookup_val As Variant, rd_min_val As Variant, rd_max_val As Variant) As String
    Dim result As String
    Dim result_rdlic As RefDataLinearIntervalChk

    result_rdlic = ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_ENUM_FROM_BOUNDARY_VALS_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), CAST_CELL_VALUE_TO_DOUBLE(rd_min_val), CAST_CELL_VALUE_TO_DOUBLE(rd_max_val))
    result = ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM(result_rdlic)

    ASSIGN_BOUNDARY_REFDATALINEARINTERVALCHK_STRING_FROM_BOUNDARY_VALS_V000 = result
End Function

Public Function ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_DOUBLE_RANGE_V000(lookup_val As Variant, rd_range As Range) As String
    Dim result As String
    Dim double_arr() As Double
    Dim result_rdlic As RefDataLinearIntervalChk

    double_arr = CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(rd_range,"DOWNWARDS_ALONG_COLS_FIRST")
    result_rdlic = ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_DOUBLE_ARR_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), double_arr)
    result = ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM(result_rdlic)

    ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_DOUBLE_RANGE_V000 = result
End Function

Public Function RETURN_DATA_ANALYSIS_ID_FROM_DBFUNCDATAFYIPONEDOUBLETYPE_RANGE_V000(lookup_val As Variant, input_id_col As Range, input_sort_num_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range, input_rdlic_str As String) As Long
    Dim result As Long
    Dim input_rdlic As RefDataLinearIntervalChk
    Dim in_rd_arr() As DBFuncDataFYIPOneDoubleType

    input_rdlic = ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING_V000(input_rdlic_str)
    in_rd_arr = CREATE_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_FROM_RANGES_V000(input_id_col, input_sort_num_col, input_in_y_i_col, input_out_f_i_col, input_in_y_i_p_one_col, input_out_f_i_p_one_col)
    result = RETURN_DATA_ANALYSIS_ID_FROM_DBFUNCDATAFYIPONEDOUBLETYPE_ARR_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), in_rd_arr, input_rdlic)

    RETURN_DATA_ANALYSIS_ID_FROM_DBFUNCDATAFYIPONEDOUBLETYPE_RANGE_V000 = result
End Function

Public Function CREATE_EDARTYPE_VIA_CELL_VALS_V000(within_rd_dafc_cell_val As Variant, above_rd_cell_val As Variant, below_rd_cell_val As Variant, equal_to_rd_cell_val As Variant, multi_match_to_rd_cell_val As Variant) As EntityDataAnalysisRulesType
    CREATE_EDARTYPE_VIA_CELL_VALS_V000 = CREATE_EDARTYPE_VIA_DAFC_STRINGS_V000(CAST_CELL_VALUE_TO_STRING(within_rd_dafc_cell_val), CAST_CELL_VALUE_TO_STRING(above_rd_cell_val), CAST_CELL_VALUE_TO_STRING(below_rd_cell_val), CAST_CELL_VALUE_TO_STRING(equal_to_rd_cell_val), CAST_CELL_VALUE_TO_STRING(multi_match_to_rd_cell_val))
End Function

'   ...
Public Function CALC_FUNC_F_I_VIA_LINEAR_ESTIMATION_OF_FY_DATA_V000(out_f_i_p_two As Variant, out_f_i_p_one As Variant, out_f_i_m_one As Variant, out_f_i_m_two As Variant, in_y_i_p_two As Variant, in_y_i_p_one As Variant, in_y_i As Variant, in_y_i_m_one As Variant, in_y_i_m_two As Variant, in_rdlic_str As String) As Double
	Dim result As Double
    Dim local_rdlic As RefDataLinearIntervalChk
	
	local_rdlic = ASSIGN_REFDATALINEARINTERVALCHK_ENUM_FROM_STRING_V000(in_rdlic_str)
	
    '   condition to prevent errors, is it necessary?
	If (in_y_i > 0) Then
'        CALC_MAGNET_FORCE_INDEX_VIA_LINEAR_INTERPOLATION = CALC_Y_DATA_POINT_I_VIA_LINEAR_INTERPOLATION(magnet_fi_i_p_one, magnet_fi_i_m_one, mag_susp_height_i_p_one, mag_susp_height_i, mag_susp_height_i_m_one)
		result = CALC_Y_I_VIA_LINEAR_ESTIMATION_OF_X_Y_DATA_SET(CAST_CELL_VALUE_TO_DOUBLE(out_f_i_p_two), CAST_CELL_VALUE_TO_DOUBLE(out_f_i_p_one), CAST_CELL_VALUE_TO_DOUBLE(out_f_i_m_one), CAST_CELL_VALUE_TO_DOUBLE(out_f_i_m_two), CAST_CELL_VALUE_TO_DOUBLE(in_y_i_p_two), CAST_CELL_VALUE_TO_DOUBLE(in_y_i_p_one), CAST_CELL_VALUE_TO_DOUBLE(in_y_i), CAST_CELL_VALUE_TO_DOUBLE(in_y_i_m_one), CAST_CELL_VALUE_TO_DOUBLE(in_y_i_m_two), local_rdlic)
	Else
		result = ERROR_DOUBLE_VAL
	End If

    CALC_FUNC_F_I_VIA_LINEAR_ESTIMATION_OF_FY_DATA_V000 = result
End Function

Public Function RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_VIA_RANGES_V000(lookup_val As Variant, lookup_filter_val As Variant, input_id_col As Range, input_sort_num_col As Range, input_filter_code_col As Range, input_in_y_i_col As Range, input_out_f_i_col As Range, _
    input_in_y_i_p_one_col As Range, input_out_f_i_p_one_col As Range, input_in_y_i_p_two_col As Range, input_out_f_i_p_two_col As Range, input_in_y_i_m_one_col As Range, input_out_f_i_m_one_col As Range, input_in_y_i_m_two_col As Range, _
    input_out_f_i_m_two_col As Range, Optional ByVal within_rd_dafc_cell_val As Variant = "LINEAR_INTERPOLATION_VIA_CLOSE_DATA", Optional ByVal above_rd_cell_val As Variant = "LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA", Optional ByVal below_rd_cell_val As Variant = "LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA", _
    Optional ByVal equal_to_rd_cell_val As Variant = "DIRECT_ASSIGNMENT", Optional ByVal multi_match_to_rd_cell_val As Variant = "DO_NOT_CALCULATE") As Double

    '   Dim input_type_arr() As DBFuncDataFYIPMTwoDoubleType
    
    '   this method doesn't seem to work in Excel VBA for some reason I couldn't figure out (type mismatch error) ; possibly passing too many instructions at once
    
    ' RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_VIA_RANGES_V000 = RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), (CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_V000(CAST_CELL_VALUE_TO_STRING(lookup_filter_val), CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_id_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    '     CONVERT_RANGE_TO_ONE_DIM_LONG_ARRAY(input_sort_num_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_STRING_ARRAY(input_filter_code_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    '     CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    '     CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_p_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), _
    '     CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_one_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_in_y_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"), CONVERT_RANGE_TO_ONE_DIM_DOUBLE_ARRAY(input_out_f_i_m_two_col,"DOWNWARDS_ALONG_COLS_FIRST"))), _
    '     CREATE_EDARTYPE_VIA_CELL_VALS_V000(within_rd_dafc_cell_val, above_rd_cell_val, below_rd_cell_val, equal_to_rd_cell_val, multi_match_to_rd_cell_val))

    RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_VIA_RANGES_V000 = RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_FROM_RANGES_V000(lookup_filter_val, input_id_col, input_sort_num_col, _
        input_filter_code_col, input_in_y_i_col, input_out_f_i_col, input_in_y_i_p_one_col, input_out_f_i_p_one_col, input_in_y_i_p_two_col, input_out_f_i_p_two_col, input_in_y_i_m_one_col, input_out_f_i_m_one_col, input_in_y_i_m_two_col, input_out_f_i_m_two_col), _
        CREATE_EDARTYPE_VIA_CELL_VALS_V000(within_rd_dafc_cell_val, above_rd_cell_val, below_rd_cell_val, equal_to_rd_cell_val, multi_match_to_rd_cell_val))

    ' input_type_arr = CREATE_DBFUNCDATAFYIPMTWODOUBLETYPE_SF_ARR_FROM_RANGES_V000(lookup_filter_val, input_id_col, input_sort_num_col, _
    '     input_filter_code_col, input_in_y_i_col, input_out_f_i_col, input_in_y_i_p_one_col, input_out_f_i_p_one_col, input_in_y_i_p_two_col, input_out_f_i_p_two_col, input_in_y_i_m_one_col, input_out_f_i_m_one_col, input_in_y_i_m_two_col, input_out_f_i_m_two_col)
    ' RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_VIA_RANGES_V000 = RETURN_DATA_ANALYSIS_F_I_FROM_DBFUNCDATAFYIPMTWODOUBLETYPE_ARR_V000(CAST_CELL_VALUE_TO_DOUBLE(lookup_val), input_type_arr, _
    '     CREATE_EDARTYPE_VIA_CELL_VALS_V000(within_rd_dafc_cell_val, above_rd_cell_val, below_rd_cell_val, equal_to_rd_cell_val, multi_match_to_rd_cell_val))
End Function