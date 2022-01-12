Attribute VB_Name = "GeneralMathModule"
'
'	@file THMModule.bas
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
'	05/11/2021		16:20		Leonard Sponza		THMModule created
'                                                   functions moved to here
'
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
    NULL_INTERVAL_CHK = NULL_LONG_VAL
    ERROR_INTERVAL_CHK = ERROR_LONG_VAL
End Enum

Public Enum DATA_ANALYSIS_FUNC_ENUM
    '   LINEAR_GRADIENT_VIA_CLOSE_DATA ? for both inter and extra
    LINEAR_INTERPOLATION_VIA_CLOSE_DATA
    LINEAR_EXTRAPOLATION_VIA_CLOSE_DATA
    EXPONENTIAL
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
    data As FuncDataFYDoubleType
End Type

Public Type DBFuncDataFYIPMTwoDoubleType
    id As Long
    sort_num As Long
    data_i As FuncDataFYDoubleType
    data_i_p_one As FuncDataFYDoubleType
    data_i_p_two As FuncDataFYDoubleType
    data_i_m_one As FuncDataFYDoubleType
    data_i_m_two As FuncDataFYDoubleType
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

Public Function ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM(input_ref_data_linear_interval_chk As RefDataLinearIntervalChk) As String
    Select Case input_ref_data_linear_interval_chk
        Case WITHIN_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "WITHIN_REF_DATA"
        Case ABOVE_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ABOVE_REF_DATA"
        Case BELOW_REF_DATA
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "BELOW_REF_DATA"
        Case NULL_INTERVAL_CHK
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "NULL_INTERVAL_CHK"
        Case ERROR_INTERVAL_CHK
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ERROR_INTERVAL_CHK"
        Case Else
            ASSIGN_REFDATALINEARINTERVALCHK_STRING_FROM_ENUM = "ERROR_INTERVAL_CHK"
    End Select
End Function

Public Function CREATE_FUNCDATAFYDOUBLETYPE(input_in_y As Double, input_out_f) As FuncDataFYDoubleType
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

Public Function JOIN_FUNCDATAFYDOUBLETYPE_TO_STRING_V000(input_type As FuncDataFYDoubleType) As String
    Dim result As String

    '   ... function incomplete

    JOIN_FUNCDATAFYDOUBLETYPE_TO_STRING_V000 = result
End Function

Public Function CREATE_DBFUNCDATAFYDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE(input_type_arr() As DBFuncDataFYDoubleType) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_type_arr, 1)
	result.u_bnd = UBound(input_type_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DBFUNCDATAFYDOUBLETYPE_ONE_DIM_ARRAYDIMSTYPE = result
End Function

'   General Calc

Public Function RETURN_Y_LINEAR_FUNCTION(gradient As Double, x_coordinate As Double, y_intercept As Double) As Double
    ' y = m * x + c
    RETURN_Y_LINEAR_FUNCTION = gradient * x_coordinate + y_intercept
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