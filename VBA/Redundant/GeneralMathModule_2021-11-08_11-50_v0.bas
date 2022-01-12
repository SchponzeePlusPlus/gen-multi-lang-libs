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

Public Function NULLIFY_FUNCDATAFYDOUBLETYPE() As FuncDataFYDoubleType
    Dim result As FuncDataFYDoubleType

    result.in_y = NULL_DOUBLE_VAL
    result.out_f = NULL_DOUBLE_VAL

    NULLIFY_FUNCDATAFYDOUBLETYPE = result
End Function

Public Function JOIN_FUNCDATAFYDOUBLETYPE_TO_STRING_V000(input As FuncDataFYDoubleType) As String
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
