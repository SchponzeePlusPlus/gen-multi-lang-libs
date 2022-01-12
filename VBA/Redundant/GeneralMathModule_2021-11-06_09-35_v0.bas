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

Public Type POINT_TWO_DIM_DOUBLE
    x As Double
    y As Double
End Type

Public Type POINT_THR_DIM_DOUBLE
    xy As POINT_TWO_DIM_DOUBLE
    z As Double
End Type

Public Type FUNCTION_DATA_FX_DOUBLE
    in_x As Double
    out_f As Double
End Type

Public Type FUNCTION_DATA_FY_DOUBLE
    in_y As Double
    out_f As Double
End Type

Public Type FUNCTION_DATA_FXY_DOUBLE
    in_xy As POINT_TWO_DIM_DOUBLE
    out_f As Double
End Type

Public Function SOLVE_SIMUL_FUNC_FY_DOUB_FOR_Y_V000() As Double
    Dim result As Double

    SOLVE_SIMUL_FUNC_FY_DOUB_FOR_Y_V000 = result
End Function
