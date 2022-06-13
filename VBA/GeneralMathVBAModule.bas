Attribute VB_Name = "GeneralMathVBAModule"
'
'	@file GeneralMathModule.bas
'
'	Write description of source file here for dOxygen.
'
'	System:         Independant Tool
'	Component Name: 
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
'	Description / Abstract: Module file
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

Public Function CALC_N_FACTORIAL_V000(n As Long) As Long
    Dim i As Long
    Dim result As Long

    '   0! = 1
    '   https://www.cuemath.com/numbers/factorial/
    result = 1
    For i = 1 To n
        result = result * i
    Next

    CALC_N_FACTORIAL_V000 = result
End Function

Public Function CALC_HYP_PYTHAGOREAN_THEOREM_VIA_A_B_V000(a As Double, b As Double) As Double
    CALC_HYP_PYTHAGOREAN_THEOREM_VIA_A_B_V000 = (a ^(2) + b ^(2))^(1/2)
End Function