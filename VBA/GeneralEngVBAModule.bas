Attribute VB_Name = "GeneralEngVBAModule"
'
'	@file GeneralEngVBAModule.bas
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

Public Function CALC_FORCE_N_VIA_M_A_V000(mass_kg As Double, accel_mpssq As Double) As Double
	CALC_FORCE_N_VIA_M_A_V000 = (mass_kg * accel_mpssq)
End Function

Public Function CALC_VOLTAGE_V_VIA_I_R_V000(current_a As Double, resistance_ohms As Double) As Double
	CALC_VOLTAGE_V_VIA_I_R_V000 = (current_a * resistance_ohms)
End Function

Public Function CALC_POWER_W_VIA_V_I_V000(voltage_v As Double, current_a As Double) As Double
	CALC_POWER_W_VIA_V_I_V000 = (voltage_v * current_a)
End Function

Public Function CALC_R_TWO_V_OUT_VIA_VOLT_DIV_V_IN_R_ONE_R_TWO_V000(v_in As Double, r_one As Double, r_two As Double) As Double
	CALC_R_TWO_V_OUT_VIA_VOLT_DIV_V_IN_R_ONE_R_TWO_V000 = (v_in * (r_two / (r_one + r_two)))
End Function