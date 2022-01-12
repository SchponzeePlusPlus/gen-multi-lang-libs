Attribute VB_Name = "GeneralEngineeringModule"
'
'	@file GeneralEngineeringModule.bas
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

'	License: N/A PRIVATE SOURCE CODE
'	Licensed Material - Property of Eriez
'	Copyright (c) 2021 Eriez Australia. All rights reserved.
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

Public Function CONVERT_VALUE_AREA_SQMM_TO_SQM_V000(input_area_sqmm As Double) As Double
	Dim result As Double
	
	result = input_area_sqmm / 10^(6)

	CONVERT_VALUE_AREA_SQMM_TO_SQM_V000 = result
End Function