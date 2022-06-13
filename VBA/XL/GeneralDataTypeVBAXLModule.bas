Attribute VB_Name = "GeneralTemplateModule"
'
'	@file GeneralTemplateModule.bas
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

Option Explicit

Public Function WRAP_CELL_VAL_TO_VARIANT_V000(in_cell As Range) As Variant
    '   does it need a check to ensure range is one cell?
    WRAP_CELL_VAL_TO_VARIANT_V000 = in_cell.Value2
End Function

Public Function CHECK_RANGE_ROW_CNT_V000(in_range As Range, expected_row_cnt As Integer) As Boolean
    Dim result As Boolean
    If (in_range.Rows.Count = expected_row_cnt) Then
        result = True
    Else
        result = False
    End If
        CHECK_RANGE_ROW_CNT_V000 = result
End Function

Public Function CHECK_RANGE_COL_CNT_V000(in_range As Range, expected_col_cnt As Integer) As Boolean
    Dim result As Boolean
    If (in_range.Columns.Count = expected_col_cnt) Then
        result = True
    Else
        result = False
    End If
        CHECK_RANGE_COL_CNT_V000 = result
End Function

Public Function CHECK_RANGE_ROW_COL_CNT_V000(in_range As Range, expected_row_cnt As Integer, expected_col_cnt As Integer) As Boolean
    Dim result As Boolean
    If ((CHECK_RANGE_ROW_CNT_V000(in_range, expected_row_cnt)) And (CHECK_RANGE_COL_CNT_V000(in_range, expected_col_cnt))) Then
        result = True
    Else
        result = False
    End If
    CHECK_RANGE_ROW_COL_CNT_V000 = result
End Function

Public Function CHECK_RANGE_SINGLE_ROW_CNT_V000(in_range As Range) As Boolean
    CHECK_RANGE_SINGLE_ROW_CNT_V000 = CHECK_RANGE_ROW_CNT_V000(in_range, 1)
End Function

Public Function CHECK_RANGE_SINGLE_COL_CNT_V000(in_range As Range) As Boolean
    CHECK_RANGE_SINGLE_COL_CNT_V000 = CHECK_RANGE_COL_CNT_V000(in_range, 1)
End Function

Public Function CHECK_RANGE_SINGLE_ROW_COL_CNT_V000(in_range As Range) As Boolean
    CHECK_RANGE_SINGLE_COL_CNT_V000 = CHECK_RANGE_ROW_COL_CNT_V000(in_range, 1, 1)
End Function

Public Function WRAP_RANGE_CELL_VAL_TO_VARIANT_V000(in_cell As Range) As Variant
End Function

Public Function TEMPLATE_FUNC(example As datatype) As resultdatatype
End Function