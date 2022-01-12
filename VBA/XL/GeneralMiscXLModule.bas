Attribute VB_Name = "GeneralMiscXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Misc Module
'   GeneralMiscXLModule
'   Leonard Sponza
'   Last Modified 18/08/2021 16:40
'   Date Time Version 00

Option Explicit

Public Function CLEAN_CAST_CELL_VALUE_TO_STRING(input_cell As Variant) As String
    Dim result As String
    
    Dim input_str As String
    Dim input_char_elem As String
    Dim i As Integer
    Dim space_ctr As Long
    Dim result_len As Long
    
    input_str = CAST_CELL_VALUE_TO_STRING(input_cell)
    
    space_ctr = 0
    
    For i = 1 To (Len(input_str))
        input_char_elem = Mid(input_str, i, 1)
        If (input_char_elem = " ") Then
            space_ctr = space_ctr + 1
        Else
            space_ctr = 0
        End If
    Next i
    
    If (space_ctr > 0) Then
        result_len = (Len(input_str)) - space_ctr
        result = Left(input_str, result_len)
    Else
        result = input_str
    End If
    
    '   result = RTrim(input_str)
    
    CLEAN_CAST_CELL_VALUE_TO_STRING = result
End Function