Attribute VB_Name = "Module1"
Function CUSTM_LNR_INTERPOLATION(y_0, y_2, x_0, x_1, x_2)
    CUSTM_LNR_INTERPOLATION = ((y_2 - y_0) / (x_2 - x_0)) * (x_1 - x_0) + y_0
End Function

Function CUSTM_VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)
    If Application.WorksheetFunction.VLookup(lookup_value, table_array, col_index_num, range_lookup) = "" Then
        CUSTM_VLOOKUP = ""
    Else
        CUSTM_VLOOKUP = Application.WorksheetFunction.VLookup(lookup_value, table_array, col_index_num, range_lookup)
    End If
End Function
