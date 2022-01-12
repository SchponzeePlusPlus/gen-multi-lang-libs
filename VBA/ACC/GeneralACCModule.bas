Attribute VB_Name = "GeneralACCModule"
'   MS Access VBA
'   General Use Module
'   Access Database
'   GeneralACCModule
'   Leonard Sponza
'   Last Modified 23/07/2021 15:15
'   Date Time Version 00

Option Compare Database
Option Explicit

Public Function RETURN_RESULTANT_VALID_STR_FIELD(default_str As String, overwrite_str As String) As String
    If (Not (IsNull(overwrite_str))) Then
        RETURN_RESULTANT_VALID_STR_FIELD = overwrite_str
    Else
        RETURN_RESULTANT_VALID_STR_FIELD = default_str
    End If
End Function

Public Function RETURN_RESULTANT_VALID_INT_FIELD(default_int_raw As Variant, overwrite_int_raw As Variant) As Integer
'Public Function RETURN_RESULTANT_VALID_INT_FIELD(default_int_raw As Integer, overwrite_int_raw As Integer) As Integer
'    Dim default_int As Integer
'    Dim overwrite_int As Integer
    Dim default_int As Variant
    Dim overwrite_int As Variant
'    default_int = Nz(default_int_raw, 0)
'    overwrite_int = Nz(overwrite_int_raw, 0)
'    default_int = CInt(default_int_raw)
'    overwrite_int = CInt(overwrite_int_raw)
    default_int = default_int_raw
    overwrite_int = overwrite_int_raw
    
'    If ((Not (IsNull(overwrite_int))) Or (Not (IsEmpty(overwrite_int))) Or (overwrite_int <> 0)) Then
'    If ((Not (IsNull(overwrite_int))) Or (overwrite_int <> 0)) Then
    If ((Not (IsNull(overwrite_int)))) Then
        RETURN_RESULTANT_VALID_INT_FIELD = CInt(Nz(overwrite_int, 0))
'        RETURN_RESULTANT_VALID_INT_FIELD = 1
    Else
        RETURN_RESULTANT_VALID_INT_FIELD = CInt(Nz(default_int, 0))
'        RETURN_RESULTANT_VALID_INT_FIELD = 2
    End If
'    RETURN_RESULTANT_VALID_INT_FIELD = overwrite_int
End Function

Public Function RETURN_RESULTANT_VALID_VARIANT_FIELD(default_var As Variant, overwrite_var As Variant) As Variant
    Dim result As Variant

    Dim overwrite_val_active As Boolean
    Dim overwrite_var_gnvs As GeneralNumVarState

    overwrite_var_gnvs = ASSIGN_VARIANT_GNVS(overwrite_var)

    If ((Not (IsNull(overwrite_var))) And (overwrite_var_gnvs <> UNASSIGNED_GNVS) And (overwrite_var_gnvs <> NULL_GNVS) And (overwrite_var_gnvs <> UNKNOWN_GNVS)) Then
'    If ((Not (IsNull(overwrite_var)))) Then
'        RETURN_RESULTANT_VALID_VARIANT_FIELD = overwrite_var
'        RETURN_RESULTANT_VALID_INT_FIELD = 1
        overwrite_val_active = True
    Else
'        RETURN_RESULTANT_VALID_VARIANT_FIELD = default_var
'        RETURN_RESULTANT_VALID_INT_FIELD = 2
        overwrite_val_active = False
    End If

    If (overwrite_val_active = True) Then
        result = overwrite_var
    Else
        result = default_var
    End If

    RETURN_RESULTANT_VALID_VARIANT_FIELD = result
End Function

Public Function RETURN_RELEVANT_SORT_NUMBER(sort_num_i As Variant, inc_dec As Integer) As Long
    If (Not (IsNull(sort_num_i))) Then
        RETURN_RELEVANT_SORT_NUMBER = sort_num_i + inc_dec
    Else
        RETURN_RELEVANT_SORT_NUMBER = NULL_LONG_VAL
    End If
End Function

Public Function CUSTOM_DLOOKUP_LONG(expression As Variant, domain As Variant, criteria As Variant, NULL_LONG_VAL As Long) As Long
    Dim process As Variant
    
    process = Application.DLookup(expression, domain, criteria)
    
    If (Not (IsNull(process))) Then
        CUSTOM_DLOOKUP_LONG = process
    Else
        CUSTOM_DLOOKUP_LONG = NULL_LONG_VAL
    End If
End Function