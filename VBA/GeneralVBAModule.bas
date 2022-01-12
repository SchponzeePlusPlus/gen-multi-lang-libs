Attribute VB_Name = "GeneralVBAModule"
'   Eriez Magnetics Australia MS VBA
'   General Use Module (for all MS applications)
'   GeneralVBAModule
'   Leonard Sponza
'   Last Modified 15/09/2021 13:35
'   Date Time Version 00

Option Explicit

'   Approximate limit acording to documentation
'   2GB limit?, according to forums
Public Const STRING_LENGTH_LIM As Long = (1 * 2 ^ 31) - 1

'   1 * 10^9 was used within the concat function previously
Public Const STRING_LENGTH_SAFE_LIM As Long = STRING_LENGTH_LIM - 100

Public Const JOIN_SPLIT_ARR_DELIMITER As String = " , "

'   more constants for join split array beginning and ends?

Public Function CHECK_DATA_TYPE_CAST_FROM_STRING(input_str As String, expected_data_type As String) As Boolean

'    Dim input_char_arr() As Char
    
    Dim valid_input As Boolean
    
    Dim decimal_ctr As Integer, neg_sign_ctr As Integer
    
    Dim i As Integer
    
    Dim input_char_elem As String
    
    valid_input = False
    
    decimal_ctr = 0
    neg_sign_ctr = 0
    
    Dim test As String
    
'    input_char_arr = input_str.ToCharArray

    Select Case expected_data_type
'       Character number not considered
        Case "String"
            valid_input = True
'       Digit number not considered
        Case "Integer"
            If (input_str <> "") Then
            
                valid_input = True
                
                For i = 1 To (Len(input_str))
                    input_char_elem = Mid(input_str, i, 1)
                    If ((input_char_elem <> "0") And (input_char_elem <> "1") And (input_char_elem <> "2") And (input_char_elem <> "3") And (input_char_elem <> "4") And (input_char_elem <> "5") And (input_char_elem <> "6") And (input_char_elem <> "7") And (input_char_elem <> "8") And (input_char_elem <> "9") And (input_char_elem <> "-")) Then
                        valid_input = False
                    End If
                    
                    If ((input_char_elem = "-") And (neg_sign_ctr = 0) And (i = 1)) Then
                        neg_sign_ctr = neg_sign_ctr + 1
                    ElseIf ((input_char_elem = "-") And (neg_sign_ctr <> 0) And (i > 1)) Then
                        valid_input = False
                    End If
                Next i
            ElseIf (input_str = "") Then
                valid_input = True
            Else
                valid_input = False
            End If
'       Digit number not considered
        Case "Double"
            If (input_str <> "") Then
            
                valid_input = True
            
                For i = 1 To (Len(input_str))
                    input_char_elem = Mid(input_str, i, 1)
                    If ((input_char_elem <> "0") And (input_char_elem <> "1") And (input_char_elem <> "2") And (input_char_elem <> "3") And (input_char_elem <> "4") And (input_char_elem <> "5") And (input_char_elem <> "6") And (input_char_elem <> "7") And (input_char_elem <> "8") And (input_char_elem <> "9") And (input_char_elem <> ".") And (input_char_elem <> "-")) Then
                        valid_input = False
                    End If
                    
                    If ((input_char_elem = "-") And (neg_sign_ctr = 0) And (i = 1)) Then
                        neg_sign_ctr = neg_sign_ctr + 1
                    ElseIf ((input_char_elem = "-") And (neg_sign_ctr <> 0) And (i > 1)) Then
                        valid_input = False
                    End If
                    
                    If ((input_char_elem = ".") And (decimal_ctr = 0) And (i > 1)) Then
                        decimal_ctr = decimal_ctr + 1
                    ElseIf ((input_char_elem = ".") And (decimal_ctr <> 0) And (i > 1)) Then
                        valid_input = False
                    End If
                Next i
            ElseIf (input_str = "") Then
                valid_input = True
            Else
                valid_input = False
            End If
        Case Else
            valid_input = False
    End Select
    
    If (valid_input = True) Then
        CHECK_DATA_TYPE_CAST_FROM_STRING = True
    Else
        CHECK_DATA_TYPE_CAST_FROM_STRING = False
    End If

End Function

Public Function CHECK_NULL_STRING_INPUT(input_str As String) As Boolean
    If (input_str = "") Then
        CHECK_NULL_STRING_INPUT = True
    Else
        CHECK_NULL_STRING_INPUT = False
    End If
End Function

Public Function PRINT_VARIANT_TYPENAME(input_var As Variant) As String
    
    Dim test_var As Variant
    test_var = "String"
    
    Dim result As String
    
'    PRINT_CELL_VALUE_TYPENAME = TypeName(input_cell.Value2)
'    PRINT_CELL_VALUE_TYPENAME = TypeName(test_var)
    result = TypeName(input_var)
'    result = TypeName(input_cell.Value2)
    PRINT_VARIANT_TYPENAME = result
End Function

Public Function CAST_VARIANT_TO_INTEGER(input_var As Variant) As Integer

    Dim check_data_type_validity As Boolean
    
    check_data_type_validity = False
    
    check_data_type_validity = CHECK_DATA_TYPE_CAST_FROM_STRING(CStr(input_var), "Integer")

    If ((IsNumeric(input_var)) And (check_data_type_validity = True)) Then
        If ((TypeName(input_var) = "Byte") Or (TypeName(input_var) = "Double") Or (TypeName(input_var) = "Long") Or (TypeName(input_var) = "Single") Or (TypeName(input_var) = "Currency") Or (TypeName(input_var) = "Decimal")) Then
            CAST_VARIANT_TO_INTEGER = CInt(input_var)
'            CAST_VARIANT_TO_INTEGER = 1#
        ElseIf (TypeName(input_var) = "Integer") Then
            CAST_VARIANT_TO_INTEGER = input_var
'            CAST_VARIANT_TO_INTEGER = 2#
        ElseIf ((TypeName(input_var) = "Null") Or (TypeName(input_var) = "Empty")) Then
            CAST_VARIANT_TO_INTEGER = NULL_INTEGER_VAL
'            CAST_VARIANT_TO_INTEGER = 2#
        ElseIf (TypeName(input_var) = "Error") Then
            CAST_VARIANT_TO_INTEGER = ERROR_INTEGER_VAL
'            CAST_VARIANT_TO_INTEGER = 2#
        Else
            CAST_VARIANT_TO_INTEGER = CInt(input_var)
'            CAST_VARIANT_TO_INTEGER = 3#
        End If
    Else
        CAST_VARIANT_TO_INTEGER = ERROR_INTEGER_VAL
    End If
End Function

Public Function CAST_VARIANT_TO_LONG(input_var As Variant) As Long

    Dim check_data_type_validity As Boolean
    
    check_data_type_validity = False
    
    check_data_type_validity = CHECK_DATA_TYPE_CAST_FROM_STRING(CStr(input_var), "Integer")

    If ((IsNumeric(input_var)) And (check_data_type_validity = True)) Then
        If ((TypeName(input_var) = "Byte") Or (TypeName(input_var) = "Double") Or (TypeName(input_var) = "Integer") Or (TypeName(input_var) = "Single") Or (TypeName(input_var) = "Currency") Or (TypeName(input_var) = "Decimal")) Then
            CAST_VARIANT_TO_LONG = CLng(input_var)
'            CAST_VARIANT_TO_LONG = 1#
        ElseIf (TypeName(input_var) = "Long") Then
            CAST_VARIANT_TO_LONG = input_var
'            CAST_VARIANT_TO_LONG = 2#
        ElseIf ((TypeName(input_var) = "Null") Or (TypeName(input_var) = "Empty")) Then
            CAST_VARIANT_TO_LONG = NULL_LONG_VAL
'            CAST_VARIANT_TO_INTEGER = 2#
        ElseIf (TypeName(input_var) = "Error") Then
            CAST_VARIANT_TO_LONG = ERROR_LONG_VAL
'            CAST_VARIANT_TO_INTEGER = 2#
        Else
            CAST_VARIANT_TO_LONG = CLng(input_var)
'            CAST_VARIANT_TO_LONG = 3#
        End If
    Else
        CAST_VARIANT_TO_LONG = ERROR_LONG_VAL
    End If
End Function

Public Function CAST_VARIANT_TO_DOUBLE(input_var As Variant) As Double
    Dim result as Double
    Dim process As Double
    Dim input_gnvs As GeneralNumVarState
    
    process = UNASSIGNED_DOUBLE_VAL
    input_gnvs = UNASSIGNED_GNVS

    If (TypeName(input_var) = "Double") Then
        process = input_var
    ElseIf ((TypeName(input_var) = "Byte") Or (TypeName(input_var) = "Integer") Or (TypeName(input_var) = "Long") Or (TypeName(input_var) = "Single") Or (TypeName(input_var) = "Currency") Or (TypeName(input_var) = "Decimal")) Then
        process = CDbl(input_var)
    ElseIf (TypeName(input_var) = "String") Then
        '   Make sure the MS Access Object Library is enabled in VBA Tools>References
        If ((IsNumeric(input_var)) And (CHECK_DATA_TYPE_CAST_FROM_STRING(CStr(Access.Nz(input_var, NULL_DOUBLE_VAL)), "Double") = True)) Then
            process = CDbl(input_var)
        ElseIf ((IsNumeric(input_var)) And (CHECK_DATA_TYPE_CAST_FROM_STRING(CStr(Access.Nz(input_var, NULL_DOUBLE_VAL)), "Integer") = True)) Then
            process = CDbl(input_var)
        Else
            process = ERROR_DOUBLE_VAL
        End If

    ElseIf ((TypeName(input_var) = "Null") Or (TypeName(input_var) = "Empty")) Then
        process = NULL_DOUBLE_VAL
    ElseIf (TypeName(input_var) = "Error") Then
        process = ERROR_DOUBLE_VAL
    Else
        process = ERROR_DOUBLE_VAL
    End If

    input_gnvs = ASSIGN_VARIANT_GNVS(process)

    If (input_gnvs = VALID_GNVS) Then
        result = process
    ElseIf (input_gnvs <> VALID_GNVS) Then
        result = ASSIGN_DOUBLE_VS_FROM_GNVS(input_gnvs)
    End If

    CAST_VARIANT_TO_DOUBLE = result
End Function

Public Function CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_TN_V000(input_var As Variant, var_type_name As String) As Double
    Dim result As Double

    If ((var_type_name = "Byte") Or (var_type_name = "Integer") Or (var_type_name = "Long") Or (var_type_name = "Single") Or (var_type_name = "Currency") Or (var_type_name = "Decimal")) Then
        result = CDbl(input_var)
    ElseIf (var_type_name = "Double") Then
        result = input_var
    '   Else covers "Null", "Empty", "Error" and anything else
    Else
        result = 0
    End If

    CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_TN_V000 = result
End Function

Public Function CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_V000(input_var As Variant) As Double
    Dim result As Double
    Dim var_type_name As String

    var_type_name = TypeName(input_var)

    result = CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_TN_V000(input_var, var_type_name)

    CAST_VARIANT_TO_DOUBLE_NO_ENUM_CODE_VIA_IV_V000 = result
End Function

Public Function CAST_VARIANT_TO_STRING(input_var As Variant) As String
    Dim process As String
    Dim result As String
    Dim check_data_type_validity As Boolean
    
    check_data_type_validity = False

    process = CStr(Access.Nz(input_var, "(Null)"))
    
    check_data_type_validity = CHECK_DATA_TYPE_CAST_FROM_STRING(process, "String")
    
    If (check_data_type_validity = True) Then
        result = process
    Else
        result = "(Error)"
    End If
    CAST_VARIANT_TO_STRING = result
End Function