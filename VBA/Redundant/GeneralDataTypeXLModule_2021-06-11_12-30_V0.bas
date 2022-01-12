Attribute VB_Name = "GeneralDataTypeXLModule"
'   Eriez Magnetics Australia Excel VBA
'   General Use Module
'   GeneralDataTypeXLModule
'   Leonard Sponza
'   Last Modified 11/06/2021 12:30
'   Date Time Version 00

Option Explicit

Public Enum CustomBoolean
    TRUE_CUST_BOOL = 1
    FALSE_CUST_BOOL = 2
    UNASSIGNED_CUST_BOOL = UNASSIGNED_LONG_VAL
    NULL_CUST_BOOL = NULL_LONG_VAL
    NOT_APPLICABLE_CUST_BOOL = NOT_APPLICABLE_LONG_VAL
    ERROR_CUST_BOOL = ERROR_LONG_VAL
    TEST_CUST_BOOL = TEST_LONG_VAL
End Enum