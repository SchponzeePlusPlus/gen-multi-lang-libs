Attribute VB_Name = "GeneralTestXLModule"
'   Excel VBA
'   General Use Testing Module
'   GeneralTestXLModule
'   Leonard Sponza
'   Last Modified 18/08/2021 16:40
'   Date Time Version 00

Option Explicit

Public Const EXPECTED_COL_NUM As Long = 5

Public Function Test1(inputRange As Range) As Variant
    Dim result As Variant
    Dim row_num As Long
    
    row_num = inputRange.Rows.Count
    If (inputRange.Columns.Count = EXPECTED_COL_NUM) Then
        result = inputRange(3, 1)
    End If
    Test1 = result
End Function

Sub Test()
    Dim MyArray() As Variant
    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    Dim rng As Range
    Dim wsr As Range

    ReDim MyArray(1 To 3, 1 To 3) ' make it flexible

    ' Fill array
    '  ...
    cnt = 0
    
    For i = LBound(MyArray, 1) To UBound(MyArray, 1)
        For j = LBound(MyArray, 2) To UBound(MyArray, 2)
            MyArray(i, j) = cnt
            cnt = cnt + 1
        Next
    Next
    
    Set wsr = ActiveWorkbook.Worksheets("Sheet4").[A1]
    Set rng = ActiveWorkbook.Worksheets("Sheet4").Range("A1")

    PrintArray MyArray, rng
End Sub

Public Function CONCAT_JOIN_RANGE(input_range As Range) As String
    Dim result As String
    Dim input_array() As Variant

    input_array = CONVERT_RANGE_TO_ONE_DIM_ARRAY(input_range, "DOWNWARDS_ALONG_COLS_FIRST")

    result = Join(input_array, " , ")

    CONCAT_JOIN_RANGE = result
End Function

Sub FindValue()
    
    Dim c As Range
    Dim firstAddress As String
    Dim found_ctr As Long

    found_ctr = 0
    
    With Worksheets(1).Range("Q1:Q6") 
        '   Set c = .Find(what:="o", lookin:=xlValues)
        Set c = .Find("o", .Cells(.Rows.Count,1), xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
        '   Set c = .Find("o", Worksheets(1).Range("A1:A1") , xlFormulas, xlPart, xlByColumns, xlNext, False, False, False)
        Debug.Print c.Address
        Debug.Print c.Value2
        If Not c Is Nothing Then 
            firstAddress = c.Address 
            Do 
                found_ctr = found_ctr + 1
                Set c = .FindNext(c)
                Debug.Print c.Address
                Debug.Print c.Value2
            Loop While (Not c Is Nothing) And (c.Address <> firstAddress)
        End If 
    End With
    
End Sub