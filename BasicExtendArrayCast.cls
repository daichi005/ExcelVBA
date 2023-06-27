VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicExtendArrayCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property
Public Function variantToAnyting(ByRef arr() As Variant, ByVal type_code As typeList, Optional ByVal based_index As Long = 1) As Variant

    Dim result_string() As String
    Dim result_long() As Long
    Dim result_date() As Date
    Dim result_double() As Double
    Dim result_variant() As Variant
    
    Dim dim1 As Long
    Dim dim2 As Long
    
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim y As Long
    
    Select Case basic.exArray.status.dimension(arr)
        Case 1
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        result_string(i) = arr(x)
                        x = x + 1
                    Next i
                    variantToAnyting = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        result_long(i) = arr(x)
                        x = x + 1
                    Next i
                    variantToAnyting = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        result_date(i) = arr(x)
                        x = x + 1
                    Next i
                    variantToAnyting = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        result_double(i) = arr(x)
                        x = x + 1
                    Next i
                    variantToAnyting = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        result_variant(i) = arr(x)
                        x = x + 1
                    Next i
                    variantToAnyting = result_variant
            End Select
        Case 2
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            dim2 = UBound(arr, 2) - LBound(arr, 2) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        For j = LBound(result_string, 2) To UBound(result_string, 2) Step 1
                            result_string(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    variantToAnyting = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        For j = LBound(result_long, 2) To UBound(result_long, 2) Step 1
                            result_long(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    variantToAnyting = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        For j = LBound(result_date, 2) To UBound(result_date, 2) Step 1
                            result_date(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    variantToAnyting = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        For j = LBound(result_double, 2) To UBound(result_double, 2) Step 1
                            result_double(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    variantToAnyting = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        For j = LBound(result_variant, 2) To UBound(result_variant, 2) Step 1
                            result_variant(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    variantToAnyting = result_variant
            End Select
        End Select

End Function
Public Function stringToAnytiong(ByRef arr() As Variant, ByVal type_code As typeList, Optional ByVal based_index As Long = 1) As Variant

    Dim result_string() As String
    Dim result_long() As Long
    Dim result_date() As Date
    Dim result_double() As Double
    Dim result_variant() As Variant
    
    Dim dim1 As Long
    Dim dim2 As Long
    
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim y As Long
    
    Select Case basic.exArray.status.dimension(arr)
        Case 1
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        result_string(i) = arr(x)
                        x = x + 1
                    Next i
                    stringToAnytiong = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        result_long(i) = arr(x)
                        x = x + 1
                    Next i
                    stringToAnytiong = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        result_date(i) = arr(x)
                        x = x + 1
                    Next i
                    stringToAnytiong = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        result_double(i) = arr(x)
                        x = x + 1
                    Next i
                    stringToAnytiong = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        result_variant(i) = arr(x)
                        x = x + 1
                    Next i
                    stringToAnytiong = result_variant
            End Select
        Case 2
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            dim2 = UBound(arr, 2) - LBound(arr, 2) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        For j = LBound(result_string, 2) To UBound(result_string, 2) Step 1
                            result_string(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    stringToAnytiong = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        For j = LBound(result_long, 2) To UBound(result_long, 2) Step 1
                            result_long(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    stringToAnytiong = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        For j = LBound(result_date, 2) To UBound(result_date, 2) Step 1
                            result_date(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    stringToAnytiong = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        For j = LBound(result_double, 2) To UBound(result_double, 2) Step 1
                            result_double(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    stringToAnytiong = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        For j = LBound(result_variant, 2) To UBound(result_variant, 2) Step 1
                            result_variant(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    stringToAnytiong = result_variant
            End Select
        End Select

End Function
Public Function longToAnytiong(ByRef arr() As Variant, ByVal type_code As typeList, Optional ByVal based_index As Long = 1) As Variant

    Dim result_string() As String
    Dim result_long() As Long
    Dim result_date() As Date
    Dim result_double() As Double
    Dim result_variant() As Variant
    
    Dim dim1 As Long
    Dim dim2 As Long
    
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim y As Long
    
    Select Case basic.exArray.status.dimension(arr)
        Case 1
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        result_string(i) = arr(x)
                        x = x + 1
                    Next i
                    longToAnytiong = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        result_long(i) = arr(x)
                        x = x + 1
                    Next i
                    longToAnytiong = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        result_date(i) = arr(x)
                        x = x + 1
                    Next i
                    longToAnytiong = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        result_double(i) = arr(x)
                        x = x + 1
                    Next i
                    longToAnytiong = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1)
                    x = LBound(arr, 1)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        result_variant(i) = arr(x)
                        x = x + 1
                    Next i
                    longToAnytiong = result_variant
            End Select
        Case 2
            dim1 = UBound(arr, 1) - LBound(arr, 1) + based_index
            dim2 = UBound(arr, 2) - LBound(arr, 2) + based_index
            Select Case type_code
                Case dkString
                    ReDim result_string(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_string, 1) To UBound(result_string, 1) Step 1
                        For j = LBound(result_string, 2) To UBound(result_string, 2) Step 1
                            result_string(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    longToAnytiong = result_string
                Case dkLong
                    ReDim result_long(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_long, 1) To UBound(result_long, 1) Step 1
                        For j = LBound(result_long, 2) To UBound(result_long, 2) Step 1
                            result_long(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    longToAnytiong = result_long
                Case dkDate
                    ReDim result_date(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_date, 1) To UBound(result_date, 1) Step 1
                        For j = LBound(result_date, 2) To UBound(result_date, 2) Step 1
                            result_date(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    longToAnytiong = result_date
                Case dkDouble
                    ReDim result_double(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_double, 1) To UBound(result_double, 1) Step 1
                        For j = LBound(result_double, 2) To UBound(result_double, 2) Step 1
                            result_double(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    longToAnytiong = result_double
                Case dkVariant
                    ReDim result_variant(based_index To dim1, based_index To dim2)
                    x = LBound(arr, 1)
                    y = LBound(arr, 2)
                    For i = LBound(result_variant, 1) To UBound(result_variant, 1) Step 1
                        For j = LBound(result_variant, 2) To UBound(result_variant, 2) Step 1
                            result_variant(i, j) = arr(x, y)
                            y = y + 1
                        Next j
                        x = x + 1
                    Next i
                    longToAnytiong = result_variant
            End Select
        End Select

End Function
