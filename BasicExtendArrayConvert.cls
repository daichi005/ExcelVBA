VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicExtendArrayConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property

Public Function toRangeKey(ByRef arr() As String) As String

    Dim x As Long
    Dim y As Long
    
    Select Case basic.exArray.status.dimension(arr)
        Case 1
            x = UBound(arr, 1) - LBound(arr, 1) + 1
            toRangeKey = "A1:" & Replace(Cells(1, x).Address, "$", "")
        Case 2
            x = UBound(arr, 1) - LBound(arr, 1) + 1
            y = UBound(arr, 2) - LBound(arr, 2) + 1
            toRangeKey = "A1:" & Replace(Cells(x, y).Address, "$", "")
    End Select

End Function
