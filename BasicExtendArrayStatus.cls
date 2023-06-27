VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicExtendArrayStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property

Function dimension(ByRef arr() As Variant) As Long

    Dim result As Long
    Dim buf As Long
    
    result = 0
    
    On Error Resume Next
    
    Do
        result = result + 1
        
        If result > 1 Then
            Err.Clear
        End If
        
        buf = UBound(arr, result)
        
    Loop Until Err.Number <> 0
    
    dimension = result - 1
    
End Function
Function IsReleased(ByRef arr() As Variant) As Boolean

    IsReleased = (arr Is Nothing)

End Function
