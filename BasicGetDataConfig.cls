VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicGetDataConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property
Public Function references(ByRef wb As Workbook) As String()

    Dim ref As Object
    
    Dim result() As String
    
    Dim refCount As Long
    Dim i As Long
    
    refCount = wb.VBProject.references.Count
    ReDim result(1 To refCount, 1 To 2)

    For i = LBound(result, 1) To UBound(result, 1) Step 1
    
        Set ref = ThisWorkbook.VBProject.references.Item(i)
        
        result(i, 1) = ref.Description
        result(i, 2) = ref.FullPath
        
    Next i
    
    references = result
    
End Function
