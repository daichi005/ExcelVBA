VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicExtendArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic
Public cast As BasicExtendArrayCast
Public status As BasicExtendArrayStatus
Public convert As BasicExtendArrayConvert

Private Sub Class_Initialize()

    Set cast = New BasicExtendArrayCast
    Set status = New BasicExtendArrayStatus
    Set convert = New BasicExtendArrayConvert

End Sub
Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj
    Set cast.setBasic = obj
    Set status.setBasic = obj
    Set convert.setBasic = obj

End Property
Public Sub writeToSheet(ByRef arr() As String, ByRef ws As Worksheet, Optional ByVal start_cell As String = "A1")

    ws.Range(start_cell).Range(basic.exArray.convert.toRangeKey(arr)).Value = arr

End Sub
Public Sub writeToSheetCastOption(ByRef arr() As String, ByRef cast() As String, ByRef ws As Worksheet, Optional ByVal start_cell As String = "A1")



End Sub
