VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic
Private wsh As WshShell

Public path As BasicGetDataPath
Public config As BasicGetDataConfig

Private Sub Class_Initialize()

    Set path = New BasicGetDataPath
    Set config = New BasicGetDataConfig

End Sub
Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj
    Set path.setBasic = obj
    Set config.setBasic = obj

End Property
Public Property Set setWsh(ByRef obj As WshShell)

    Set wsh = obj
    Set path.setWsh = obj

End Property
