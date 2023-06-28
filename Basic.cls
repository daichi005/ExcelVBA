VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'参照設定：Microsoft ActiveX Data Objects 6.1 Library
'参照設定：Microsoft Scripting Runtime
'参照設定：Windows Script Host Object Model

Option Explicit

Public wsh As WshShell
Public getData As BasicGetData
Public openFileDialog As BasicOpenFileDialog
Public exArray As BasicExtendArray

Private Sub Class_Initialize()

    Set wsh = New WshShell

    Set getData = New BasicGetData
    Set getData.setBasic = Me
    Set getData.setWsh = wsh
    
    Set openFileDialog = New BasicOpenFileDialog
    Set openFileDialog.setBasic = Me
    
    Set exArray = New BasicExtendArray
    Set exArray.setBasic = Me

End Sub
