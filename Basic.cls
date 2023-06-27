VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public wsh As WshShell
Public getPath As basicGetPath
Public openFileDialog As BasicOpenFileDialog
Public exArray As BasicExtendArray

Private Sub Class_Initialize()

    Set wsh = New WshShell

    Set getPath = New basicGetPath
    Set getPath.setBasic = Me
    Set getPath.setWsh = wsh
    
    Set openFileDialog = New BasicOpenFileDialog
    Set openFileDialog.setBasic = Me
    
    Set exArray = New BasicExtendArray
    Set exArray.setBasic = Me

End Sub
