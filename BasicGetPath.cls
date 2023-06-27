VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basicGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic
Private wsh As WshShell

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property
Public Property Set setWsh(ByRef obj As WshShell)

    Set wsh = obj

End Property

Public Function deskTop() As String
    
    deskTop = wsh.SpecialFolders("DeskTop")

End Function
Public Function startup() As String

    startup = wsh.SpecialFolders("Startup")

End Function
Public Function downloads() As String

    downloads = Replace(deskTop, "Desktop", "Downloads")

End Function

