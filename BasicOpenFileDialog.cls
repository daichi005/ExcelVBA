VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicOpenFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private basic As basic
Private fd As FileDialog

Public Property Set setBasic(ByRef obj As basic)

    Set basic = obj

End Property
Private Function fileDialogSingleProcess(ByVal title As String, ByVal fileCategory As String, ByVal fileType As String) As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.title = title
    fd.AllowMultiSelect = False
    Call fd.Filters.Add(fileCategory, fileType, 1)
    
    If fd.Show = -1 Then
    
        fileDialogSingleProcess = CStr(fd.SelectedItems(1))
        
    Else
    
        fileDialogSingleProcess = ""
        
    End If
    
    Call fd.Filters.Clear

End Function
Private Function fileDialogMultiProcess(ByVal title As String, ByVal fileCategory As String, ByVal fileType As String) As String()

    Dim pathTable() As String
    Dim bufTable() As Variant
    
    Dim i As Long
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.title = title
    fd.AllowMultiSelect = True
    Call fd.Filters.Add(fileCategory, fileType, 1)
    
    If fd.Show = -1 Then
    
        ReDim bufTable(1 To fd.SelectedItems.Count)

        For i = 1 To fd.SelectedItems.Count
        
            bufTable(i) = fd.SelectedItems(i)
        
        Next i
        
        pathTable = basic.exArray.cast.variantToAnyting(bufTable, dkString, 1)
        
    End If
    
    fileDialogMultiProcess = pathTable
    
    Call fd.Filters.Clear

End Function
Public Function excelSingle(Optional ByVal title As String = "Excel�t�@�C����I������") As String

    excelSingle = fileDialogSingleProcess(title, "Excel�t�@�C��", "*.xls; *.xlsx; *.xlsm")

End Function
Public Function csvlSingle(Optional ByVal title As String = "csv�t�@�C����I������") As String

    csvlSingle = fileDialogSingleProcess(title, "csv�t�@�C��", "*.csv")

End Function
Public Function anytinglSingle(Optional ByVal title As String = "�t�@�C����I������") As String

    anytinglSingle = fileDialogSingleProcess(title, "���ׂẴt�@�C��", "*.*")

End Function
Public Function excelMulti(Optional ByVal title As String = "Excel�t�@�C����I������") As String()

    excelMulti = fileDialogMultiProcess(title, "Excel�t�@�C��", "*.xls; *.xlsx; *.xlsm")

End Function
Public Function csvlMulti(Optional ByVal title As String = "csv�t�@�C����I������") As String()

    csvlMulti = fileDialogMultiProcess(title, "csv�t�@�C��", "*.csv")

End Function
Public Function anytinglMulti(Optional ByVal title As String = "�t�@�C����I������") As String()

    anytinglMulti = fileDialogMultiProcess(title, "���ׂẴt�@�C��", "*.*")

End Function