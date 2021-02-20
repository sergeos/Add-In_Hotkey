VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FormExp (file.cls)
' http://www.balagurov.com/software/formexp/

Option Explicit

Private File As Integer

Private Sub Class_Initialize()
    File = 0
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub

Public Function CreateFile(FileName As String) As Boolean
    On Error Resume Next
    
    Kill FileName
    Err.Clear

    File = FreeFile
    Open FileName For Binary Access Write Lock Read Write As File
    
    CreateFile = Err.Number = 0
End Function

Public Sub CloseFile()
    If File > 0 Then
        Close File
        File = 0
    End If
End Sub

Public Sub WriteLine(Optional ByVal Line As String = "")
    Put File, , Line & vbNewLine
End Sub
