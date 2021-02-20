VERSION 5.00
Begin VB.Form ChooseForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5970
   HelpContextID   =   1
   Icon            =   "Choose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox OverwriteCheck 
      Caption         =   "&Overwrite existing files without prompting"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   3975
   End
   Begin VB.CommandButton AboutButton 
      Caption         =   "Abou&t..."
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BrowseCommand 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox FolderText 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   3975
   End
   Begin VB.CommandButton UnselectAllButton 
      Caption         =   "&Unselect All"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton SelectAllButton 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox FormsList 
      Height          =   4350
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label FolderLabel 
      Caption         =   "Save exported form(s) in the following &folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label PromptLabel 
      Caption         =   "&Select form(s) to export:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "ChooseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FormExp (choose.frm)
' http://www.balagurov.com/software/formexp/

Option Explicit

Private IDE As VBIDE.VBE

Private Const SectionName As String = "Options"
Private Const FolderKeyName As String = "Target Folder"
Private Const OverwriteKeyName As String = "Overwrite without prompt"

Public Sub Initialize(ByRef IDE_ As VBIDE.VBE)
    On Error Resume Next
    
    Set IDE = IDE_
    
    Caption = AddInName

    OKButton.Enabled = False

    FolderText.Text = GetSetting(AddInName, SectionName, FolderKeyName)
    If (Len(FolderText.Text) = 0) Or (PathIsDirectory(FolderText.Text) = 0) Then
        FolderText.Text = GetSpecialFolderLocation(CSIDL_PERSONAL, Me.hwnd)
    End If
    
    OverwriteCheck.Value = GetSetting(AddInName, SectionName, OverwriteKeyName, 0)
    
    If IDE.ActiveVBProject Is Nothing Then Exit Sub
    
    Dim Component As VBIDE.VBComponent
    For Each Component In IDE.ActiveVBProject.VBComponents
        If Component.Type = vbext_ct_MSForm Then
            FormsList.AddItem Component.Name
        End If
    Next
    
    If FormsList.ListCount > 0 Then
        OKButton.Enabled = True
        FormsList.Selected(0) = True
        
        If Not (IDE.SelectedVBComponent Is Nothing) Then
            If IDE.SelectedVBComponent.Type = vbext_ct_MSForm Then
                Dim I As Long
                For I = 0 To FormsList.ListCount - 1
                    If FormsList.List(I) = IDE.SelectedVBComponent.Name Then
                        FormsList.Selected(0) = False
                        FormsList.Selected(I) = True
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub FormsList_Click()
    OKButton.Enabled = FormsList.SelCount > 0
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    
    Dim FolderName As String
    FolderName = AddBackslash(FolderText.Text)
    SaveSetting AddInName, SectionName, FolderKeyName, FolderName
    
    SaveSetting AddInName, SectionName, OverwriteKeyName, OverwriteCheck.Value

    Dim ExitLoop As Boolean
    ExitLoop = False

    Dim I As Long
    For I = 0 To FormsList.ListCount - 1
        If FormsList.Selected(I) Then
            Dim FormName As String
            FormName = FormsList.List(I)
            
            Dim Component As VBIDE.VBComponent
            For Each Component In IDE.ActiveVBProject.VBComponents
                If Component.Type = vbext_ct_MSForm Then
                    If Component.Name = FormName Then
                        Dim Exporter As ExportClass
                        Set Exporter = New ExportClass
                        ExitLoop = Not Exporter.ExportForm(IDE, Component, FolderName, OverwriteCheck.Value)
                        Set Exporter = Nothing
                        Exit For
                    End If
                End If
            Next
            
            If ExitLoop Then Exit For
        
            If Err <> 0 Then
                ErrorBoxEx "Cannot save " & FormName & " form."
                Err.Clear
            End If
        End If
    Next

    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub FormsList_DblClick()
    OKButton_Click
End Sub

Private Sub SelectAllButton_Click()
    SelectAll True
End Sub

Private Sub UnselectAllButton_Click()
    SelectAll False
End Sub

Private Sub SelectAll(Mode As Boolean)
    Dim I As Long
    For I = 0 To FormsList.ListCount - 1
        FormsList.Selected(I) = Mode
    Next
End Sub

Private Sub BrowseCommand_Click()
    FolderText.Text = BrowseForFolder(Me.hwnd, FolderText.Text, "Select folder to save exported form(s):")
End Sub

Private Sub AboutButton_Click()
    Dim Dlg As New AboutForm
    Dlg.Show 1
End Sub

