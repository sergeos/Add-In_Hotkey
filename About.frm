VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CloseCommand 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1343
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label VersionLabel 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label WwwLabel 
      Alignment       =   2  'Center
      Caption         =   "http://www.balagurov.com/software/formexp/"
      Height          =   255
      Left            =   120
      MouseIcon       =   "About.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label EmailLabel 
      Alignment       =   2  'Center
      Caption         =   "vassili@balagurov.com"
      Height          =   255
      Left            =   120
      MouseIcon       =   "About.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label CopyrightLabel 
      Alignment       =   2  'Center
      Caption         =   "Copyright (c) 2002 Vassili Balagurov"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label AppNameLabel 
      Alignment       =   2  'Center
      Caption         =   "This add-in exports Microsoft Office 2000 forms and their associated code to Visual Basic .frm files."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FormExp (about.frm)
' http://www.balagurov.com/software/formexp/

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    VersionLabel.Caption = "FormExp Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub CloseCommand_Click()
    Unload Me
End Sub

Private Sub EmailLabel_Click()
    OpenLink "mailto:" & EmailLabel.Caption
End Sub

Private Sub WwwLabel_Click()
    OpenLink WwwLabel.Caption
End Sub

Private Sub OpenLink(Link As String)
    ShellExecute Me.hwnd, "open", Link, vbNullString, vbNullString, 1
End Sub
