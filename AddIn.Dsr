VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} AddIn 
   ClientHeight    =   14160
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15900
   _ExtentX        =   28046
   _ExtentY        =   24977
   _Version        =   393216
   Description     =   "This add-in exports Microsoft Office 2000 forms and their associated code to Visual Basic .frm files."
   DisplayName     =   "Export Forms to Visual Basic"
   AppName         =   "Visual Basic for Applications IDE"
   AppVer          =   "6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0"
End
Attribute VB_Name = "AddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' FormExp (addin.dsr)
' http://www.balagurov.com/software/formexp/

Option Explicit

Public IDE As VBIDE.VBE
Public MenuItem As Office.CommandBarButton
Public WithEvents MenuItemHandler As VBIDE.CommandBarEvents
Attribute MenuItemHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error Resume Next

    Set IDE = Application
    
    Dim AddInsMenu As Office.CommandBar
    Set AddInsMenu = IDE.CommandBars("Add-Ins")
    
    If Not (AddInsMenu Is Nothing) Then
        Set MenuItem = AddInsMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = AddInName & "..."
    
        Set MenuItemHandler = IDE.Events.CommandBarEvents(MenuItem)
    End If
    
    'создание виртуального окна
    InitHK
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    MenuItem.Delete
End Sub

Private Sub MenuItemHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next

    Dim Dlg As New ChooseForm
    Dlg.Initialize IDE
    Dlg.Show 1
End Sub

