VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FormExp (export.cls)
' http://www.balagurov.com/software/formexp/

Option Explicit

Private IDE As VBIDE.VBE

Private File As FileClass

Public Function ExportForm(ByRef IDE_ As VBIDE.VBE, ByRef Component As VBIDE.VBComponent, _
        FolderName As String, Overwrite As Boolean) As Boolean
    
    On Error Resume Next

    Set IDE = IDE_
    
    ExportForm = False
    
    Dim Form As UserForm
    Set Form = Component.Designer
    
    Dim FormName As String
    FormName = Component.Name
    
    Dim FileName As String
    FileName = FolderName & FormName & ".frm"
    
    If Not Overwrite Then
        If PathFileExists(FileName) Then
            Select Case YesNoCancelBox("The following file already exists:" & vbNewLine & FileName & _
                    vbNewLine & vbNewLine & "Do you want to overwrite it?")
                Case vbNo
                    ExportForm = True
                    Exit Function
                Case vbCancel
                    Exit Function
            End Select
        End If
    End If
    
    Set File = New FileClass
    
    If Not File.CreateFile(FileName) Then
        ErrorBoxEx "Cannot create file: '" & FileName & "'"
        Exit Function
    End If
    
    File.WriteLine "VERSION 5.00"
    File.WriteLine "Begin VB.Form " & FormName
    
    File.WriteLine "   BackColor       =   " & ColorToStr(Form.BackColor)
    File.WriteLine "   BorderStyle     =   3  'Fixed Dialog"
    File.WriteLine "   Caption         =   """ & Component.Properties("Caption").Value & """"
    File.WriteLine "   ClientHeight    =   " & Form.InsideHeight * 20
    File.WriteLine "   ClientLeft      =   0"
    File.WriteLine "   ClientTop       =   0"
    File.WriteLine "   ClientWidth     =   " & Form.InsideWidth * 20
    File.WriteLine "   BeginProperty Font"
    File.WriteLine "      Name            =   """ & Form.Font.Name & """"
    File.WriteLine "      Size            =   " & Form.Font.Size
    File.WriteLine "      Charset         =   " & Form.Font.Charset
    File.WriteLine "      Weight          =   " & Form.Font.Weight
    File.WriteLine "      Underline       =   " & BoolToStr(Form.Font.Underline)
    File.WriteLine "      Italic          =   " & BoolToStr(Form.Font.Italic)
    File.WriteLine "      Strikethrough   =   " & BoolToStr(Form.Font.Strikethrough)
    File.WriteLine "   EndProperty"
    File.WriteLine "   ForeColor       =   " & ColorToStr(Form.ForeColor)
    File.WriteLine "   MaxButton       =   0   'False"
    File.WriteLine "   MinButton       =   0   'False"
    File.WriteLine "   ScaleHeight    =   " & Form.InsideHeight * 20
    File.WriteLine "   ScaleWidth     =   " & Form.InsideWidth * 20
    File.WriteLine "   ShowInTaskbar   =   0   'False"
    File.WriteLine "   StartUpPosition =   1  'CenterOwner"
    
    Dim Control As MSForms.Control
    For Each Control In Form.Controls
        If TypeName(Control.Parent) = "UserForm" Then
            If TypeName(Control) = "Frame" Then
                ExportFrame Control, 1
            Else
                ExportControl Control, 1
            End If
        End If
    Next
    
    File.WriteLine "End"
    
'    Dim P As VBIDE.Property
 '   For Each P In Component.Properties
  '      File.WriteLine P.Name
   '     File.WriteLine "   = " & P.Value
    'Next
    
    File.WriteLine "Attribute VB_Name = """ & FormName & """"
    File.WriteLine "Attribute VB_GlobalNameSpace = False"
    File.WriteLine "Attribute VB_Creatable = False"
    File.WriteLine "Attribute VB_PredeclaredId = True"
    File.WriteLine "Attribute VB_Exposed = False"
    
    Dim Code As String
    Code = Component.CodeModule.CountOfLines
    
    
    Dim Code As String
    Code = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
    Code = Replace(Code, "Private Sub UserForm_", "Private Sub Form_", 1, -1, vbTextCompare)
    File.WriteLine Code
    
    File.CloseFile

    Set File = Nothing
End Function

Private Sub ExportFrame(ByRef Frame As MSForms.Frame, Level As Long)
    On Error Resume Next

    ExportControl Frame, Level, True
    
    Dim Control As MSForms.Control
    For Each Control In Frame.Controls
        If Control.Parent.object Is Frame.object Then
            If TypeName(Control) = "Frame" Then
                ExportFrame Control, Level + 1
            Else
                ExportControl Control, Level + 1
            End If
        End If
    Next
            
    File.WriteLine Space(Level * 3) & "End"
End Sub

Private Sub ExportControl(ByRef Control As MSForms.Control, Level As Long, _
        Optional NoEnd As Boolean = False)
    
    On Error Resume Next
    
    Dim Margin As String
    Margin = Space(Level * 3)
    
    File.WriteLine Margin & "Begin " & GetControlProgID(Control) & " " & Control.Name
    
    Dim Caption As String
    Caption = Control.object.Caption
    If Err.Number = 0 Then
        Dim Acc As String
        Acc = Control.Accelerator
        If Err.Number = 0 Then Caption = Replace(Caption, Acc, "&" & Acc, 1, 1, vbBinaryCompare)
        File.WriteLine Margin & "   Caption         =   " & StringToStr(Caption)
    End If
    
    File.WriteLine Margin & "   BackColor       =   " & ColorToStr(Control.BackColor)
    If Not Control.Enabled Then _
        File.WriteLine Margin & "   Enabled         =   0   'False"
    If Not Control.TabStop And TypeName(Control) <> "CommandButton" Then _
        File.WriteLine Margin & "   ForeColor       =   " & ColorToStr(Control.ForeColor)
    File.WriteLine Margin & "   Height          =   " & Control.Height * 20
    If Control.HelpContextID > 0 Then _
        File.WriteLine Margin & "   HelpContextID   =   " & Control.HelpContextID
    File.WriteLine Margin & "   Left            =   " & Control.Left * 20
    File.WriteLine Margin & "   TabIndex        =   " & Control.TabIndex
    If Not Control.TabStop And TypeName(Control) <> "Label" Then _
        File.WriteLine Margin & "   TabStop         =   0   'False"
    If Len(Control.Tag) > 0 Then _
        File.WriteLine Margin & "   Tag             =   " & StringToStr(Control.Tag)
    If Control.TextAlign <> 1 Then
        Dim TextAlignArray As Variant
        TextAlignArray = Array(0, 2, 1)
        File.WriteLine Margin & "   Alignment       =   " & TextAlignArray(Control.TextAlign - 1)
    End If
    If Len(Control.ControlTipText) > 0 Then _
        File.WriteLine Margin & "   ToolTipText     =   " & StringToStr(Control.ControlTipText)
    File.WriteLine Margin & "   Top             =   " & Control.Top * 20
    If Not Control.Visible Then _
        File.WriteLine Margin & "   Visible         =   0   'False"
    File.WriteLine Margin & "   Width           =   " & Control.Width * 20
    
    Select Case TypeName(Control)
        Case "CommandButton"
            If Control.Cancel Then _
                File.WriteLine Margin & "   Cancel          =   -1  'True"
            If Control.Default Then _
                File.WriteLine Margin & "   Default         =   -1  'True"
        Case "Label"
            If Control.AutoSize Then _
                File.WriteLine Margin & "   AutoSize        =   -1  'True"
            If Control.BorderStyle <> 0 Then _
                File.WriteLine Margin & "   BorderStyle     =   " & Control.BorderStyle
            File.WriteLine Margin & "   WordWrap        =   " & BoolToStr(Control.WordWrap)
        Case "TextBox"
            If Not Control.HideSelection Then _
                File.WriteLine Margin & "   HideSelection   =   0  'False"
            If Control.Locked Then _
                File.WriteLine Margin & "   Locked          =   -1 'True"
            If Control.MaxLength > 0 Then _
                File.WriteLine Margin & "   MaxLength       =   " & Control.MaxLength
            If Control.MultiLine Then _
                File.WriteLine Margin & "   MultiLine       =   -1 'True"
            If Len(Control.PasswordChar) > 0 Then _
                File.WriteLine Margin & "   PasswordChar    =   " & StringToStr(Control.PasswordChar)
            If Control.ScrollBars <> 0 Then _
                File.WriteLine Margin & "   ScrollBars      =   " & Control.ScrollBars
            If Len(Control.Text) > 0 Then _
                File.WriteLine Margin & "   Text            =   " & StringToStr(Control.Text)
        Case "CheckBox"
            If Control.Value <> 0 Then _
                File.WriteLine Margin & "   Value           =   1"
        Case "OptionButton"
            If Control.Value <> 0 Then _
                File.WriteLine Margin & "   Value           =   -1 'True"
        Case "ListBox"
            If Not Control.IntegralHeight <> 0 Then _
                File.WriteLine Margin & "   IntegralHeight  =   0 'False"
            If Control.MultiSelect <> 0 Then _
                File.WriteLine Margin & "   MultiSelect     =   " & Control.MultiSelect
        Case "ComboBox"
            If Control.Locked Then _
                File.WriteLine Margin & "   Locked          =   -1 'True"
            If Control.Style <> 0 Then _
                File.WriteLine Margin & "   Style           =   " & Control.Style
    End Select
    
    If Not NoEnd Then File.WriteLine Margin & "End"

    Err.Clear
End Sub

Private Function GetControlProgID(ByRef Control As MSForms.Control)
    On Error Resume Next

    Dim ControlType As String
    ControlType = TypeName(Control)

    Const VBControls As String = _
        ".CheckBox.ComboBox.CommandButton.Frame.Image.Label.ListBox.OptionButton.TextBox."
    If InStr(1, VBControls, "." & ControlType & ".", vbTextCompare) > 0 Then
        GetControlProgID = "VB." & ControlType
        Exit Function
    End If
    
    If ControlType = "ScrollBar" Then
        Dim ScrollBar As MSForms.ScrollBar
        Set ScrollBar = Control
        If ScrollBar.Orientation = fmOrientationHorizontal Then
            GetControlProgID = "VB.HScrollBar"
        Else
            GetControlProgID = "VB.VScrollBar"
        End If
        Exit Function
    End If

    On Error GoTo OnError

    Dim Reference As VBIDE.Reference
    For Each Reference In IDE.ActiveVBProject.References
        If Not Reference.IsBroken Then
            Dim ProgID As String
            ProgID = Reference.Name & "." & ControlType
            
            Dim hKey As Long
            If RegOpenKey(HKEY_CLASSES_ROOT, ProgID, hKey) = 0 Then
                GetControlProgID = ProgID
                RegCloseKey hKey
                Exit Function
            End If
        End If
    Next

OnError:
    GetControlProgID = "<unknown>." & ControlType
    Err.Clear
End Function

Private Function BoolToStr(Value As Boolean)
    If Value Then
        BoolToStr = "-1"
    Else
        BoolToStr = "0"
    End If
End Function

Private Function ColorToStr(Value As OLE_COLOR)
    ColorToStr = Hex(Value)
    If Len(ColorToStr) < 8 Then ColorToStr = String(8 - Len(ColorToStr), "0")
    ColorToStr = "&H" & ColorToStr & "&"
End Function

Private Function StringToStr(Value As String)
    StringToStr = """" & Replace(Value, """", """""") & """"
End Function


