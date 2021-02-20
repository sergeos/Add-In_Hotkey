Attribute VB_Name = "Common"
' FormExp (common.bas)
' http://www.balagurov.com/software/formexp/

Option Explicit

Public Const AddInName As String = "Export Forms to Visual Basic"

Private Const WM_USER = &H400

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Public Const CSIDL_PERSONAL = &H5 ' My Documents
Public Const CSIDL_DESKTOPDIRECTORY = &H10 ' Desktop

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
   (ByVal HWndOwner As Long, ByVal Folder As Long, pidl As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal Path As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
  (bi As BROWSEINFO) As Long
  
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal p As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Dest As Any, Source As Any, ByVal Length As Long)
    
Private Declare Function LocalAlloc Lib "kernel32" _
    (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
   
Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" _
    (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  
Private Function PidlToPath(pidl As Long, Optional DefaultFolder As String = "") As String
        PidlToPath = DefaultFolder
        
        If pidl = 0 Then Exit Function

        Dim Path As String
        Path = Space(1024)
        
        If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
            PidlToPath = Left(Path, InStr(Path, Chr(0)) - 1)
        End If
    
        Call CoTaskMemFree(pidl)
End Function

Public Function GetSpecialFolderLocation(CSIDL As Long, Optional HWndOwner As Long = 0) As String
    Dim Path As String
    Dim pidl As Long
    
    If SHGetSpecialFolderLocation(HWndOwner, CSIDL, pidl) = 0 Then  ' S_OK
        GetSpecialFolderLocation = PidlToPath(pidl)
    End If
End Function

Public Function BrowseCallbackProc( _
        ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    
    If uMsg = BFFM_INITIALIZED Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If
End Function

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

Public Function BrowseForFolder(HWndOwner As Long, Optional InitialFolder As String = "", _
        Optional Description As String = "") As String
    
    Dim mem As Long
    mem = LocalAlloc(LMEM_FIXED Or LMEM_ZEROINIT, Len(InitialFolder) + 1)
    CopyMemory ByVal mem, ByVal InitialFolder, Len(InitialFolder) + 1

    Dim bi As BROWSEINFO
    With bi
        .hOwner = HWndOwner
        .pidlRoot = 0&
        .lpszTitle = Description
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        .lParam = mem
    End With
    
    Dim pidl As Long
    pidl = SHBrowseForFolder(bi)
    
    BrowseForFolder = PidlToPath(pidl, InitialFolder)
    
    Call LocalFree(mem)
End Function

Public Function AddBackslash(Path As String) As String
    AddBackslash = Path
    If Len(Path) = 0 Then Exit Function
    If Right(Path, 1) <> "\" Then AddBackslash = Path & "\"
End Function

Private Function MsgBoxEx(ByVal Text As String, Buttons As VbMsgBoxStyle) As VbMsgBoxResult
    MsgBoxEx = MsgBox(Text, Buttons, AddInName)
End Function

Public Sub ErrorBox(ByVal Text As String)
    MsgBoxEx Text, vbOKOnly + vbCritical
End Sub

Public Sub ErrorBoxEx(ByVal Text As String)
    ErrorBox Text & vbNewLine & vbNewLine & "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function YesNoCancelBox(ByVal Text As String)
    YesNoCancelBox = MsgBoxEx(Text, vbYesNoCancel + vbExclamation)
End Function
