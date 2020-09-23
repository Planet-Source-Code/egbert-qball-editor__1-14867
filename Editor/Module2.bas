Attribute VB_Name = "Commandialog"
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private OFN As OPENFILENAME
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Function ShowOpen(Filter As String, DialogTitel As String) As String
Dim RetValue As Long
Dim iDelim As Integer
InitOFN Filter, DialogTitel
RetValue = GetOpenFileName(OFN)
If RetValue > 0 Then
iDelim = InStr(OFN.lpstrFile, vbNullChar)
If iDelim Then ShowOpen = Left$(OFN.lpstrFile, iDelim - 1)
Else
ShowOpen = ""
End If
End Function

Function ShowSave(Filter As String, DialogTitel As String) As String
Dim RetValue As Long
Dim iDelim As Integer
InitOFN Filter, DialogTitel
RetValue = GetSaveFileName(OFN)
If RetValue > 0 Then
iDelim = InStr(OFN.lpstrFile, vbNullChar)
If iDelim Then ShowSave = Left$(OFN.lpstrFile, iDelim - 1)
iDelim = InStr(Format(ShowSave, "<"), ".Qba")
If iDelim = 0 Then ShowSave = ShowSave & ".Qba"
Else
ShowSave = ""
End If
End Function

Private Sub InitOFN(Filter As String, DialogTitle As String)
Dim sTemp As String, I As Integer
Dim uFlag, mFlags As Long
With OFN
.lStructSize = Len(OFN)
.flags = uFlag
sTemp = "c:\mp3"
.lpstrInitialDir = sTemp
.lpstrFile = "" & String$(255 - Len(sTemp), 0)
.nMaxFile = 255
.lpstrFileTitle = String$(255, 0)
.nMaxFileTitle = 255
sTemp = Filter
For I = 1 To Len(sTemp)
If Mid(sTemp, I, 1) = "|" Then
Mid(sTemp, I, 1) = vbNullChar
End If
Next
sTemp = sTemp & String$(2, 0)
.lpstrFilter = sTemp
.lpstrTitle = DialogTitle
.hInstance = App.hInstance
End With
End Sub

Function CreatExtension() ''creats edit in menu right click
Dim Path As String
Dim sKeyName As String
Dim sKeyValue As String
Dim Ret&
Dim lphKey&

Path = App.Path
If Right(Path, 1) <> "\" Then Path = Path & "\"

sKeyName = "QBall Level File"
sKeyValue = "QBall Level File"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
sKeyName = ".Qba"
sKeyValue = "QBall Level File"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
sKeyName = "QBall Level File"
sKeyValue = Path & "Editor.exe %1"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "shell\Edit\command", REG_SZ, sKeyValue, MAX_PATH)
Ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, Path & "Editor.exe", MAX_PATH)
End Function
