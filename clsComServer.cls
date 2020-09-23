VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelfReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API's to use
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0

Public Function RegisterServer(hWnd As Long, DllServerPath As String, bRegister As Boolean)
    Dim strFile As String
    
On Error Resume Next

    ' going to call the DllRegisterServer/DllUnRegisterServer API of the specified library.
    ' there's no need to use the Regsvr32.exe anymore.
    ' Make sure the path is correct and that the file exists, otherwise VB will crash.

    Dim lb As Long, pa As Long
    Dim strErrorText As String
    
    'Load the Library
    lb = LoadLibrary(DllServerPath)
    
    If bRegister Then
        'Get Address of the DllRegisterServer Method
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
    'Get Address of the DllUnregisterServer Method
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If
    
    'Get the File name without the path
    strFile = Mid(DllServerPath, InStrRev(DllServerPath, "\") + 1, Len(DllServerPath)) + " - "
    
    Dim strText As String
    
    'Exec the Method using the Handle
    CallWindowProc pa, hWnd, ByVal 0&, ByVal 0&, ByVal 0&
    
    'Update the Return Text
    If bRegister Then
        strText = strFile + IIf(pa > 0, "Registered", "Error")
    Else
        strText = strFile + IIf(pa > 0, "Unregistered", "Error")
    End If
    
    'Update the ListView
    If pa > 0 Then
        frmMain.ListView1.ListItems.Add , , strText, , 1
    Else
        frmMain.ListView1.ListItems.Add , , strText, , 2
    End If
    
    'unmap the library's address
    FreeLibrary lb

End Function





