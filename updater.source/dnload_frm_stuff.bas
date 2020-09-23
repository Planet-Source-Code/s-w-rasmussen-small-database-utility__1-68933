Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const SPI_GETWORKAREA = 48                       'Desktop Area with task bar consideration.
Public Const WM_CLOSE = &H10

Public Type FILETIME
   dwLowDateTime            As Long
   dwHighDateTime           As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes         As Long
   ftCreationTime           As FILETIME
   ftLastAccessTime         As FILETIME
   ftLastWriteTime          As FILETIME
   nFileSizeHigh            As Long
   nFileSizeLow             As Long
   dwReserved0              As Long
   dwReserved1              As Long
   cFileName                As String * 260
   cAlternate               As String * 14
End Type

Public Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Public Type WIN_VERSION
    WindowsVersion  As String
    UseAutoLaunch   As Boolean
End Type
Public WinVer As WIN_VERSION

Public Const INFINITE = &HFFFF

'STARTINFO constants
Private Const STARTF_USESHOWWINDOW = &H1
Public Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
 
Public Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
 



'------------------------------------------------------------------------------
'http://www.vb6.us/tutorials/advanced-shell
'------------------------------------------------------------------------------
Public Function SuperShell( _
                            ByVal App As String, _
                            ByVal WorkDir As String, _
                            dwMilliseconds As Long, _
                            ByVal start_size As enSW, _
                            ByVal Priority_Class As enPriority_Class) _
                            As Boolean

Dim pclass As Long
Dim sinfo As STARTUPINFO
Dim pinfo As PROCESS_INFORMATION 'Not used, but needed
Dim sec1 As SECURITY_ATTRIBUTES
Dim sec2 As SECURITY_ATTRIBUTES
    
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sinfo.cb = Len(sinfo)
    
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    sinfo.wShowWindow = start_size
    
    pclass = Priority_Class
    
    If CreateProcess( _
                      vbNullString, _
                      App, _
                      sec1, _
                      sec2, _
                      False, _
                      pclass, _
                      0&, _
                      WorkDir, _
                      sinfo, _
                      pinfo) Then
        WaitForSingleObject pinfo.hProcess, dwMilliseconds
        SuperShell = True
    Else
        SuperShell = False
    End If

End Function

Public Function FileExist(ByVal FilePath As String) As Boolean
        
On Error GoTo errhandler

Dim hFile                   As Long
Dim WFD                     As WIN32_FIND_DATA

    hFile = FindFirstFile(FilePath, WFD)
    FileExist = hFile <> -1
   
    Call FindClose(hFile)
    
    Exit Function
    
errhandler:
    Exit Function
End Function

