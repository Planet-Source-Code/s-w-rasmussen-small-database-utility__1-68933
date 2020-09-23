VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpdate 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "frmUpdate.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Max             =   1
   End
   Begin VB.Label lblUpdater 
      BackColor       =   &H00E0E0E0&
      Caption         =   "auto updater"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   900
      TabIndex        =   5
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Initialising update process..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   960
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label InfoLine 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Installing update, please wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008B5641&
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   1455
      Width           =   2295
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Private Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Private Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData As Any, lpcbData&)
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CURRENT_USER = &H80000001
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_SZ = 1

' setup information for updater program
Private Const ExeFile = "sdu_uk.exe"                                                ' Name of exe file to be updated
Private Const BakFile = "sdu_uk.bak"                                                ' Name of the old exe file
Private Const DownloadURL = "http://www.swr.dk/software/downloads/sdu_uk.exe"       ' URL containing update
Private Const MainAppClosingCaption = " Small Database Utility is closing..."       ' Form caption of the closing calling application
Private Const MainAppTitle = "Small Database Utility"                               ' Titel of the calling application

Private ApplicationPath_ExeFile     As String
Private DesktopPath_ExeFile         As String

Private ApplicationPath_BakFile     As String
Private SuccessAll                  As Boolean

Private WithEvents mydl             As VicsDL
Attribute mydl.VB_VarHelpID = -1





'------------------------------------------------------------------------------
' Download exe from remote to desktop
'------------------------------------------------------------------------------
Private Function B_DownloadExeFile() As Boolean
    
On Error GoTo errhandler

Dim filelist                As String

    filelist = DownloadURL & "," & DesktopPath_ExeFile & "," & "1": WaitABit 0.5
    Call ShowDownLoad(filelist, Me)
    
    B_DownloadExeFile = True
    
    Exit Function
    
errhandler:
    B_DownloadExeFile = False
    Exit Function
End Function

'------------------------------------------------------------------------------
' In application folder: kill old bak -> copy exe to bak -> kill exe
'------------------------------------------------------------------------------
Private Function C_RemoveOldExeFile() As Boolean

On Error GoTo errhandler
    
    If FileExist(ApplicationPath_BakFile) Then Kill ApplicationPath_BakFile: WaitABit 0.5
    If FileExist(ApplicationPath_ExeFile) Then FileCopy ApplicationPath_ExeFile, ApplicationPath_BakFile: WaitABit 0.5
    If FileExist(ApplicationPath_ExeFile) Then Kill ApplicationPath_ExeFile: WaitABit 0.5
            
    C_RemoveOldExeFile = True
    
    Exit Function
    
errhandler:
    C_RemoveOldExeFile = False
    Exit Function
End Function


'------------------------------------------------------------------------------
' On desktop: copy exe to app -> kill exe on desktop
'------------------------------------------------------------------------------
Private Function D_ReplaceExeFile() As Boolean

On Error GoTo errhandler

    If FileExist(DesktopPath_ExeFile) Then FileCopy DesktopPath_ExeFile, ApplicationPath_ExeFile: WaitABit 0.5
    If FileExist(DesktopPath_ExeFile) Then Kill DesktopPath_ExeFile: WaitABit 0.5
        
    D_ReplaceExeFile = True
    
    Exit Function
    
errhandler:
    D_ReplaceExeFile = False
        Exit Function
End Function

'------------------------------------------------------------------------------
' Wait for 'Seconds' with a Doevents every 10 milliseconds. Minimum waiting
' time = 0.01 second
' Created 10. August 2011, swr
'------------------------------------------------------------------------------
Public Sub WaitABit(ByVal Seconds As Single)

On Error GoTo errhandler

Dim Start                   As Single
    
    Seconds = Seconds * 100
    Start = Timer * 100
    Do While Timer * 100 - Start <= Seconds
        Sleep (10)                         ' Sleep for 10 milliseconds between Doevents
        DoEvents
    Loop
    
errhandler:
    Exit Sub
End Sub

Private Function GetWindowsVersion(WinVer As WIN_VERSION) As Boolean

Dim osv As OSVERSIONINFO

    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                WinVer.WindowsVersion = "Win32s on Windows 3.1"
                WinVer.UseAutoLaunch = True
            Case VER_PLATFORM_WIN32_NT
                WinVer.WindowsVersion = "Windows NT"
                WinVer.UseAutoLaunch = True
                Select Case osv.dwVerMajor
                    Case 3
                        WinVer.WindowsVersion = "Windows NT 3.5"
                        WinVer.UseAutoLaunch = True
                    Case 4
                        WinVer.WindowsVersion = "Windows NT 4.0"
                        WinVer.UseAutoLaunch = True
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                WinVer.WindowsVersion = "Windows 2000"
                                WinVer.UseAutoLaunch = True
                            Case 1
                                WinVer.WindowsVersion = "Windows XP"
                                WinVer.UseAutoLaunch = True
                            Case 2
                                WinVer.WindowsVersion = "Windows Server 2003"
                                WinVer.UseAutoLaunch = True
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                WinVer.WindowsVersion = "Windows Vista/Server 2008"
                                WinVer.UseAutoLaunch = False
                            Case 1
                                WinVer.WindowsVersion = "Windows 7/Server 2008 R2"
                                WinVer.UseAutoLaunch = False
                        End Select
                End Select
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        WinVer.WindowsVersion = "Windows 95"
                        WinVer.UseAutoLaunch = True
                    Case 90
                        WinVer.WindowsVersion = "Windows Me"
                        WinVer.UseAutoLaunch = True
                    Case Else
                        WinVer.WindowsVersion = "Windows 98"
                        WinVer.UseAutoLaunch = True
                End Select
        End Select
    Else
        WinVer.WindowsVersion = "Unknown Windows version"
        WinVer.UseAutoLaunch = False
    End If

End Function


Public Function ShowDownLoad(filelist As String, CallingForm As Form)
    
    Set mydl = New VicsDL
        
    CallingForm.SetFocus
    DoEvents
    
    'split files to download from FileList
    Dim i
    Dim x As Integer
    Dim File2DownLoad As String
    Dim File2Save As String
    Dim DeleteCache As Boolean
    Dim TopLimit As Integer
    Dim TempDelete As String
    Dim OffSet As Integer
    
    i = Split(filelist, ",")
    
    TopLimit = (UBound(i) - 2) / 3 'filelist comes in as:File2DownLoad,File2Save,DeleteCache
    OffSet = 0
    For x = 0 To TopLimit
        File2DownLoad = i(OffSet)
        File2Save = i(OffSet + 1)
        TempDelete = i(OffSet + 2)
        If TempDelete = "1" Then
            DeleteCache = True
        Else
            DeleteCache = False
        End If
        'OffSet = OffSet + 3 'increment the offset for next file
        frmUpdate.ProgressBar1.Value = 0 'initialize the progress bar
        
        If DeleteCache Then
            If mydl.DeleteVicCache(File2DownLoad) = 1 Then 'file was found and deletedLf
            Else
            End If
        End If
        frmUpdate.Label1.Caption = File2DownLoad
        
        'proceed with the download part
        If mydl.DownloadSuccess(File2DownLoad, File2Save) Then
            ShowDownLoad = True
        Else
            ShowDownLoad = False
        End If
    Next
    
BailingOut:
    Set mydl = Nothing          'free memory
    
End Function
Public Function GetDesktopPath() As String

On Error GoTo errhandler

Const nLG                   As Long = 256
Dim sValue                  As String * nLG
Dim hKey                    As Long
Dim nType                   As Long
Dim nCR                     As Long

    If (RegOpenKeyExA(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        
        If (RegQueryValueExA(hKey, "Desktop", 0, nType, ByVal sValue, nLG) = ERROR_SUCCESS) Then
            If (nType = REG_SZ) Then
                GetDesktopPath = Left(sValue, InStr(sValue, vbNullChar) - 1)
            End If
        End If

        nCR = RegCloseKey(hKey)
    
    End If
    
errhandler:

End Function

'------------------------------------------------------------------------------
' Wait for program exit
'------------------------------------------------------------------------------
Private Function A_WaitForProgramExit() As Boolean

On Error GoTo errhandler
        
Dim hWnd                    As Long
Dim Result                  As Long

    hWnd = 100
    Do While hWnd > 0
        hWnd = FindWindow(vbNullString, MainAppClosingCaption)
        If hWnd > 0 Then
            Result = SendMessage(hWnd, &H10, 0, 0)
        End If
        WaitABit 0.1
    Loop
    
    A_WaitForProgramExit = True
    
    Exit Function
    
errhandler:
    A_WaitForProgramExit = False
End Function

Private Sub Form_Activate()

On Error GoTo errhandler

    SuccessAll = SuccessAll + A_WaitForProgramExit
    
    SuccessAll = SuccessAll + B_DownloadExeFile
            
    SuccessAll = SuccessAll + C_RemoveOldExeFile
    
    SuccessAll = SuccessAll + D_ReplaceExeFile
    
    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()
'------------------------------------------------------------------------------
' Allround updater program to replace file(s) in an application folder with
' file(s) with identical file names located on a webserver.
'
' The updater is started by the program to be updated in its unload event
' immediately before it is unloaded (controlled by the app to be updated).
'
' The update file(s) are downloaded from the website and is placed on the
' desktop.
'
' Then - after the program to be updated has terminated - the downloaded
' file(s) is copied to the application folder replacing the old file(s).
'
' Created 2. august 2011, swr
'------------------------------------------------------------------------------
On Error Resume Next

    Me.Left = Screen.Width - Me.Width - 1200
    Me.Top = Screen.Height - Me.Height - 1200
            
    Me.Caption = "auto updater"
    lblTitle.Caption = MainAppTitle
    Me.Show
        
    ' app path is the same as sdu_uk.exe
    ApplicationPath_ExeFile = App.Path & "\" & ExeFile
    ApplicationPath_BakFile = App.Path & "\" & BakFile
    DesktopPath_ExeFile = GetDesktopPath & "\" & ExeFile
    
End Sub

'------------------------------------------------------------------------------
' Call ShellExecute(0&, vbNullString, ApplicationPath_BatFile, vbNullString, vbNullString, vbHidden)
'
' ShellExecute does not work on Windows 7. This implies that the user must
' restart SDU manually after updating by double-clicking the desktop icon.
' created 18. august 2011, swr
'------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo errhandler

Dim Success                 As Boolean
        
    ' calculate success - and inform user...
    If SuccessAll Then
        InfoLine.Caption = MainAppTitle & " successfully updated...": WaitABit 2
    Else
        InfoLine.Caption = MainAppTitle & " update failed...":  WaitABit 2
        GoTo JustUnload
    End If
    
    Call GetWindowsVersion(WinVer)
    
    If WinVer.UseAutoLaunch Then
        'Success = SuperShell(App.Path & "\sdu_uk.exe", App.Path, 0, SW_NORMAL, NORMAL_PRIORITY_CLASS)
        'Success = Shell(App.Path & "\" & "sdu_uk.exe", vbNormalFocus)
        Success = ShellExecute(0&, vbNullString, ApplicationPath_ExeFile, vbNullString, vbNullString, vbNormal)
    Else
        Success = False
    End If
    
    If Success = False Then
        Me.Hide
        frmCompleted.lblTitle.Caption = "Update Complete !"
        frmCompleted.lblInstructions.Caption = "Unable to auto-start the program after updating!" & _
                                                    vbCrLf & vbCrLf & _
                                                    "Please double-click the desktop icon to manually run the updated version of " & MainAppTitle & "..."
        frmCompleted.Show 1
    End If
    
    Unload frmCompleted
    DoEvents
    
JustUnload:
    
errhandler:
    Exit Sub
End Sub

Private Sub mydl_VicDLProg(ByVal VicBytesIn As Long, ByVal VicTotalBytes As Long)

On Error GoTo OhCrap

    If VicBytesIn >= 0 And VicBytesIn <= VicTotalBytes Then
        frmUpdate.ProgressBar1.Max = VicTotalBytes              ' set/re-set the progress bar's max value after it is known for sure
        frmUpdate.ProgressBar1.Value = VicBytesIn               ' set the current level of progress
        DoEvents                                                ' force a refresh...
    End If
    
Exit Sub

OhCrap:
    Resume Next
End Sub


Private Sub Picture1_Click()

    Unload Me
    End
    
End Sub



