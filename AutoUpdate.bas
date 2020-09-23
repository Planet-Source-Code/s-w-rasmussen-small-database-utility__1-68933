Attribute VB_Name = "AutoUpdate"
Option Explicit

Public Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Public Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Public Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData As Any, lpcbData&)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Public Const BINDF_GETNEWESTVERSION As Long = &H10
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CURRENT_USER = &H80000001
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const REG_SZ = 1

Public NewestVersionOnWebFormatted_G    As String
Public NewestVersionOnWebRaw_G          As Long
Public UpdateToNewVersion_G             As Boolean

'------------------------------------------------------------------------------
' get version number from website:
' www.sonderskovhjemmet.dk/intranet/phonelist/sdu_version.txt
' 25. july 2011, swr
'------------------------------------------------------------------------------
Public Function GetVersionFromWeb() As Boolean

On Error Resume Next

Dim fdata                   As String
Dim CurrentVersionRaw       As String
    
    NewestVersionOnWebRaw_G = Val(App.Major & App.Minor & Format$(App.Revision, "000"))
    
    ' url to update84.log
    If Not WebServerConnectionOK_G Then
        Main.StatusBar1.Panels.Item(3).Text = " No connection to webserver..." '
        GoTo errhandler
    End If
            
    ' get sdu.log from download web page: http://www.swr.dk/software/downloads
    Main.Inet1.AccessType = icUseDefault
    Main.Inet1.RequestTimeout = 10
        
    fdata = Main.Inet1.OpenURL(uploaddb.UpdateExeURL & "sdu_version.txt", icString)
                        
    ' display 'unknown' if retrieval of version number of newest version for download fails
    If Len(fdata) = 5 And IsNumeric(fdata) Then
        NewestVersionOnWebFormatted_G = Left$(fdata, 1) & "." & Mid$(fdata, 2, 1) & "." & Mid$(fdata, 3)
        NewestVersionOnWebRaw_G = Val(fdata)
        GetVersionFromWeb = True
        Exit Function
    Else
        GoTo errhandler
    End If
        
    Exit Function

errhandler:
    NewestVersionOnWebFormatted_G = "not available"
    NewestVersionOnWebRaw_G = Val(CurrentVersionRaw)
    GetVersionFromWeb = False
    Exit Function
    
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
