Attribute VB_Name = "AppGeneral"
Option Explicit

Option Compare Text

Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long)
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FtpPutFile Lib "Wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare Function FtpGetFile Lib "Wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwLocalFlagsAndAttributes As Long, ByVal dwInternetFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare Function FtpFindFirstFile Lib "wininet" Alias "FtpFindFirstFileA" (ByVal hConnect As Long, ByVal lpszSearchFile As String, lpFindFileData As Any, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function FtpSetCurrentDirectory Lib "Wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function InternetOpen Lib "Wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetConnect Lib "Wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "Wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetAttemptConnect Lib "wininet" (ByVal dwReserved As Long) As Long
Public Declare Function InternetGetConnectedState Lib "Wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function FtpGetCurrentDirectory Lib "Wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Public Declare Function InternetGetLastResponseInfo Lib "Wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function compress Lib "ZLIB.DLL" (ByVal compr As String, comprLen As Any, ByVal buf As String, ByVal buflen As Long) As Long
Public Declare Function uncompress Lib "ZLIB.DLL" (ByVal uncompr As String, uncomprLen As Any, ByVal compr As String, ByVal lcompr As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Const vbDot = 46
Public Const MAXDWORD As Long = &HFFFFFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   count As Long
End Type

Public oIE7                    As Object

Public Type UPLOAD_SETTINGS
    WebsiteURL              As String
    RemoteServerIP          As String
    UserName                As String
    PassWord                As String
    RemoteFileName          As String
    RemoteFolderPath        As String
    LocalFileName           As String
    ProgramInfoURL          As String
    UpdateExeURL            As String
End Type
Public uploaddb             As UPLOAD_SETTINGS

Public Type BrowseInfo
   hOwner                   As Long
   pIDLRoot                 As Long
   pszDisplayName           As String
   lpszTitle                As String
   ulFlags                  As Long
   lpfn                     As Long
   lParam                   As Long
   iImage                   As Long
End Type

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

Public Type DATABASES_INFO
    TotalNumber             As Long
    LastUsedDB              As String
    FirstDB                 As String
End Type
Public Databases            As DATABASES_INFO

Public Type DATABASES_COMPARE
    HeadersIdentical        As Boolean
    RecordsIdentical        As Boolean
    NumLocalRecords         As Long
    NumRemoteRecords        As Long
    NumUpdatedRecords       As Long
End Type
Public DBCompare            As DATABASES_COMPARE

' Folders
Public MAIN_DIR_G               As String       ' main data folder
Public NOTES_DIR_G              As String       ' notes folder
Public EXCEL_DIR_G              As String       ' excel folder
Public BACKUP_DIR_G             As String       ' backup folder
Public IMGS_DIR_G               As String       ' images folder

Public WebServerConnectionOK_G  As Boolean
Public Configuration_Is_Dirty_G As Boolean
Public NumRecords_G             As Long
Public CurrRecord_G             As Long
Public BookMark_G               As Long
Public Password_G               As String
Public ValidPasswords_G         As String
Public MasterUser_G             As Boolean
Public Loaded_G                 As Boolean
Public ImgIndex_G               As String
Public FromSearch_G             As Boolean
Public FromIncomplete_G         As Boolean
Public AllKeyWordsChecked_G     As Boolean
Public GotDatabaseFromWeb_G     As Boolean
Public NewDatabaseSuccess_G     As Boolean
Public sr()                     As Long         ' search results

Public Type RECORD_DATA
    ID                      As String       ' Record ID, value 19 char long string
    txtField(1 To 27)       As String       ' Text Fields (1-26), values = string
    chkKeyWord(1 To 10)     As Long         ' KeyWords (1-10), values = 0 or 1
    Comments                As String       ' Comments, value string
End Type

Public nr() As RECORD_DATA

Public Const swXORKEY = "ghj56XZkljhrKLJ78MM2093=)((/ghjælkjyutrwq+09=)/&f5ff234¤¤#432876(((&76H8L7609?0nm,><>\£$ty$3&/(%¤#HGFhgflplpoiyy"
Public Const swDot = "."

' upload
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

' tray icon
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONCLK = &H202

Public Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uID As Long
       uFlags As Long
       uCallbackMessage As Long
       hIcon As Long
       szTip As String * 64
End Type

Public IconData As NOTIFYICONDATA

Public Const SDU_STATUS_EDIT_OFF = "IS OFF"
Public Const SDU_STATUS_EDIT_ON = "IS ON"
Public Const SDU_STATUS_WORKING = " working..."
Public Const SDU_STATUS_ERROR = " Error..."
Public Const SDU_STATUS_COLOR_BLACK = &H808080
Public Const SDU_STATUS_COLOR_RED = &H40C0&
Public Const SDU_STATUS_COLOR_GREEN = &H8000&

Public Const vbPut = "PUT"
Public Const vbGet = "GET"

Public captions()           As String
Public SDU()                As String
Public lst()                As Long
Public users()              As String
Public selected()           As String
Public app_LastOpen()       As String
Public FolderList()         As String

Public cmp1()               As String
Public cmp2()               As String
Public new2()               As Long
Public diff2()              As String

Public TitleFontName_G      As String
Public TitleFontSize_G      As Long
Public TitleColor_G         As Long
Public TitleBold_G          As Long
Public TitleItalic_G        As Long
Public TitleOpaque_G        As Long
Public TitleBackcolor_G     As Long
Public WebLockStatus_G      As String
Public KeyWordLabelIndex_G  As Long


Public Function CreateNewLocalFolder(ByVal OldFolderName As String, ByVal NewFolderName As String) As Boolean
    
On Error GoTo errhandler

Dim ExistOld                As Boolean
Dim ExistNew                As Boolean
    
    ' remove terminal backslash from both folder names
    If Right$(OldFolderName, 1) = "\" Then OldFolderName = Left$(OldFolderName, Len(OldFolderName) - 1)
    If Right$(NewFolderName, 1) = "\" Then NewFolderName = Left$(NewFolderName, Len(NewFolderName) - 1)
    
    ' see if both old and new folder exist
    ExistOld = Len(Dir$(OldFolderName, vbDirectory)) > 0
    
    ' if NewFolderName already exists on C-Drive then exit - else create it
    If Len(Dir$(NewFolderName, vbDirectory)) > 0 Then
        GoTo errhandler
    Else
        MkDir NewFolderName
    End If
    
    ' verify that ExistNew exists
    ExistNew = Len(Dir$(NewFolderName, vbDirectory)) > 0
    
    If ExistOld And ExistNew Then
        Name OldFolderName As NewFolderName
        CreateNewLocalFolder = True
        Exit Function
    End If
    
errhandler:
    CreateNewLocalFolder = False
    Exit Function
End Function

'------------------------------------------------------------------------------
' routine to read, write and delete Registry entries
' lResult = fReadValue  ("HKCU", "Software\YourKey\LastKey\YourApp", "AppName", "S", "", sValue)
' lResult = fWriteValue ("HKCU", "Software\YourKey\LastKey\YourApp", "AppName", "S", "MyApp")
' lResult = fDeleteValue("HKCU", "Software\YourKey\LastKey\YourApp", "EntryToDelete")
' see: http://www.thescarms.com/VBasic/registry.aspx
'------------------------------------------------------------------------------
Public Function GetRunSettingsFromRegistry(ByVal ReadResident0_SetResident1_SetRunOnce2 As Long) As Boolean
    
On Error GoTo errhandler
    
Dim sValue                  As String
Dim lReturn                 As Long
Dim RegPath                 As String

Const RegValue = "C:\Progra~1\SDU_UK\sdu_uk.exe /resident"
Const RegName = "SmallDatabaseUtility"
    
    ' to avoid McAfee triggering a Trojan alert!
    RegPath = Join(Array("SOFTWARE", "Microsoft", "Windows", "CurrentVersion", "Run"), "\")

    Select Case ReadResident0_SetResident1_SetRunOnce2
        Case 0                                                              ' Read registry entry
            lReturn = fReadValue("HKLM", RegPath, RegName, "S", "", sValue)
            If sValue = RegValue Then GetRunSettingsFromRegistry = True Else GetRunSettingsFromRegistry = False
            Exit Function
    
        Case 1                                                              ' Write registry entry
            lReturn = fWriteValue("HKLM", RegPath, RegName, "S", RegValue)
            If lReturn = 0 Then GetRunSettingsFromRegistry = True Else GetRunSettingsFromRegistry = False
            Exit Function
            
        Case 2                                                              ' Delete registry entry
            lReturn = fDeleteValue("HKLM", RegPath, RegName)
            If lReturn = 0 Then GetRunSettingsFromRegistry = True Else GetRunSettingsFromRegistry = False
            Exit Function
            
    End Select
        
    Exit Function
    
errhandler:
    GetRunSettingsFromRegistry = False
    Exit Function
End Function



Public Sub Database_EraseA()

On Error GoTo errhandler

Dim N                       As Long

On Error Resume Next

    ' redeclare arrays
    ReDim nr(1 To 1) As RECORD_DATA             ' clear records array: nr()
    ReDim sr(1 To 2, 1 To 1)                    ' clear search array: sr()
    ReDim captions(1 To 80)                     ' clear captions array: captions()
    ReDim users(1 To 2, 1 To 6)                 ' clear administrator array: users()
    ReDim selected(1 To 2, 1 To 37)             ' clear selection on build list form: selected()
    'ReDim app_LastOpen(1 To 3, 1 To 10)        ????!!!!
    
    ' clear labels on main form
    For N = 1 To 26
        Main.lblField(N).Caption = vbNullString
    Next N
    For N = 28 To 37
        Main.lblField(N).Caption = vbNullString
    Next N
    
    For N = 1 To 7
        Main.lblCaption(N).Caption = vbNullString
    Next N
    
    With uploaddb
        .WebsiteURL = vbNullString          ' clear upload information: upload.
        .RemoteServerIP = vbNullString
        .UserName = vbNullString
        .PassWord = vbNullString
        .RemoteFileName = vbNullString
        .RemoteFolderPath = vbNullString
        .LocalFileName = vbNullString
        .ProgramInfoURL = vbNullString
        .UpdateExeURL = vbNullString
    End With
    
    NumRecords_G = 1                            ' set number of records, NumRecords_G, to 1
    CurrRecord_G = 1                            ' set current record number, CurrRecord_G, to 1
        
    captions(45) = "Empty Database"             ' set database title to: "Empty Database"
    captions(46) = "no records"                 ' set database subtitle to: "no records"
    
    Call Record_ClearCurrentA                   ' clear Main form
    
    IMGS_DIR_G = App.Path & "\imgs\"            ' set path to images in app folder
    
    DoEvents
    
errhandler:
    Exit Sub
End Sub

Public Function IsFormLoaded(ByVal frmName As String) As Boolean

Dim frm As Form

    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    
    IsFormLoaded = False
    
End Function

Function UpdateRecordForm(ByVal RecordNumber As Long, Optional UpdateLabels As Boolean = False) As Boolean

On Error GoTo errhandler

Dim N                       As Long
    
    If UpdateLabels Then
        ' field labels
        For N = 1 To 37
            Record.lblField(N).Caption = Main.lblField(N).Caption
        Next N
        For N = 1 To 7
            Record.lblCaption(N).Caption = Main.lblCaption(N).Caption
        Next N
    End If
    
    ' data fields
    For N = 1 To 26
        Record.txtField(N).Text = Main.txtField(N).Text
    Next N
    
    ' checkbox values
    For N = 1 To 10
        Record.chkKeyWord(N).Value = Main.chkKeyWord(N).Value
    Next N
    
    ' comments
    Record.txtComments.Text = Main.txtComments.Text
    
    UpdateRecordForm = True
    
    Exit Function
    
errhandler:
    UpdateRecordForm = False
    Exit Function
End Function

'------------------------------------------------------------------------------
' Parse and load the download disaster recovery copy of the current database.
' Created 12. August 2011, swr
'------------------------------------------------------------------------------
Public Function Parse_Database_Text(ByVal DataBaseText As String) As Boolean

On Error GoTo errhandler

Dim P                       As Long
Dim N                       As Long
Dim dummy                   As String
Dim Header                  As String
Dim Records                 As String
Dim Configuration           As String
Dim UploadSettings          As String
Dim Administrator           As String
Dim Selection               As String
Dim x()                     As String
Dim y()                     As String
Dim count                   As Long
                
    If DataBaseText = "ERROR" Or Len(DataBaseText) = 0 Then
        GoTo errhandler
    End If
                
    ' clear/reset database
    Call Database_EraseA
    
    ' if database is empty just parse the Header
    x() = Split(DataBaseText, "]" & vbCrLf)
    Header = Trim$(Left$(x(1), InStr(x(1), "[") - 1))
    Records = Trim$(Left$(x(2), InStr(x(2), "[") - 1))
    Configuration = Trim$(Left$(x(3), InStr(x(3), "[") - 1))
    UploadSettings = Trim$(Left$(x(4), InStr(x(4), "[") - 1))
    Administrator = Trim$(Left$(x(5), InStr(x(5), "[") - 1))
    Selection = Trim$(x(6))

    ' [HEADER]
    x() = Split(Header, vbCrLf)
    
    dummy = x(0)
    Main.lblLastSaved = x(1)
    captions(45) = x(2)
    captions(46) = x(3)
    Erase x()
    
    ' [RECORDS]
    x() = Split(Records, "-//-" & vbCrLf)
    count = 0
    NumRecords_G = UBound(x)
    ReDim nr(1 To NumRecords_G)
        
    For N = 0 To UBound(x) - 1
        count = count + 1
        Erase y()
        y() = Split(x(N), vbCrLf)
        nr(count).ID = y(0)                                     ' Y(0) = id
        For P = 1 To 26
            nr(count).txtField(P) = y(P)                        ' Y(2-27) = text fields
        Next P
        For P = 1 To 10
            nr(count).chkKeyWord(P) = y(P + 26)                 ' Y(28 - 37) = keywords
        Next P
        nr(count).Comments = Replace(y(37), "/¤¤/", vbCrLf)     ' Y (38)= comments
        DoEvents
    Next N
    Erase x()
    Erase y()
    CurrRecord_G = 1
    Call Record_ShowSingleA(CurrRecord_G)
        
    ' [CONFIGURATION]
    x() = Split(Configuration, vbCrLf)
    ReDim captions(1 To 80)
    count = 0
    For N = 0 To UBound(x) - 2
        count = count + 1
        captions(count) = x(N)
    Next N
    Erase x()
    
    Call FieldLabels_CopyFromArrayToMain
    
    ImgIndex_G = captions(47)
    TitleFontName_G = captions(48)
    TitleFontSize_G = captions(49)
    TitleColor_G = captions(50)
    TitleOpaque_G = captions(51)
    TitleBackcolor_G = captions(52)
    TitleBold_G = captions(53)
    TitleItalic_G = captions(54)
    Password_G = StringDecode(captions(55))
    ValidPasswords_G = StringDecode(captions(56))
    KeyWordLabelIndex_G = Val(captions(57))
    'NB values 58 - 80 are not currently in use
    
    ' [UPLOAD SETTINGS]
    x() = Split(UploadSettings, vbCrLf)
    With uploaddb
        .RemoteServerIP = x(0)
        .UserName = x(1)        ' note that username is encrypted at this point
        .PassWord = x(2)        ' note that password is encrypted at this point
        .RemoteFileName = x(3)
        .RemoteFolderPath = x(4)
        .LocalFileName = x(5)
        .WebsiteURL = x(6)
        .ProgramInfoURL = x(7)
        If x(8) = vbNullString Then
            .UpdateExeURL = "my download URL"
        Else
            .UpdateExeURL = x(8)
        End If
    End With
    Erase x()
    
    ' [ADMINISTRATOR]
    x() = Split(Administrator, vbCrLf)
    count = 0
    ReDim users(1 To 2, 1 To 6)
    ValidPasswords_G = "{@€[[Kra2t1"                        ' build-in master password
    For N = 0 To UBound(x) - 2
        If Len(x(N)) > 0 Then
            count = count + 1
            y() = Split(x(N), "::")
            users(1, count) = StringDecode(y(0))
            users(2, count) = StringDecode(y(1))
            ValidPasswords_G = ValidPasswords_G & "{@€[[" & users(2, count)
        End If
    Next N
    Erase x()
    Erase y()
    
    ' [SELECTION]
    x() = Split(Selection, vbCrLf)
    count = 0
    ReDim selected(1 To 2, 1 To 37)
    For N = 0 To UBound(x) - 1
        If Len(x(N)) > 0 Then
            count = count + 1
            y() = Split(x(N), ";")
            If y(0) = 0 Then
                selected(1, count) = 0
                selected(2, count) = vbNullString
            Else
                selected(1, count) = Val(y(0))  ' value
                selected(2, count) = Val(y(1))  ' index
            End If
        End If
    Next N
    Erase x()
    Erase y()
        
    Parse_Database_Text = True

Exit Function

errhandler:
    Parse_Database_Text = False
    Exit Function
End Function





Public Function String_Compress_Save(ByVal StringToCompress As String, ByVal FullFilePath As String) As Boolean

On Error GoTo errhandler

Dim fil                     As Long
Dim sCompressed             As String
Dim lCompressedLen          As Long
Dim lStringLen              As Long
    
    ' calculate buffer size
    lStringLen = Len(StringToCompress)
    lCompressedLen = (lStringLen * 1.01) + 13
    sCompressed = Space(lCompressedLen)
    
    ' perform compression
    If compress(sCompressed, lCompressedLen, StringToCompress, lStringLen) <> 0 Then
        GoTo errhandler
    End If
    
    sCompressed = Left(sCompressed, lCompressedLen)
        
    ' save compressed file
    fil = FreeFile
    Open FullFilePath For Binary As #fil
        Put #fil, , lStringLen & ":" & sCompressed
    Close fil
    
    String_Compress_Save = True
    
    Exit Function
    
errhandler:
    String_Compress_Save = False
    Exit Function
    
End Function

Public Function String_DeCompress(ByVal Compressed As String) As String

On Error GoTo errhandler

Dim sUncompressed           As String
Dim sCompressed             As String
Dim lenUncompressed         As Long
Dim lenCompressed           As Long
Dim ReturnValue             As Long
        
    ' get uncompressed length from msCompressed
    lenUncompressed = Val(Compressed)
    sUncompressed = Space(lenUncompressed)
    sCompressed = Mid(Compressed, InStr(Compressed, ":") + 1)
    lenCompressed = Len(sCompressed)
        
    ReturnValue = uncompress(sUncompressed, lenUncompressed, sCompressed, lenCompressed)
    
    Select Case ReturnValue
        Case 0
            String_DeCompress = sUncompressed
        Case Else
            GoTo errhandler
    End Select
            
    Exit Function
        
errhandler:
    String_DeCompress = "ERROR"
    Exit Function
    
End Function



'------------------------------------------------------------------------------
' Validate upload and download filetitel.
' Replace æ, ø, å and other illegal filename characters.
' Created 30. september 2007, swr
'------------------------------------------------------------------------------
Public Function GetRemoteFileName(ByVal UserFileName As String) As String

On Error GoTo errhandler

Dim tmp                     As String
    
    ' convert to lower case
    tmp = LCase$(Trim$(UserFileName))
    If Len(tmp) = 0 Then
        tmp = "sdu_database.zlb"
        Exit Function
    End If
     
    ' replace illegal characters
    Do While InStr(tmp, "æ") > 0: tmp = Replace(tmp, "æ", "ae"): Loop
    Do While InStr(tmp, "ø") > 0: tmp = Replace(tmp, "ø", "oe"): Loop
    Do While InStr(tmp, "å") > 0: tmp = Replace(tmp, "å", "aa"): Loop
    
    ' make sure that extention is "zlb" - add extentiln ".zlb"
    If InStr(tmp, ".") > 0 Then
        tmp = Left$(tmp, InStr(tmp, ".")) & "zlb"
    Else
        tmp = tmp & ".zlb"
    End If
    
    Do While InStr(tmp, " ") > 0: tmp = Replace(tmp, " ", "_"): Loop
    
    GetRemoteFileName = tmp
    
    Exit Function
    
errhandler:
    GetRemoteFileName = vbNullString
    Exit Function
End Function

'------------------------------------------------------------------------------
' See http://vbnet.mvps.org/
' Create a Ledger-Style Listview Report Background
'------------------------------------------------------------------------------
Public Sub SetListViewLedger(lv As ListView, ByVal ColorPattern As Long)

Dim iBarHeight              As Long
Dim lBarWidth               As Long
Dim diff                    As Long
Dim twipsy                  As Long
Dim lvWidth                 As Long
Dim C1                      As Long
Dim C2                      As Long
    
    Select Case ColorPattern
        Case 1: C1 = &HD7D8FE: C2 = &HEFF0FF    ' red
        Case 2: C1 = &HBCEAD4: C2 = &HE8F8F0    ' green
        Case 3: C1 = &HFDDFD2: C2 = &HFEF4F0    ' blue
        Case 4: C1 = &HA5EAEA: C2 = &HC5F2F2    ' yellow
        Case 5: C1 = &HE0EEF0: C2 = &HF5FAFA    ' grey
    End Select
    
    lvWidth = lv.Width
    lv.Visible = False
    lv.Width = 4 * Screen.Width
    
    iBarHeight = 0
    lBarWidth = 0
    diff = 0
   
    On Local Error GoTo SetListViewColor_Error
   
    twipsy = Screen.TwipsPerPixelY
   
    If lv.View = lvwReport Then
        With lv
            .Picture = Nothing
            .Refresh
            .Visible = 1
            .PictureAlignment = lvwTile
            lBarWidth = .Width
        End With
     
        With RecordList.Picture1
            .AutoRedraw = False
            .Picture = Nothing
            .BackColor = vbWhite
            .Height = 1
            .AutoRedraw = True
            .BorderStyle = vbBSNone
            .ScaleMode = vbTwips
            .Top = RecordList.Top - 10000
            .Width = Screen.Width
            .Visible = False
            .Font = lv.Font
        
        With .Font
            .Bold = lv.Font.Bold
            .Charset = lv.Font.Charset
            .Italic = lv.Font.Italic
            .Name = lv.Font.Name
            .Strikethrough = lv.Font.Strikethrough
            .Underline = lv.Font.Underline
            .Weight = lv.Font.Weight
            .Size = lv.Font.Size
        End With
              
        iBarHeight = .TextHeight("W")
        diff = Main.ImageList1.ImageHeight - (iBarHeight \ twipsy)
        iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
              
        .Height = iBarHeight * 2
        .Width = lBarWidth
         
        RecordList.Picture1.Line (0, 0)-(lBarWidth, iBarHeight), C1, BF
        RecordList.Picture1.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), C2, BF
        .AutoSize = True
        .Refresh
         
      End With
        lv.Width = lvWidth
        lv.Picture = RecordList.Picture1.Image
        lv.Visible = True
        lv.Refresh
    Else
        lv.Picture = Nothing
        
   End If

SetListViewColor_Exit:
On Local Error GoTo 0
Exit Sub
    
SetListViewColor_Error:
    With lv
        .Picture = Nothing
        .Refresh
    End With
   
    Resume SetListViewColor_Exit
    
End Sub

'------------------------------------------------------------------------------
' Create sub folder structure for MainDataPath and makes sure that
' all inifiles are present. If required inifiles are missing they are copied
' and renamed from the default inifiles located in app.path. After being copied
' the MAUN_DIR_G value is inserted in the first line of all inifiles.
' modified 23. september 2007
'------------------------------------------------------------------------------
Public Function CreateSystemFoldersA(ByVal MainDataPath As String) As Boolean

On Error GoTo errhandler
    
    ' exit if MainDataPath is not a valid database name
    If InStr(MainDataPath, "SDU_") = 0 Then GoTo errhandler
    
    ' remove trailing "\" if it exists before checking if folder already exists
    If Right$(MainDataPath, 1) = "\" Then
        MainDataPath = Left$(MainDataPath, Len(MainDataPath) - 1)
    End If
    
    ' check if database already exists on c-drive
    If Len(Dir(MainDataPath, vbDirectory)) > 0 Then
        CreateSystemFoldersA = False
    End If
        
    ' create MAIN_DIR_G if it does not exist
    If Len(Dir(MainDataPath, vbDirectory)) = 0 Then                                         ' MAIN
        MkDir MainDataPath
    End If
    MainDataPath = QualifyPath(MainDataPath)
    
' SUB FOLDERS ===========================================================================

    NOTES_DIR_G = MainDataPath & "NOTES"                                                    ' NOTES
    If Len(Dir(NOTES_DIR_G, vbDirectory)) = 0 Then
        MkDir NOTES_DIR_G
    End If
    NOTES_DIR_G = QualifyPath(NOTES_DIR_G)
                                                          ' NOTES\BACKUP
    If Len(Dir(MainDataPath & "NOTES" & "\BACKUP", vbDirectory)) = 0 Then
        MkDir MainDataPath & "NOTES" & "\BACKUP"
    End If
        
    EXCEL_DIR_G = MainDataPath & "EXCEL"                                                    ' EXCEL
    If Len(Dir(EXCEL_DIR_G, vbDirectory)) = 0 Then
        MkDir EXCEL_DIR_G
    End If
    EXCEL_DIR_G = QualifyPath(EXCEL_DIR_G)
    
    BACKUP_DIR_G = MainDataPath & "BACKUP"                                                  ' BACKUP
    If Len(Dir(BACKUP_DIR_G, vbDirectory)) = 0 Then
        MkDir BACKUP_DIR_G
    End If
    BACKUP_DIR_G = QualifyPath(BACKUP_DIR_G)
            
    CreateSystemFoldersA = True
    
    Exit Function
        
errhandler:
    CreateSystemFoldersA = False
    Exit Function
End Function







Public Function GetUniqueID() As String

On Error GoTo errhandler
    
    GetUniqueID = Format$(Date, "YYMMDD") & Format$(Now, "HHMMSS") & "-" & Format$(Rnd() * 1000000000000#, "000000000000")
    
errhandler:
    Exit Function
End Function



Public Function Record_GetNotes( _
                                  ByVal RecordNumber As Long, _
                                  Optional ByVal ShowAlways As Boolean = True, _
                                  Optional ByVal Silent As Boolean = False _
                                ) As Boolean

On Error GoTo errhandler

Dim tmp                     As String
Dim fil                     As Long

Const CRLF1 = "\par " & vbCrLf & "\par " & vbCrLf & "\par }"
Const CRLF2 = "\par " & vbCrLf & "\par }"
    
    If RecordNumber = 0 Then GoTo errhandler
    
    ' open note file
    If FileExist(NOTES_DIR_G & nr(RecordNumber).ID & ".note") Then
        fil = FreeFile
        Open NOTES_DIR_G & nr(RecordNumber).ID & ".note" For Input As #fil
            tmp = Input(LOF(fil), fil)
        Close #fil
        
        Do While InStr(tmp, CRLF1)
            tmp = Replace(tmp, CRLF1, CRLF2)
        Loop
        
        InfoBox.txtNotes.TextRTF = tmp
        If Len(InfoBox.txtNotes.Text) > 5 Then
           Record_GetNotes = True
           Main.cmdNotes.BackColor = &H7B8DF4           ' red
        Else
            tmp = vbNullString
            Record_GetNotes = False
            Main.cmdNotes.BackColor = 11468799          ' yellow
        End If
    Else
        tmp = vbNullString
        Record_GetNotes = False
        Main.cmdNotes.BackColor = 11468799              ' yellow
    End If
        
    If Silent Then
        Exit Function
    Else
        ' always show notes form
        If ShowAlways Then
            InfoBox.Caption = "Notes:   " & nr(RecordNumber).ID & ".note"
            InfoBox.txtNotes.TextRTF = tmp
            InfoBox.Show
        
        ' show notes form if text is longer than 110 chrs
        ElseIf Len(tmp) > 0 Then
            InfoBox.Caption = "Notes:   " & nr(RecordNumber).ID & ".note"
            InfoBox.txtNotes.TextRTF = tmp
            InfoBox.Show
        
        ' don't show notes form
        Else
            Exit Function
        End If
    End If
    
errhandler:
    Exit Function
End Function

'------------------------------------------------------------------------------
' Encodes user and license information before the information is
' stored in the Windows registry.
'------------------------------------------------------------------------------
Public Function StringEncode(ByVal InText As String) As String

On Error GoTo errhandler

Dim Plen                    As Long
Dim Pcurr                   As Long
Dim pMax                    As Long
Dim Ptxt                    As String
Dim Temp1                   As String
Dim Temp2                   As String
Dim char                    As String
Dim N                       As Long
Dim x                       As Long

    '--------------------------------------------------------------------------
    ' Exit with nothing if InText is empty
    '--------------------------------------------------------------------------
    If InText = vbNullString Then
        StringEncode = vbNullString
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' InText Xor Key60
    '--------------------------------------------------------------------------
    x = 0
    
    Pcurr = 1
    pMax = 5000
    Temp1 = Space$(pMax)
    
    For N = 1 To Len(InText)
        char = Mid$(InText, N, 1)
        x = x + 1
        If x > Len(swXORKEY) Then
            x = 1
        End If
        Ptxt = Chr$(Asc(char) Xor Asc(Mid$(swXORKEY, x, 1)))
        Plen = Len(Ptxt)
        If Pcurr + Plen > pMax Then Temp1 = Temp1 & Space$(100 * Plen): pMax = Len(Temp1)
        Mid$(Temp1, Pcurr) = Ptxt
        Pcurr = Pcurr + Plen
    Next N
    Temp1 = Left$(Temp1, Pcurr - 1)
    
    ' Convert Xor-String to Hex-String
    Pcurr = 1
    pMax = 5000
    Temp2 = Space$(pMax)
    For N = 1 To Len(Temp1)
        Ptxt = Hex(Asc(Mid$(Temp1, N, 1)))
        If Len(Ptxt) = 1 Then Ptxt = "0" & Ptxt
        Plen = 2
        If Pcurr + Plen > pMax Then Temp2 = Temp2 & Space$(100 * Plen): pMax = Len(Temp2)
        Mid$(Temp2, Pcurr) = Ptxt
        Pcurr = Pcurr + Plen
    Next N
    StringEncode = RTrim$(Temp2)
        
    Exit Function
    
errhandler:
    StringEncode = vbNullString
    Exit Function

End Function

'------------------------------------------------------------------------------
' Decodes user and license information read from the Registry.
' Convert Hex -> Chr -> Xor -> Return.
'------------------------------------------------------------------------------
Public Function StringDecode(ByVal InText As String) As String

On Error GoTo errhandler

Dim Plen                    As String
Dim Pcurr                   As Long
Dim pMax                    As Long
Dim Ptxt                    As String

Dim Temp1                   As String
Dim Temp2                   As String

Dim char                    As String
Dim N                       As Long
Dim x                       As Long

    '--------------------------------------------------------------------------
    ' Exit with nothing if InText is empty
    '--------------------------------------------------------------------------
    If InText = vbNullString Then
        StringDecode = vbNullString
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' Convert Hex to string
    '--------------------------------------------------------------------------
    Pcurr = 1
    pMax = 5000
    Temp1 = Space$(pMax)
    For N = 1 To Len(InText) Step 2
        char = Mid$(InText, N, 2)
        Ptxt = Chr$(swVal("&H" & char))
        Plen = Len(Ptxt)
        If Pcurr + Plen > pMax Then Temp1 = Temp1 & Space$(100 * Plen): pMax = Len(Temp1)
        Mid$(Temp1, Pcurr) = Ptxt
        Pcurr = Pcurr + Plen
    Next N
    Temp1 = Left$(Temp1, Pcurr - 1)
    
    ' Xor string and Crpt_Key60
    x = 0
    Pcurr = 1
    pMax = 5000
    Temp2 = Space$(pMax)
    For N = 1 To Len(Temp1)
        char = Mid$(Temp1, N, 1)
        x = x + 1
        If x > Len(swXORKEY) Then
            x = 1
        End If
        Ptxt = Chr$(Asc(char) Xor Asc(Mid$(swXORKEY, x, 1)))
        Plen = Len(Ptxt)
        If Pcurr + Plen > pMax Then Temp2 = Temp2 & Space$(100 * Plen): pMax = Len(Temp2)
        Mid$(Temp2, Pcurr) = Ptxt
        Pcurr = Pcurr + Plen
    Next N
    StringDecode = RTrim$(Temp2)
        
    Exit Function
    
errhandler:
    StringDecode = vbNullString
    Exit Function
    
End Function



'------------------------------------------------------------------------------
' Scan C:\ to build list of existing databases. Returns the last used db if it
' exists - if it does NOT exist then return first db that do exist.
' If No db's are found on the C-drive then return an empty string.
' Modified 19.8.2011, swr
'------------------------------------------------------------------------------
Public Function Installed_Databases(Optional ByVal LastUsedDatabase As String = vbNullString) As DATABASES_INFO

On Error Resume Next

Dim N                       As Long
Dim CountDB                 As Long
Dim fil                     As Long
Dim tmp                     As String
Dim lenFile                 As Long
Dim selDB                   As String
Dim count                   As Long
Dim FilePath                As String
Dim x()                     As String
Dim SDU()                   As String
Dim FP                      As FILE_PARAMS
       
    ' unload database menu submenu items
    For N = 1 To Main.DataBaseItem.UBound
       Unload Main.DataBaseItem(N)
    Next N
   
    ' set up search params
    With FP
        .sFileRoot = "C:\"        'start path
        .sFileNameExt = "*.*"     'file type of interest
    End With
    
    ' load FolderList with all folders on c-drive
    Call SearchForFolders(FP)
    
    ' process FolderList() to get all SDU_ database folder names
    count = 0
    For N = 1 To UBound(FolderList)
        If Left$(UCase$(FolderList(N)), 4) = "SDU_" Then
            count = count + 1
            ReDim Preserve SDU(1 To count)
            SDU(count) = FolderList(N)
        End If
    Next N
    
    Installed_Databases.TotalNumber = count
    
    ' if no SDU folders were found then return a nullstring and exit function
    If count = 0 Then
        GoTo errhandler
    End If
    
    ' clear app_LastOpen()
    ReDim app_LastOpen(1 To 3, 1 To 1)
    
    ' get information from database.txt in SDU_ folders
    CountDB = 0
    For N = 1 To UBound(SDU)
        SDU(N) = QualifyPath("C:\" & SDU(N))
        FilePath = SDU(N) & "database.txt"
        If FileExist(FilePath) Then
            fil = FreeFile
            Open FilePath For Input As #fil
                lenFile = LOF(fil)
                If lenFile > 500 Then lenFile = 500
                tmp = Input(lenFile, fil)
                x() = Split(tmp, vbCrLf)
                If UBound(x) > 4 Then
                    CountDB = CountDB + 1
                    ReDim Preserve app_LastOpen(1 To 3, 1 To CountDB)
                    app_LastOpen(1, CountDB) = x(3)                     ' Title of database
                    app_LastOpen(2, CountDB) = x(4)                     ' Subtitle of database
                    app_LastOpen(3, CountDB) = SDU(N)                   ' Folder name, Main_Dir_G
                    
                    If N = 1 Then Installed_Databases.FirstDB = SDU(N)  'x(5)
                    
                    If UCase$(LastUsedDatabase) = UCase$(SDU(N)) Then   ' Main_Dir_G
                        selDB = UCase$(x(5))                            ' Database exists and was used the last time
                        Installed_Databases.LastUsedDB = x(5)
                    End If
                End If
            Close #fil
        End If
    Next N
    
    ' update database menu on Main form
    CountDB = 0
    For N = 1 To UBound(app_LastOpen, 2)
        If Len(app_LastOpen(1, N)) > 0 Then
            CountDB = CountDB + 1
            Load Main.DataBaseItem(N)
            Main.DataBaseItem(N).Caption = Trim$(app_LastOpen(1, N)) & " - " & Trim$(app_LastOpen(2, N)) & Space$(5) & "[ " & app_LastOpen(3, N) & " ]"
        End If
    Next N
    
    Exit Function
    
errhandler:
    Installed_Databases.TotalNumber = 0
    Installed_Databases.FirstDB = vbNullString
    Installed_Databases.LastUsedDB = vbNullString
    Exit Function
End Function

Private Sub SearchForFolders(FP As FILE_PARAMS)

On Error GoTo errhandler

Dim WFD                     As WIN32_FIND_DATA
Dim hFile                   As Long
Dim sRoot                   As String
Dim spath                   As String
Dim sTmp                    As String
   
    sRoot = QualifyPath(FP.sFileRoot)
    spath = sRoot & FP.sFileNameExt
   
    'obtain handle to the first match
    hFile = FindFirstFile(spath, WFD)
   
    'if valid ...
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) And _
                Asc(WFD.cFileName) <> vbDot Then
            
                'must be a folder, so remove trailing nulls
                sTmp = TrimNull(WFD.cFileName)

                FP.count = FP.count + 1
                ReDim Preserve FolderList(1 To FP.count)
            
                FolderList(FP.count) = sTmp
            End If
      Loop While FindNextFile(hFile, WFD)
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If
   
errhandler:
   Exit Sub
End Sub

Public Function QualifyPath(spath As String) As String

On Error GoTo errhandler

    ' assures that a passed path ends in a slash
    If Right$(spath, 1) <> "\" Then
        QualifyPath = spath & "\"
    Else
        QualifyPath = spath
    End If
    
errhandler:
    Exit Function
End Function

Private Function TrimNull(startstr As String) As String

On Error GoTo errhandler

Dim pos As Integer
   
    pos = InStr(startstr, Chr$(0))
   
    If pos Then
        TrimNull = Left$(startstr, pos - 1)
        Exit Function
    End If
  
    TrimNull = startstr
    
errhandler:
    Exit Function
End Function

'------------------------------------------------------------------------------
' Read application captions from ini-file into application
'------------------------------------------------------------------------------
Public Sub FieldLabels_CopyFromArrayToMain()

On Error GoTo errhandler

Dim N                       As Long
    
    ' FieldCaptions for Fields 1 to 37
    For N = 1 To 37
        Main.lblField(N).Caption = Space$(1) & captions(N)
    Next N
           
    ' GroupCaptions (1 to 7)
    For N = 1 To 7
        Main.lblCaption(N).Caption = Space$(1) & captions(N + 37)
    Next N
        
    Main.lblApplicationTitle(0) = captions(45)
    Main.lblApplicationTitle(1) = captions(46)
    Main.Caption = captions(45) & " - " & captions(46)
    
errhandler:
    Exit Sub
End Sub

Public Sub LogFile_ReadA()

On Error GoTo errhandler

Dim fil                     As Long
Dim tmp                     As String
Dim N                       As Long
        
    If Not FileExist(MAIN_DIR_G & "program.log") Then
        ReDim SDU(0 To 0)
        GoTo errhandler
    End If
        
    fil = FreeFile
    Open MAIN_DIR_G & "program.log" For Input As #fil
        tmp = Input(LOF(fil), fil)
    Close #fil
        
    SDU() = Split(tmp, vbCrLf)
        
    ' truncate logfile after entry number 500
    If UBound(SDU) > 500 Then
        ReDim Preserve SDU(0 To 500)
    End If
    
    ' clear old logfile before writing
    AppConfig.lstLog.Clear
    
    ' write logfile entries from config.log
    For N = 0 To UBound(SDU)
        AppConfig.lstLog.AddItem SDU(N)
    Next N
    
errhandler:
   Exit Sub
End Sub

Public Sub LogFile_WriteA(ByVal Edit00_Upload01_Recover02_Backup03_Configuration04_Disaster05 As Long)

On Error GoTo errhandler

Dim fil                     As Long
Dim N                       As Long
Dim NewLine                 As String
Dim eName                   As String
Dim eAction                 As String
Dim t()                     As String
        
    Call LogFile_ReadA
            
    eName = GetUserName
    
    If eName = "Unknown" Then
        eAction = "Read access"
    Else
        eAction = "Edit access"
    End If
            
    Select Case Edit00_Upload01_Recover02_Backup03_Configuration04_Disaster05
        Case 0:    eAction = "Edit" & Space$(14 - Len("Edit"))
        Case 1:    eAction = "Upload" & Space$(14 - Len("Upload"))
        Case 2:    eAction = "Recover" & Space$(14 - Len("Recover"))
        Case 3:    eAction = "Backup" & Space$(14 - Len("Backup"))
        Case 4:    eAction = "Settings" & Space$(14 - Len("Settings"))
        Case 5:    eAction = "Disaster" & Space$(14 - Len("Disaster"))
        Case Else: eAction = "Other"
    End Select
    
    NewLine = Space$(1) & eAction & Chr(9) & Format$(Date, "DD. MMM YYYY") & " [" & Format$(Time, "HH:MM") & "]" & Space$(5) & eName
        
    fil = FreeFile
    Open MAIN_DIR_G & "program.log" For Output As #fil
        Print #fil, NewLine
        For N = 0 To UBound(SDU)
            Print #fil, SDU(N)
        Next N
    Close #fil
    
    ' update sdu()
    ReDim t(0 To UBound(SDU) + 1)   ' create tmp(), index sdu() + 1
    t(0) = NewLine                  ' add the new data line at index 0
    For N = 1 To UBound(t)          ' add entire sdu() from index 1 to ubound(t)
        t(N) = SDU(N - 1)           ' index sdu() = 0 to ubound(t) - 1
    Next N
    ReDim SDU(0 To UBound(t))       ' redim sdu() 0 to ubound(t)
    For N = 0 To UBound(t)
        SDU(N) = t(N)               ' fill sdu() with values from t()
    Next N
    Erase t()                       ' clear t()
        
errhandler:
    Exit Sub
End Sub

Public Function GetUserName() As String

On Error Resume Next

Dim N                       As Long
        
    For N = 1 To UBound(users, 2)
        If users(2, N) = Password_G Then
            GetUserName = users(1, N)
            Exit Function
        End If
    Next N
    
errhandler:
    GetUserName = "Unknown"
    Exit Function
    
End Function

'------------------------------------------------------------------------------
' replaces the Val function to avoid problems with decimal separators.
' correct the dot-comma problem and always return 0 if an error occurs
' exp greater than app. 300 returns 0
' created 04.02.2004, swr
'------------------------------------------------------------------------------
Public Function swVal(ByVal nString As String) As Double
    
On Error GoTo errhandler

Dim posexp                  As Long
Dim exp                     As Long
                
    nString = Trim$(LCase$(nString))
        
    ' exit is string is empty
    If nString = vbNullString Then
        swVal = 0
        Exit Function
    End If
        
    ' handle hex values
    If AscW(nString) = 38 And InStr("ABCDEFG1234567890", Mid$(nString, 2, 1)) > 0 Then      ' 38 = "&"
        swVal = CDbl(nString)
        Exit Function
    End If
        
    ' handle exp values
    posexp = InStr(nString, "e")
    If posexp > 0 Then
        exp = Val(Mid$(nString, posexp + 1))
        If exp > 300 Then
            swVal = 1E+300
            Exit Function               ' exit if value is greater than 1E+300
        ElseIf exp < -300 Then
            swVal = 1E-300              ' exit if value is smaller than 1E-300
            Exit Function
        End If
        ' add a "1" in front of exp values starting with "e"
        If AscW(nString) = 101 Then     ' 101 = "e"
            If InStrB(nString, "-") <> 0 Or InStrB(nString, "+") <> 0 Then
                nString = "1" & nString
            End If
        End If
    End If
    
    ' correct decimal separator problem
    If InStrB(CStr(1 / 2), ",") <> 0 Then
        nString = Replace$(nString, swDot, ",")
    End If
    
    ' CDbl wants the decimal separator specified by the operation system
    ' i.e., either a comma OR a dot depending on the country setting
    If IsNumeric(nString) Then
        swVal = CDbl(nString)
        Exit Function
    End If
    
    Exit Function
    
errhandler:
    swVal = 0
    Exit Function
    
End Function




'------------------------------------------------------------------------------
' Cause the selected form to remain on top - rescues forms positioned outside
' screen area
'

' created 15.7.2002, swr
'------------------------------------------------------------------------------
Public Sub form_StayOnTop( _
                           ByVal frmOnTop As Object, _
                           ByVal hWnd As String, _
                           ByVal Position As String _
                         )

On Error GoTo errhandler

Dim TpPT        As Long
Dim TpPL        As Long
Dim TpPH        As Long
Dim TpPW        As Long
Dim H           As Long
Dim oT          As Long
            
    Select Case UCase$(hWnd)
        Case "ABSOLUTE":        oT = -1   ' HWND_TOPMOST      above all non-topmost windows, maintains its topmost position when deactivated
        Case "TOPMOST":         oT = 2    ' HWND_NOTOPMOST    above all non-topmost windows
        Case "TOP":             oT = 0    ' HWND_TOP          top of z-order
        Case "BOTTOM":          oT = 1    ' HWND_BOTTOM       bottom of z-order
        Case Else:              oT = 0    ' HWND              top of z-order
    End Select
    
H = frmOnTop.Height
       
    '--------------------------------------------------------------------------
    ' Get conversion factors from twips to pixel for X and Y
    '--------------------------------------------------------------------------
    TpPH = Screen.TwipsPerPixelY    ' = pixels
    TpPW = Screen.TwipsPerPixelX

    Select Case UCase$(Position)
        Case Space$(1)                                               ' No change
            TpPT = frmOnTop.Top
            TpPL = frmOnTop.Left
            
            ' prevent form position from being outside screen area
            If frmOnTop.Top > Screen.Height - 100 Or _
               frmOnTop.Left > Screen.Width - 100 Then
                TpPT = 100
                TpPL = 100
            End If
            
        Case "UL", "LU"                                             ' Upper Left
            TpPT = (Screen.Height * 0.08) / TpPH
            TpPL = (Screen.Width * 0.02) / TpPW
        
        Case "ML", "LM"                                             ' Mid Left
            TpPT = (Screen.Height / 2) / TpPH
            TpPL = (Screen.Width * 0.02) / TpPW
        
        Case "LL"                                                   ' Lower Left
            TpPT = (Screen.Height - frmOnTop.Height * 1.5) / TpPH
            TpPL = (Screen.Width * 0.02) / TpPW
                
        Case "UR", "RU"                                             ' Upper Right
            TpPT = (Screen.Height * 0.08) / TpPH
            TpPL = (Screen.Width - frmOnTop.Width * 1.1) / TpPW
        
        Case "MR", "RM"                                             ' Mid Right
            TpPT = (Screen.Height / 2) / TpPH
            TpPL = (Screen.Width - frmOnTop.Width * 1.1) / TpPW
        
        Case "LR", "RL"                                             ' Lower Right
            TpPT = (Screen.Height - frmOnTop.Height * 1.5) / TpPH
            TpPL = (Screen.Width - frmOnTop.Width * 1.1) / TpPW
            
        Case "TL", "LT"                                             ' Top Left
            TpPT = 0 / TpPH
            TpPL = 0 / TpPW
        
        Case "TR", "RT"                                             ' Top Right
            TpPT = 0 / TpPH
            TpPL = (Screen.Width - frmOnTop.Width) / TpPW
            
        Case "BL", "LB"                                             ' Botom Left
            TpPT = (Screen.Height - frmOnTop.Height) / TpPH
            TpPL = 0 / TpPW
            
        Case "BR", "RB"                                             ' Bottom Right
            TpPT = (Screen.Height - frmOnTop.Height) / TpPH
            TpPL = (Screen.Width - frmOnTop.Width) / TpPW
            
        Case "C"                                                    ' Center
            TpPT = (Screen.Height / 2 - frmOnTop.Height / 2) / TpPH
            TpPL = (Screen.Width / 2 - frmOnTop.Width / 2) / TpPW
            
        Case Else                                                   ' No change
            TpPT = frmOnTop.Top
            TpPL = frmOnTop.Left
            
            ' prevent form position from being outside screen area
            If frmOnTop.Top > Screen.Height - 100 Or _
               frmOnTop.Left > Screen.Width - 100 Then
                TpPT = 100
                TpPL = 100
            End If
    End Select
    
    TpPH = frmOnTop.Height / TpPH   ' FormHeight in pixels
    TpPW = frmOnTop.Width / TpPW    ' FormWidth in pixels
            
    '----------------------------------------------------------------------
    ' Call API SetWindowPos sub
    '----------------------------------------------------------------------
    SetWindowPos frmOnTop.hWnd, oT, TpPL, TpPT, TpPW, TpPH, 0 '&H40
        
    frmOnTop.Height = H
    
    Exit Sub
    
errhandler:
    Exit Sub
            
End Sub


'------------------------------------------------------------------------------
' Receives the full path of a file (whether or not it exists)   N:\NNN\NNN\filename.nnn
' Returns the path WITH a backslash at the end.                 N:\NNN\NNN\
' This function is similar to the GetDirectoryPathSystem.
' Only the latter returns the path WITHOUT a "\" at the end     N:\NNN\NNN
'------------------------------------------------------------------------------
Public Function GetDirectoryPath( _
                                  ByVal NewPath As String _
                                ) As String

On Error GoTo errhandler
     
Dim pos                     As Long
    
    If Len(NewPath) > 0 Then
        pos = InStrRev(NewPath, "\")
        If pos > 0 And pos < Len(NewPath) Then
            NewPath = Left$(NewPath, pos)
        End If
    Else
        GoTo errhandler
    End If
           
    ' return path if it exists else default directory
    If Len(Dir(NewPath, vbDirectory)) > 0 Then
        GetDirectoryPath = UCase$(NewPath)
    End If
    
    Exit Function
    
errhandler:
    Exit Function
    
End Function


Public Sub Database_SaveBackupA(ByVal SDU_FolderName As String)

On Error GoTo errhandler

Dim DateStamp               As String
    
    If Len(SDU_FolderName) = 0 Then Exit Sub
    
    ' save database before making the backup
    Call Compressed_Database_Write(SDU_FolderName, False)
    
    ' create backup folder if it doesn't exist
    If Len(Dir(SDU_FolderName & "BACKUP\", vbDirectory)) = 0 Then
        MkDir SDU_FolderName & "BACKUP"
    End If
    
    ' compose data stamp
    DateStamp = Format$(Date, "YYMMDD") & Format$(Time, "HHMMSS")
    
    ' do backup
    FileCopy SDU_FolderName & "database.txt", SDU_FolderName & "BACKUP\database.txt_" & DateStamp
    FileCopy SDU_FolderName & "database.zlb", SDU_FolderName & "BACKUP\database.zlb_" & DateStamp
    
errhandler:
    Exit Sub
End Sub

Public Sub Notes_SaveBackupA()

On Error GoTo errhandler

Dim tDir                    As String
    
    If Len(Dir(NOTES_DIR_G & "BACKUP\", vbDirectory)) = 0 Then
        MkDir NOTES_DIR_G & "BACKUP"
    End If
        
    tDir = Dir(NOTES_DIR_G, vbNormal)
    Do While Len(tDir) > 0
        If Right$(tDir, 4) = "note" Then
            FileCopy NOTES_DIR_G & tDir, NOTES_DIR_G & "BACKUP\" & tDir
        End If
        tDir = Dir
        DoEvents
    Loop
    
errhandler:
    Exit Sub
End Sub

Public Sub DataBase_LoadBackupA(ByVal FullPathBackupFile As String)
    
On Error GoTo errhandler

Dim FileTitle               As String
Dim FolderName              As String
Dim BackUpDate              As String
Dim Success                 As Boolean
Dim x()                     As String
        
    ' save currently loaded database files (database.txt and database.zlb)
    Success = Compressed_Database_Write(MAIN_DIR_G)
            
    ' get file title and folder name for destination
    x() = Split(FullPathBackupFile, "\")
    FileTitle = x(UBound(x))
    FolderName = Left$(FullPathBackupFile, InStr(FullPathBackupFile, "BACKUP") - 1)
    BackUpDate = Mid$(FullPathBackupFile, InStrRev(FullPathBackupFile, "_") + 1)
    BackUpDate = Mid$(BackUpDate, 5, 2) & "-" & _
                 Mid$(BackUpDate, 3, 2) & "-" & "20" & _
                 Mid$(BackUpDate, 1, 2) & "  " & _
                 Mid$(BackUpDate, 7, 2) & ":" & _
                 Mid$(BackUpDate, 9, 2) & ":" & _
                 Mid$(BackUpDate, 11, 2)
                
    ' copy backup database file from backup folder to main folder
    If InStr(FileTitle, ".zlb") > 0 Then
        FileCopy FullPathBackupFile, MAIN_DIR_G & "database.zlb"
        
    ElseIf InStr(FileTitle, ".txt") > 0 Then
        FileCopy FullPathBackupFile, MAIN_DIR_G & "database.txt"
        
    Else
        GoTo errhandler
    End If
            
    ' load selected database
    MAIN_DIR_G = FolderName
    SaveSetting "SDU_UK", "User", "DataPath", MAIN_DIR_G
    
    Success = Compressed_Database_Read(MAIN_DIR_G)
    
    ' set database information on Main form
    If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
    Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
    Main.lblApplicationTitle(0).Caption = captions(45)
    Main.lblApplicationTitle(1).Caption = captions(46)
    Main.Caption = captions(45) & " - " & captions(46)
    
    ' lock new database
    Main.ToolsItem_Click 1
    
    Main.StatusBar1.Panels(2).Text = MAIN_DIR_G
    Main.StatusBar1.Panels(3).Text = " Database successfully retrieved from backup created: " & BackUpDate
    Main.lblWebDBVersion.Caption = "BACKUP"
    
errhandler:
    Exit Sub
End Sub

'------------------------------------------------------------------------------
' Clear all text fields and checkbox values but leaves the ID value unaltered
' Modified 29. October 2008, swr
'------------------------------------------------------------------------------
Public Sub Record_ClearCurrentA()

On Error GoTo errhandler

Dim N                       As Long
        
    ' clear text fields
    For N = 1 To 26
        Main.txtField(N).Text = vbNullString
    Next N
        
    ' clear keyword values
    For N = 1 To 10
        Main.chkKeyWord(N).Value = 0
    Next N
    
    ' clear comments
    Main.txtComments = vbNullString
        
    ' clear ID number display
    Main.lblUniqueID.Caption = "RECORD IS EMPTY !"
    
    ' paint fields
    Call PaintTextFields
    
errhandler:
    Exit Sub
End Sub


'----------------------------------------------------------------------------------------
' Compares local and remote copies of the database.
' If different then:
'                   1. always add local to remote - never opposite!
'                   2. do not delete if remote contains records not present in local
'                   3. when done then download remote replacing local
' Created 22. August 2011, swr
'----------------------------------------------------------------------------------------
Public Function UpdateRemoteRecords( _
                                     ByVal Local_Folder As String, _
                                     ByVal Remote_Titel As String) _
                                     As DATABASES_COMPARE
On Error GoTo errhandler

Dim N                       As Long
Dim P                       As Long
Dim count                   As Long
Dim fil                     As Long
Dim Success                 As Boolean
Dim ret                     As Long
Dim Local_Text              As String
Dim Remote_Text             As String
Dim tmpR()                  As String

Dim yConfiguration          As String
Dim yUploadSettings         As String
Dim yAdministrator          As String
Dim ySelection              As String
Dim yRemote                 As String

Dim x()                     As String
Dim xHeader                 As String
Dim xRecords                As String
Dim xR()                    As String

Dim y()                     As String
Dim yHeader                 As String
Dim yRecords                As String
Dim sRecords                As String
Dim yR()                    As String
    
    Local_Folder = QualifyPath(Local_Folder)
        
    If Not Get_Upload_Information(Local_Folder) Then
        GoTo errhandler
    End If
    
'GET LOCAL RECORDS
    If FileExist(Local_Folder & "database.zlb") Then
        fil = FreeFile
        Open Local_Folder & "database.zlb" For Binary As #fil
            Local_Text = Space(LOF(fil))
            Get #fil, , Local_Text
        Close fil
        If Len(Local_Text) > 1000 Then
            Local_Text = String_DeCompress(Local_Text)
        Else
            GoTo errhandler
        End If
    Else
        GoTo errhandler
    End If

'GET REMOTE RECORDS
        ret = WebsiteTransfer(vbGet, Local_Folder & "remote.zlb", Remote_Titel)
        If FileExist(Local_Folder & "remote.zlb") Then
        fil = FreeFile
        Open Local_Folder & "remote.zlb" For Binary As #fil
            Remote_Text = Space(LOF(fil))
            Get #fil, , Remote_Text
        Close fil
        If Len(Remote_Text) > 1000 Then
            Remote_Text = String_DeCompress(Remote_Text)
        Else
            GoTo errhandler
        End If
    Else
        GoTo errhandler
    End If
    
'SPLIT LOCAL TEXT INTO SECTIONS
    x() = Split(Local_Text, "]" & vbCrLf)
    xHeader = Trim$(Left$(x(1), InStr(x(1), "[") - 1))
    xRecords = Trim$(Left$(x(2), InStr(x(2), "[") - 1))
    
'SPLIT REMOTE TEXT INTO SECTIONS
    y() = Split(Remote_Text, "]" & vbCrLf)
    yHeader = Trim$(Left$(y(1), InStr(y(1), "[") - 1))
    yRecords = Trim$(Left$(y(2), InStr(y(2), "[") - 1))
    yConfiguration = Trim$(Left$(x(3), InStr(x(3), "[") - 1))
    yUploadSettings = Trim$(Left$(x(4), InStr(x(4), "[") - 1))
    yAdministrator = Trim$(Left$(x(5), InStr(x(5), "[") - 1))
    ySelection = Trim$(x(6))

'ANALYSE HEADERS
    If StrComp(xHeader, yHeader, vbTextCompare) = 0 Then
        UpdateRemoteRecords.HeadersIdentical = True
    Else
        UpdateRemoteRecords.HeadersIdentical = False
    End If
        
'ANALYSE RECORDS - GENERAL
    If StrComp(xRecords, yRecords, vbTextCompare) = 0 Then
        UpdateRemoteRecords.RecordsIdentical = True
        Exit Function
    Else
        UpdateRemoteRecords.RecordsIdentical = False
        xR() = Split(xRecords, "-//-" & vbCrLf)
        UpdateRemoteRecords.NumLocalRecords = UBound(xR)
        yR() = Split(yRecords, "-//-" & vbCrLf)
        UpdateRemoteRecords.NumRemoteRecords = UBound(yR)
    End If
    
'UPDATE RECORDS - LOCAL -> REMOTE
    count = 0                                                           ' collect non-empty, identical records in tmpR
    ReDim tmpR(0 To 1)
    For N = 0 To UBound(xR)
        For P = 0 To UBound(yR)
            If Len(x(N)) > 0 And Len(y(P)) > 0 Then
                If x(N) = y(P) Then
                    count = count + 1
                    ReDim Preserve tmpR(1 To count)
                    tmpR(count) = x(N)
                    x(N) = vbNullString
                    y(P) = vbNullString
                End If
            End If
        Next P
    Next N
    
    count = UBound(tmpR)                                                ' add non-empty, remote records to tmpR
    UpdateRemoteRecords.NumUpdatedRecords = 0
    For P = 0 To UBound(yR)
        If Len(yR(P)) > 0 Then
            count = count + 1
            ReDim Preserve tmpR(1 To count)
            tmpR(count) = yR(P)
            yR(P) = vbNullString
            UpdateRemoteRecords.NumUpdatedRecords = UpdateRemoteRecords.NumUpdatedRecords + 1
        End If
    Next P
        
'BUILD and SAVE REMOTE DATABASE STRING
    For N = 1 To UBound(tmpR)                                           ' build record section text string
        yRecords = yRecords & tmpR(N) & "-//-" & vbCrLf
    Next N
    
    yRemote = "[HEADER]" & vbCrLf & yHeader & vbCrLf & _
              "[RECORDS]" & vbCrLf & yRecords & vbCrLf & _
              "[CONFIGURATION]" & vbCrLf & yConfiguration & vbCrLf & _
              "[UPLOAD SETTINGS]" & vbCrLf & yUploadSettings & vbCrLf & _
              "[ADMINISTRATOR]" & vbCrLf & yAdministrator & vbCrLf & _
              "[SELECTION]" & vbCrLf & ySelection & vbCrLf
    Success = String_Compress_Save(yRemote, Local_Folder & "remote.zlb")
    
'UPLOAD UPDATED DATABASE
    If WebsiteTransfer(vbPut, Local_Folder & "remote.zlb", Remote_Titel) = 0 Then
        GoTo errhandler
    End If
       
    Exit Function
        
errhandler:
    UpdateRemoteRecords.NumUpdatedRecords = 0
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

'------------------------------------------------------------------------------
' Set the flag: WebServerConnectionOK_G. True if connection is OK - else False
' Flag is set in the Load Welcome form sub
'------------------------------------------------------------------------------
Public Sub SetInternetConnectionFlag()

On Error GoTo errhandler

Dim hOpen                   As Long
Dim hConnect                As Long
    
    Screen.MousePointer = 11
    
    ' test if a live IN connection is present   ????!!!!
    If InternetAttemptConnect(ByVal 0&) = 0 Then
    
        ' display user message form while attempting to connect ????!!!!
        hOpen = InternetOpen("sdu_uk", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        
        hConnect = InternetConnect(hOpen, _
                                          uploaddb.RemoteServerIP, _
                                          INTERNET_INVALID_PORT_NUMBER, _
                                          StringDecode(uploaddb.UserName), _
                                          StringDecode(uploaddb.PassWord), _
                                          INTERNET_SERVICE_FTP, _
                                          INTERNET_FLAG_PASSIVE, _
                                          0)
        
        'hConnect = InternetOpenUrl(hOpen, "http://www.google.dk/", vbNullString, ByVal 0&, &H80000000, ByVal 0&)
        
        ' close internet handles
        If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
        If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
        
        ' return connection status
        If hConnect = 0 Then
            WebServerConnectionOK_G = False
        Else
            WebServerConnectionOK_G = True
        End If
    Else
        GoTo errhandler
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
    
errhandler:
    If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
    If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
    Screen.MousePointer = 0
    WebServerConnectionOK_G = False
    Exit Sub
End Sub
'------------------------------------------------------------------------------
' Load and parse upload information from local copy - txt or compressed -
' of database.
' Return values: -1 if successful AND a live internet connection is available
'                0 True if successful - else False
' Created 26. july 2011 swr
'------------------------------------------------------------------------------
Public Function Get_Upload_Information( _
                                        ByVal SDU_FolderName As String, _
                                        Optional ByVal UserInfo As Boolean = False, _
                                        Optional ByVal GetLastUsedSettings As Boolean = False) _
                                        As Boolean
     
On Error GoTo errhandler

' if an internet connection is not available
If Not WebServerConnectionOK_G And Not GetLastUsedSettings Then
    Main.lblWebDBVersion.Caption = "LOCAL"
    Get_Upload_Information = False
    Screen.MousePointer = 0
    Exit Function
End If
    
' create
Dim fil                     As Long
Dim DataBaseText            As String
Dim Header                  As String
Dim Records                 As String
Dim Configuration           As String
Dim UploadSettings          As String
Dim Administrator           As String
Dim Selection               As String
Dim x()                     As String
Dim Response                As Long
Dim tmp                     As String
Dim N                       As Long
Dim msg                     As String
Dim msg1                    As String
Dim UploadInfoDirty         As Boolean
Dim Err_Number()            As String
Dim tDir                    As String
Dim Success                 As Long
        
ReDim Err_Number(1 To 10)

'///////////////////////////////////////////////////////////////////////////// DECOMPRESS DATABASE (.zlb -> .txt)
    ' decompress the zlb file if it exists - else move on and try
    ' loading the plain text file from the local folder
    
    DataBaseText = "ERROR"
    If FileExist(SDU_FolderName & "database.zlb") Then
        fil = FreeFile
        Open SDU_FolderName & "database.zlb" For Binary As #fil
            DataBaseText = Space(LOF(fil))
            Get #fil, , DataBaseText
        Close fil
        DataBaseText = String_DeCompress(DataBaseText)
        If Len(DataBaseText) > 1000 And DataBaseText <> "ERROR" Then
            GoTo ProcessDatabaseFile
        End If
    End If
    
'///////////////////////////////////////////////////////////////////////////// RESCUE LOAD FUNCTION
    
    ' if zlb file does not exist or if decompression fails then
    ' try loading the plain text file if it exists
    
    If DataBaseText = "ERROR" Then
        If FileExist(SDU_FolderName & "database.txt") Then
            fil = FreeFile
            Open SDU_FolderName & "database.txt" For Binary As #fil
                DataBaseText = Space$(LOF(fil))
                Get #fil, , DataBaseText
            Close fil
        End If
        If Len(DataBaseText) > 1000 And DataBaseText <> "ERROR" Then
            Main.lblWebDBVersion.Caption = "LOCAL"
            GoTo ProcessDatabaseFile
        Else
            GoTo errhandler
        End If
    End If
        
'///////////////////////////////////////////////////////////////////////////// PARSE String_DeCompressED DATABASE FILE
ProcessDatabaseFile:
    
    ' get database header and upload settings
    x() = Split(DataBaseText, "]" & vbCrLf)
    Header = Trim$(Left$(x(1), InStr(x(1), "[") - 1))
    UploadSettings = Trim$(Left$(x(4), InStr(x(4), "[") - 1))

'==================================================================================================
    ' skip this section if just retrieving last used settings
    If Not GetLastUsedSettings Then
    
        'HEADER
        ' correct if current folder is different from database folder name
        x() = Split(Header, vbCrLf)
    
        ' if the directory name on c-drive is different from the folder path included in the database file (= x(4)
        ' then create x(4) directory on the c-drive and replace folder path in database file
        If SDU_FolderName <> x(4) Then
        
            msg = "The current folder is: " & SDU_FolderName & Space$(5) & vbCrLf & _
                  "but the database was created in: " & x(4) & Space$(5) & vbCrLf & vbCrLf & _
                  "Correct the problem, Load another database or Quit the program:" & vbCrLf & vbCrLf & _
                  "YES" & Chr(9) & "- to correct the problem     " & vbCrLf & _
                  "NO" & Chr(9) & "- to load a different database     " & vbCrLf & _
                  "CANCEL" & Chr(9) & "- to exit Small Database Utility     "
            Response = MsgBox(msg, vbYesNoCancel + vbCritical, " WARNING")
            
            If Response = vbYes Then
                If Len(Dir$(x(4), vbDirectory)) = 0 Then
                    CreateSystemFoldersA (x(4))
                End If
            
                ' copy all files from sdu_foldername to x(4)
                tDir = Dir(SDU_FolderName & "*.*")
                Do
                    If Len(tDir) = 0 Then Exit Do
                    Success = CopyFile(SDU_FolderName & tDir, x(4) & tDir, False)                                                  ' do not overwrite
                    tDir = Dir
                Loop
                ' copy all files from sdu_foldername\NOTES to x(4)\NOTES
                tDir = Dir(SDU_FolderName & "NOTES\" & "*.*")
                Do
                    If Len(tDir) = 0 Then Exit Do
                    Success = CopyFile(SDU_FolderName & "NOTES\" & tDir, x(4) & "NOTES\" & tDir, False)                                                  ' do not overwrite
                    tDir = Dir
                Loop
                ' copy all files from sdu_foldername\EXCEL to x(4)\EXCEL
                tDir = Dir(SDU_FolderName & "EXCEL\" & "*.*")
                Do
                    If Len(tDir) = 0 Then Exit Do
                    Success = CopyFile(SDU_FolderName & "EXCEL\" & tDir, x(4) & "EXCEL\" & tDir, False)                                                  ' do not overwrite
                    tDir = Dir
                Loop
                ' copy all files from sdu_foldername\BACKUP to x(4)\BACKUP
                tDir = Dir(SDU_FolderName & "BACKUP\" & "*.*")
                Do
                    If Len(tDir) = 0 Then Exit Do
                    Success = CopyFile(SDU_FolderName & "BACKUP\" & tDir, x(4) & "BACKUP\" & tDir, False)                                                  ' do not overwrite
                    tDir = Dir
                Loop
                
                ' Correct folder problem on database file
                MAIN_DIR_G = x(4)
                Main.StatusBar1.Panels.Item(2).Text = MAIN_DIR_G
                fil = FreeFile
                Open SDU_FolderName & "database.txt" For Binary As #fil
                    DataBaseText = Space(LOF(fil))
                    Get #fil, , DataBaseText
                Close fil
                DataBaseText = Replace(DataBaseText, SDU_FolderName, MAIN_DIR_G)
                fil = FreeFile
                Open MAIN_DIR_G & "database.txt" For Binary As #fil
                    Put #fil, , DataBaseText
                Close fil
                Success = String_Compress_Save(DataBaseText, MAIN_DIR_G & "database.zlb")
                                
            ElseIf vbNo Then                                                ' LOAD ANOTHER DATABASE
                Main.Visible = True
                DoEvents
                GoTo errhandler
                
            Else                                                            ' EXIT SMALL DATABASE UTILITY
                Unload Main
                DoEvents
                
            End If
                    
        End If
    End If
    
'==================================================================================================

    Erase x()
    
    ' parse upload settings into uploaddb variable
    x() = Split(UploadSettings, vbCrLf)
    With uploaddb
        .RemoteServerIP = x(0)
        .UserName = x(1)        ' note that username is encrypted at this point
        .PassWord = x(2)        ' note that password is encrypted at this point
        .RemoteFileName = x(3)
        .RemoteFolderPath = x(4)
        .LocalFileName = x(5)
        .WebsiteURL = x(6)
        .ProgramInfoURL = x(7)
        If x(8) = vbNullString Then
            .UpdateExeURL = "my download URL"
        Else
            .UpdateExeURL = x(8)
        End If
    End With
    Erase x()

    If GetLastUsedSettings Then
        Get_Upload_Information = True
        Exit Function
    End If

    ' validate upload settings
    UploadInfoDirty = False
    With uploaddb
        tmp = Replace(.RemoteServerIP, ".", ""): If IsNumeric(tmp) = False Then Err_Number(1) = "Missing remote server IP address": UploadInfoDirty = True
        If InStr(.UserName, "my Username") > 0 Then Err_Number(2) = "Username is missing": UploadInfoDirty = True
        If InStr(.PassWord, "my Password") > 0 Then Err_Number(3) = "Password is missing": UploadInfoDirty = True
        If InStr(.RemoteFileName, "*.zlb") > 0 Then Err_Number(4) = "No remote file name": UploadInfoDirty = True
        If InStr(.RemoteFolderPath, "my Remote Folder Path") > 0 Then Err_Number(5) = "No remote folder path": UploadInfoDirty = True
        If InStr(.LocalFileName, "database.zlb") = 0 Then Err_Number(6) = ""
        If InStr(.WebsiteURL, "http://") = 0 Then tmp = Err_Number(7) = "URL to website incorrect": UploadInfoDirty = True
        If InStr(.ProgramInfoURL, "http://") = 0 Then tmp = Err_Number(8) = "Program info URL incorrect": UploadInfoDirty = True
        If InStr(.UpdateExeURL, "http://") = 0 Then tmp = Err_Number(9) = "Update URL incorrect": UploadInfoDirty = True
    End With
    
    ' inform user about status of the upload settings
    If UploadInfoDirty = True And UserInfo = True Then
    
        ' build error report...
        msg1 = ""
        For N = 1 To 10
            If Len(Err_Number(N)) > 0 Then
                msg1 = msg1 & Space$(5) & "- " & Err_Number(N) & vbCrLf
            End If
        Next N
        msg = "The up- and download information in the database file" & vbCrLf & _
              "you are about to load is incomplete:" & vbCrLf & vbCrLf & _
              msg1 & vbCrLf & _
              "In editing mode open the 'Advanced Functions' form," & vbCrLf & _
              "click 'Upload Settings' and enter the correct values" & vbCrLf & vbCrLf & _
              "When done upload the database to the website to verify" & vbCrLf & _
              "that the values are correct."
        
        Response = MsgBox(msg, vbExclamation + vbOKOnly, " UPLOAD DATA INFO")
    End If
    
    If UploadInfoDirty = True Then
        Get_Upload_Information = False
    Else
        Get_Upload_Information = True
    End If
    
    Screen.MousePointer = 0
        
    Exit Function
    
errhandler:
    Get_Upload_Information = False
    Main.lblWebDBVersion.Caption = "ERROR"
    Screen.MousePointer = 0
    Exit Function
End Function

'------------------------------------------------------------------------------
' Load and parse compressed database.
' Return values: 1 downloaded from website and successfully installed
'                2 download failed but database retrieved from local disk and successfully installed
'                3 retrieved from local disk and successfully installed
'                0 unknown error, trapped in errhandler
'               -1 local database file does not exists
'               -2 download failed and no local db file exists
' Modified 12. August 2011 swr
'------------------------------------------------------------------------------
Public Function Compressed_Database_Read( _
                                          ByVal SDU_FolderName As String, _
                                          Optional ByVal DownloadCompressedFile As Boolean = False _
                                        ) As Long
     
On Error GoTo errhandler

Dim fil                     As Long
Dim Success                 As Boolean
Dim DataBaseText            As String
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String

Screen.MousePointer = 11
Compressed_Database_Read = 999

    ' if no upload information available, load the local copy of data file instead
    If Not Get_Upload_Information(SDU_FolderName) Then
        Compressed_Database_Read = 2
        Main.lblWebDBVersion.Caption = "LOCAL"
        Main.lblWebDBVersion.ToolTipText = " Database is loaded from local disk "
        GoTo GetDatabaseLocally
    End If
        
'///////////////////////////////////////////////////////////////////////////// DOWNLOAD DATABASE
    If DownloadCompressedFile Then
        LocalFilePath = SDU_FolderName & uploaddb.LocalFileName
        RemoteFileName = uploaddb.RemoteFileName
        ret = WebsiteTransfer(vbGet, LocalFilePath, RemoteFileName)
                
        ' data file successfully downloaded and saved locally
        Compressed_Database_Read = 1                                                        ' file downloaded
        Main.lblWebDBVersion.Caption = "WEBSITE"
        Main.lblWebDBVersion.ToolTipText = " Database is loaded from the website "
    End If
        
'///////////////////////////////////////////////////////////////////////////// DECOMPRESS DATABASE (.zlb -> .txt)

GetDatabaseLocally:
    
    ' decompress the zlb file if it exists - else move on and try
    ' loading the plain text file from the local folder
    
    DataBaseText = "ERROR"
    If FileExist(SDU_FolderName & "database.zlb") Then
        fil = FreeFile
        Open SDU_FolderName & "database.zlb" For Binary As #fil
            DataBaseText = Space(LOF(fil))
            Get #fil, , DataBaseText
        Close fil
        DataBaseText = String_DeCompress(DataBaseText)
        
        ' if successfull then return 3 if database file was not downloaded -
        ' else 1 if download was successfull or 2 if download failed but a
        ' local, compressed database was successfully loaded
        
        If Len(DataBaseText) > 1000 Then
            If Compressed_Database_Read <> 1 And Compressed_Database_Read <> 2 Then
                Compressed_Database_Read = 3                                                ' retrieved from local file
            End If
        End If
    End If
    
'///////////////////////////////////////////////////////////////////////////// RESCUE LOAD FUNCTION
    
    ' if zlb file does not exist or if decompression fails then
    ' try loading the plain text file if it exists
    
    If DataBaseText = "ERROR" Then
        If FileExist(SDU_FolderName & "database.txt") Then
            fil = FreeFile
            Open SDU_FolderName & "database.txt" For Binary As #fil
                DataBaseText = Space$(LOF(fil))
                Get #fil, , DataBaseText
            Close fil
            
            If Compressed_Database_Read <> 1 And Compressed_Database_Read <> 2 Then
                Compressed_Database_Read = 3                                                ' retrieved from local file
            End If
            
        Else
            If Compressed_Database_Read = 2 Then
                Compressed_Database_Read = -2                                               ' no download, no local file
            Else
                Compressed_Database_Read = -1                                               ' no local file
            End If
            Main.lblWebDBVersion.Caption = "ERROR"
            Main.lblWebDBVersion.ToolTipText = " Unable to load database from local disk "
        End If
    End If
        
'///////////////////////////////////////////////////////////////////////////// PARSE DataBaseText string
    
    Success = Parse_Database_Text(DataBaseText)
    If Not Success Then
        GoTo errhandler
    End If
        
    ' if database was downloaded then check status of web edit lock on website
    If DownloadCompressedFile Then
        Call WebEditLock_READ(SDU_FolderName)
    End If
        
    Screen.MousePointer = 0
    
    Exit Function
    
errhandler:
    Compressed_Database_Read = 0
    Main.lblWebDBVersion.Caption = "ERROR"
    Main.lblWebDBVersion.ToolTipText = " Unidentified error - database not loaded "
    Screen.MousePointer = 0
    Exit Function
End Function


'------------------------------------------------------------------------------
' Read edit flag on website. Return string value. Modified November 29 2009
'------------------------------------------------------------------------------
Public Function WebEditLock_READ(ByVal SDU_FolderName As String) As String
     
On Error GoTo errhandler

Dim EditStatusFileName      As String
Dim fil                     As Long
Dim WebLockValue            As String
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String

    ' exit if upload information is not available
    If Not Get_Upload_Information(SDU_FolderName) Then
        Main.lblWebLockStatus.Caption = "?"
        Exit Function
    End If
        
    Screen.MousePointer = 11
    
    ' create edit status file name
    EditStatusFileName = Mid$(Replace(LCase$(SDU_FolderName), "\", vbNullString), 3) & "_editstatus.txt"
            
    LocalFilePath = SDU_FolderName & EditStatusFileName
    RemoteFileName = EditStatusFileName
    ret = WebsiteTransfer(vbGet, LocalFilePath, RemoteFileName)
            
    fil = FreeFile
    Open SDU_FolderName & EditStatusFileName For Binary As #fil
        WebLockValue = Space(LOF(fil))
        Get #fil, , WebLockValue
    Close #fil
    
    If WebLockValue = "ON" Then
        WebEditLock_READ = "ON"
        WebLockStatus_G = "ON"
        Main.lblWebLockStatus.ForeColor = &H40C0&
        Main.lblWebLockStatus.Caption = "ON"
        Main.lblWebLockStatus.ToolTipText = " Databasefile on the website is locked and cannot be downloaded..."
    Else
        WebEditLock_READ = "OFF"
        WebLockStatus_G = "OFF"
        Main.lblWebLockStatus.ForeColor = &H8000&
        Main.lblWebLockStatus.Caption = "OFF"
        Main.lblWebLockStatus.ToolTipText = " Databasefile on the website is free for download and editing..."
    End If
    
    Screen.MousePointer = 0
    Exit Function
    
errhandler:
    WebEditLock_READ = "ERROR"
    Main.lblWebLockStatus.ToolTipText = " Edit status of databasefile on the website cannot be determined..."
    Main.lblWebLockStatus.BackColor = &HC00000
    Screen.MousePointer = 0
    Exit Function
End Function


'------------------------------------------------------------------------------
' Compares the web and the local copy of the compressed database and warns user
' if the two files are different. Note that the comparison is performed after
' decompression of the databases.
' Return values:    True when the two databases are identical
'                   False when the two databases are different
' Modified 9. August 2011, swr
'------------------------------------------------------------------------------
Public Function Compare_Local_And_Remote_DB( _
                                             ByVal CompareOnProgramExit As Boolean, _
                                             Optional NoUploadOption As Boolean = False, _
                                             Optional ByVal Silent As Boolean = False) _
                                             As Boolean
     
On Error GoTo errhandler

Dim fil                     As Long
Dim msg0                    As String
Dim msg1                    As String
Dim msg2                    As String
Dim Response                As Long
Dim tmpDB_LOCAL             As String
Dim tmpDB_WEB               As String
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String
    
'///////////////////////////////////////////////////////////////////////////// DOWNLOAD and OPEN REMOTE COPY OF DATABASE as "tmpDB_WEB"
            
    If Not Get_Upload_Information(MAIN_DIR_G) Then
        GoTo errhandler
    Else
        LocalFilePath = MAIN_DIR_G & "tmpDB_WEB.zlb"
        RemoteFileName = uploaddb.RemoteFileName
        ret = WebsiteTransfer(vbGet, LocalFilePath, RemoteFileName)
                
        If FileExist(MAIN_DIR_G & "tmpDB_WEB.zlb") Then
            fil = FreeFile
            Open MAIN_DIR_G & "tmpDB_WEB.zlb" For Binary As #fil
                tmpDB_WEB = Space(LOF(fil))
                Get #fil, , tmpDB_WEB
            Close fil
            tmpDB_WEB = String_DeCompress(tmpDB_WEB)
            Kill MAIN_DIR_G & "tmpDB_WEB.zlb"
        Else
            GoTo errhandler
        End If
    End If
    
'///////////////////////////////////////////////////////////////////////////// OPEN LOCAL COPY as tmpDB_LOCAL
   
   ' load LOCAL copy of compressed database into tmpDB_LOCAL
    If FileExist(MAIN_DIR_G & "database.zlb") Then
        fil = FreeFile
        Open MAIN_DIR_G & "database.zlb" For Binary As #fil
            tmpDB_LOCAL = Space(LOF(fil))
            Get #fil, , tmpDB_LOCAL
        Close fil
        tmpDB_LOCAL = String_DeCompress(tmpDB_LOCAL)
    Else
        GoTo errhandler
    End If
    
'///////////////////////////////////////////////////////////////////////////// COMPARE LOCAL and REMOTE COPY

    ' if the web and local database strings are identical then exit with True
    If StrComp(tmpDB_LOCAL, tmpDB_WEB, vbBinaryCompare) = 0 Then
        Compare_Local_And_Remote_DB = True
        Exit Function
    Else
        If Silent Then
            Compare_Local_And_Remote_DB = False
            Exit Function
        Else
            ' create warning messages
            msg0 = "The LOCAL and the WEB copy of the database are DIFFERENT.     " & vbCrLf & vbCrLf & _
                   "The database may not have been uploaded to the website" & vbCrLf & _
                   "the last time it was edited." & vbCrLf & vbCrLf & _
                   "Upload the database to the website now ?"
            
            msg1 = "The LOCAL and the WEB copy of the database are DIFFERENT.     " & vbCrLf & vbCrLf & _
                   "The database may not have been uploaded to the website" & vbCrLf & _
                   "the last time it was edited."
                   
            msg2 = "Upload the database to the website before closing ? "
            
            If NoUploadOption Then
                Response = MsgBox(msg1, vbQuestion + vbOKOnly, " WARNING")                          ' on opening the program
                Compare_Local_And_Remote_DB = False
                Exit Function
            Else
                If CompareOnProgramExit And MasterUser_G Then
                    Response = MsgBox(msg2, vbDefaultButton1 + vbYesNo + vbQuestion, " WARNING")    ' on closing while Edit is enabled
                    DoEvents
                    If Response = vbYes Then
                        Main.FileItem_Click 11
                        Compare_Local_And_Remote_DB = True
                    Else
                        Compare_Local_And_Remote_DB = False
                    End If
                Else
                    Response = MsgBox(msg2, vbDefaultButton1 + vbYesNo + vbQuestion, " WARNING")     ' on program exit
                    DoEvents
                    If Response = vbYes Then
                        Main.FileItem_Click 11
                        Compare_Local_And_Remote_DB = True
                    Else
                        Compare_Local_And_Remote_DB = False
                    End If
                End If
            End If
        End If
    End If
    
    Exit Function
    
errhandler:
    Compare_Local_And_Remote_DB = False
    Exit Function
End Function



'------------------------------------------------------------------------------
' Build, write and upload (if selected) the database file.
'
' Return value = -1, function successfull
' Return value = 0, function failed
' Return value = 99, upload failed
'
' Created 25. october 2008 swr
'------------------------------------------------------------------------------
Public Function Compressed_Database_Write( _
                                           ByVal SDU_FolderName As String, _
                                           Optional ByVal UploadCompressedFile As Boolean = False _
                                         ) As Long
    
On Error GoTo errhandler

' create
Dim P                       As Long
Dim N                       As Long
Dim fil                     As Long
Dim DataBaseString          As String
Dim Success                 As Boolean
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String

Screen.MousePointer = 11

'///////////////////////////////////////////////////////////////////////////// BACKUP datafiles (.zlb and .txt)
    If Len(Dir(SDU_FolderName & "database.txt", vbNormal)) > 0 Then
        FileCopy SDU_FolderName & "database.txt", SDU_FolderName & "database.txt_bak"
        Kill SDU_FolderName & "database.txt"
    End If
    
    If Len(Dir(SDU_FolderName & "database.zlb", vbNormal)) > 0 Then
        FileCopy SDU_FolderName & "database.zlb", SDU_FolderName & "database.zlb_bak"
        Kill SDU_FolderName & "database.zlb"
    End If
    
'///////////////////////////////////////////////////////////////////////////// WRITE DATABASE FILE (.txt)
    fil = FreeFile
    Open SDU_FolderName & "database.txt" For Output As #fil
        
        Print #fil, "[HEADER]" '...............................................[HEADER]
        Print #fil, "Small Database Utility"
        Print #fil, "Database saved on " & Format$(Date, "DD. MMM YYYY") & " - " & Format$(Time, "HH:MM:SS")
        Print #fil, captions(45)
        Print #fil, captions(46)
        Print #fil, SDU_FolderName
        Print #fil,
        
        Print #fil, "[RECORDS - " & UBound(nr) & "]" '.........................[RECORDS]
        For N = 1 To UBound(nr)
            If Len(nr(N).ID) > 0 Then
                Print #fil, nr(N).ID                                ' record id             00
                For P = 1 To 26
                    Print #fil, nr(N).txtField(P)                   ' all text fields       01 - 26
                Next P
                For P = 1 To 10
                    Print #fil, nr(N).chkKeyWord(P)                 ' all keywords          27 - 37
                Next P
                Print #fil, Replace(nr(N).Comments, vbCrLf, "/¤¤/") ' comments              38
                Print #fil, "-//-"                                  ' end of record
            End If
        Next N
        Print #fil,
        
        Print #fil, "[CONFIGURATION]" '........................................[CONFIGURATION]
            For N = 1 To 80
                Print #fil, captions(N)
            Next N
        Print #fil,
        
        Print #fil, "[UPLOAD SETTINGS]" '......................................[UPLOAD SETTINGS]
        Print #fil, uploaddb.RemoteServerIP
        Print #fil, uploaddb.UserName
        Print #fil, uploaddb.PassWord
        Print #fil, uploaddb.RemoteFileName
        Print #fil, uploaddb.RemoteFolderPath
        Print #fil, uploaddb.LocalFileName
        Print #fil, uploaddb.WebsiteURL
        Print #fil, uploaddb.ProgramInfoURL
        Print #fil, uploaddb.UpdateExeURL
        Print #fil,
        
        Print #fil, "[ADMINISTRATOR]" '........................................[ADMINISTRATOR]
        For N = 1 To 6
            Print #fil, StringEncode(Trim$(users(1, N)));
            Print #fil, "::";
            Print #fil, StringEncode(Trim$(users(2, N)))
        Next N
        Print #fil,
        
        Print #fil, "[SELECTION]" '............................................[SELECTION]
        For N = 1 To 37
            If Len(selected(1, N)) > 0 Then
                Print #fil, selected(1, N) & ";" & selected(2, N)
            Else
                Print #fil, vbNullString
            End If
        Next N
    Close #fil
    
'///////////////////////////////////////////////////////////////////////////// COMPRESS DATABASE (.txt -> .zlb)
    Open SDU_FolderName & "database.txt" For Binary As #fil
        DataBaseString = Space(LOF(fil))
        Get #fil, , DataBaseString
    Close fil
    Success = String_Compress_Save(DataBaseString, SDU_FolderName & "database.zlb")
    
'///////////////////////////////////////////////////////////////////////////// UPLOAD DATABASE
    If UploadCompressedFile Then
        
        If Not Get_Upload_Information(SDU_FolderName) Then
            MsgBox "Upload information is missing, cannot upload.     ", vbInformation, " WARNING"
            Compressed_Database_Write = 99
            Exit Function
        End If
        
        LocalFilePath = SDU_FolderName & uploaddb.LocalFileName
        RemoteFileName = uploaddb.RemoteFileName
        ret = WebsiteTransfer(vbPut, LocalFilePath, RemoteFileName)
                
    End If
        
    Compressed_Database_Write = -1
    Screen.MousePointer = 0
    Exit Function
    
errhandler:
    Compressed_Database_Write = 0
    Screen.MousePointer = 0
    Exit Function
End Function


'------------------------------------------------------------------------------
' Set edit flag (..._editstatus.txt) on website. Created 29. November 2009 swr
'------------------------------------------------------------------------------
Public Function WebEditLock_ON(ByVal SDU_FolderName As String) As Boolean
    
On Error GoTo errhandler

Dim EditStatusFileName      As String
Dim fil                     As Long
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String
    
    ' exit if upload information is not available
    If Not Get_Upload_Information(SDU_FolderName) Then
        Main.lblWebLockStatus.Caption = "?"
        Exit Function
    End If
    
    ' Exit with True if status is already correct
    If WebLockStatus_G = "ON" Then
        WebEditLock_ON = True
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    ' create edit status file name
    EditStatusFileName = Mid$(Replace(LCase$(SDU_FolderName), "\", vbNullString), 3) & "_editstatus.txt"
        
    ' create edit status file
    fil = FreeFile
    Open SDU_FolderName & EditStatusFileName For Output As #fil
        Print #fil, "ON";
    Close #fil
        
    LocalFilePath = SDU_FolderName & EditStatusFileName
    RemoteFileName = EditStatusFileName
    ret = WebsiteTransfer(vbPut, LocalFilePath, RemoteFileName)
        
    WebEditLock_ON = True
    WebLockStatus_G = "ON"
    Main.lblWebLockStatus.ForeColor = &H40C0&
    Main.lblWebLockStatus.Caption = "ON"
    Main.lblWebLockStatus.ToolTipText = " Databasefile on the website is locked and cannot be downloaded..."
    Screen.MousePointer = 0
    Exit Function
    
errhandler:
    WebEditLock_ON = False
    Main.lblWebLockStatus.BackColor = &HC00000
    Main.lblWebLockStatus.ToolTipText = " Edit status of databasefile on the website cannot be determined..."
    Screen.MousePointer = 0
    Exit Function
End Function

'------------------------------------------------------------------------------
' Clear edit flag (..._editstatus.txt) on website. Created 29. November 2009 swr
'------------------------------------------------------------------------------
Public Function WebEditLock_OFF(ByVal SDU_FolderName As String) As Boolean
    
On Error GoTo errhandler

Dim fil                     As Long
Dim EditStatusFileName      As String
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String
                             
    ' exit if upload information is not available
    If Not Get_Upload_Information(SDU_FolderName) Then
        Main.lblWebLockStatus.Caption = "?"
        Exit Function
    End If
    
    ' Exit with True if status is already correct
    If WebLockStatus_G = "OFF" Then
        WebEditLock_OFF = True
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    ' create edit status file name
    EditStatusFileName = Mid$(Replace(LCase$(SDU_FolderName), "\", vbNullString), 3) & "_editstatus.txt"
    
    ' create local edit status file
    fil = FreeFile
    Open SDU_FolderName & EditStatusFileName For Output As #fil
        Print #fil, "OFF";
    Close #fil
        
    LocalFilePath = SDU_FolderName & EditStatusFileName
    RemoteFileName = EditStatusFileName
    ret = WebsiteTransfer(vbPut, LocalFilePath, RemoteFileName)
        
    WebLockStatus_G = "OFF"
    WebEditLock_OFF = True
    Main.lblWebLockStatus.ForeColor = &H8000&
    Main.lblWebLockStatus.Caption = "OFF"
    Main.lblWebLockStatus.ToolTipText = " Databasefile on the website is free for download and editing..."
    Screen.MousePointer = 0
    
    Exit Function
    
errhandler:
    WebEditLock_OFF = False
    Main.lblWebLockStatus.ForeColor = &HC00000
    Main.lblWebLockStatus.ToolTipText = " Edit status of databasefile on the website cannot be determined..."
    Screen.MousePointer = 0
    Exit Function
End Function






Public Function GetFileTitle(ByVal FullFilePath As String, Optional ByVal NoExtension As Boolean = False) As String

On Error GoTo errhandler

Dim P                       As Long
Dim tmp                     As String
    
    P = InStrRev(FullFilePath, "\")
    tmp = Mid$(FullFilePath, P + 1)
    
    If NoExtension Then
        P = InStrRev(tmp, ".")
        If P > 0 Then
            tmp = Left$(tmp, P - 1)
        End If
    End If
    
    GetFileTitle = tmp
    
errhandler:
    Exit Function
End Function

'------------------------------------------------------------------------------
' Issue a warning before deleting record by setting ID value to vbnullstring
' Modified 29. October 2008, swr
'------------------------------------------------------------------------------
Public Sub Record_DeleteSingleA(ByVal RecordNumber As Long)

On Error GoTo errhandler

If Not MasterUser_G Then
    MsgBox "Database is locked     ", vbInformation + vbOKOnly, " DATABASE IS LOCKED"
    Exit Sub
End If

Dim msg                     As String
Dim Response                As Long
Dim N                       As Long
Dim P                       As Long
Dim dummy()                 As RECORD_DATA

    msg = "You are about to delete Record nummer  " & CurrRecord_G & "  from the database     " & vbCrLf & vbCrLf & _
          Space$(6) & nr(RecordNumber).ID & vbCrLf & _
          Space$(6) & nr(RecordNumber).txtField(1) & vbCrLf & _
          Space$(6) & nr(RecordNumber).txtField(2) & vbCrLf & _
          Space$(6) & nr(RecordNumber).txtField(4) & vbCrLf & _
          Space$(6) & nr(RecordNumber).txtField(8) & vbCrLf & _
          Space$(6) & nr(RecordNumber).txtField(9) & vbCrLf & vbCrLf & _
          "Are you sure that you wish to permanently delete this Record?     "
          
    Response = MsgBox(msg, vbYesNo + vbQuestion, " WARNING", 0, 0)
    If Response = vbNo Then    ' answer = no
        Exit Sub
    End If
        
    ' clear all text fields and checkmarks
    Call Record_ClearCurrentA
    
    ' set record id of record to be deleted to nothing
    nr(RecordNumber).ID = vbNullString
    
    ' copy nr() array to dummy()
    ReDim dummy(1 To UBound(nr))
    For N = 1 To UBound(nr)
        dummy(N) = nr(N)
    Next N
    
    ' moves records with higher index than the record to be deleted one place down
    P = 0
    For N = 1 To UBound(dummy)
        If Len(dummy(N).ID) > 0 Then
            P = P + 1
            nr(P) = dummy(N)
        End If
    Next N
    
    ' redim nr to its actual size and change the total number of records
    ReDim Preserve nr(1 To P)
    NumRecords_G = UBound(nr)
    
    ' show the new record "RecordNumber"
    Record_ShowSingleA (RecordNumber)
                
errhandler:
    Exit Sub
End Sub



Public Sub SetActiveControls(ByVal DisableAll As Long)

On Error GoTo errhandler

Dim N                       As Long
    
    Select Case DisableAll
        
        Case 1                                                  ' UNLOCKED
            ' FileItem
            For N = 1 To 13
                Main.FileItem(N).Visible = True
            Next
            
            ' main commands buttons
            Main.cmdAction(1).Enabled = True    ' Delete
            Main.cmdAction(2).Enabled = True    ' Store
            Main.cmdAction(3).Enabled = True    ' New
            Main.cmdAction(5).Enabled = True    ' List
            Main.cmdAction(6).Enabled = True    ' Bookmark
            Main.cmdAction(7).Enabled = True    ' Previous/first
            Main.cmdAction(8).Enabled = True    ' Next/last
        
            ' search navigation command buttons
            Main.cmdNavSearch(0).Enabled = True
            Main.cmdNavSearch(1).Enabled = True
            
            ' search and jump text fields
            Main.txtSearch(0).Enabled = True
            Main.txtSearch(1).Enabled = True
            Main.txtJump.Enabled = True
            Main.lblJump.Enabled = True
            Main.lblFind(0).Enabled = True
            Main.lblFind(1).Enabled = True
            Main.lblMatches.Enabled = True
            
            ' section captions
            For N = 1 To 7
                Main.lblCaption(N).Enabled = True
            Next N
            Main.lblCaption(5).Enabled = False
            Main.lblCaption(6).Enabled = False
            
            ' text fields
            For N = 1 To 26
                Main.txtField(N).Enabled = True
            Next N
                        
            ' checkmarks keywords
            For N = 1 To 10
                Main.chkKeyWord(N).Enabled = True
            Next N
            
            ' preferences
            For N = 1 To 5
                Main.ToolsItem(N).Visible = True
            Next N
            
            Main.lblLocked.ForeColor = SDU_STATUS_COLOR_GREEN
            Main.lblLocked.Caption = SDU_STATUS_EDIT_ON
            Main.lblLocked.ToolTipText = " Click to set EDIT to OFF "
            
            MasterUser_G = True
            
            If WebServerConnectionOK_G Then         ' webserver connection OK
                Main.FileItem(10).Visible = True
                Main.FileItem(11).Visible = True
                Main.ToolsItem(1).Visible = True
                Main.HelpItem(1).Visible = True
                Main.HelpItem(2).Visible = True
                Main.HelpItem(3).Visible = True
                Main.HelpItem(4).Visible = True
            Else                                    ' webserver connection OFF
                Main.FileItem(10).Visible = False
                Main.FileItem(11).Visible = False
                Main.ToolsItem(1).Visible = False
                Main.HelpItem(1).Visible = False
                Main.HelpItem(2).Visible = False
                Main.HelpItem(3).Visible = False
                Main.HelpItem(4).Visible = False
            End If
            
        Case 2                                                  ' LOCKED
            ' FileItem
            For N = 1 To 13
                Main.FileItem(N).Visible = False
            Next
            Main.FileItem(0).Visible = True     ' open database
            Main.FileItem(2).Visible = True     ' sep
            Main.FileItem(3).Visible = True     ' create new database
            Main.FileItem(12).Visible = True    ' sep
            Main.FileItem(13).Visible = True    ' hide
            Main.FileItem(14).Visible = True    ' sep
            Main.FileItem(15).Visible = True    ' exit
            
            ' main commands buttons
            Main.cmdAction(1).Enabled = False   ' - Delete
            Main.cmdAction(2).Enabled = False   ' - Store
            Main.cmdAction(3).Enabled = False   ' - New
            Main.cmdAction(5).Enabled = True    ' List
            Main.cmdAction(6).Enabled = True    ' Bookmark
            Main.cmdAction(7).Enabled = True    ' Previous
            Main.cmdAction(8).Enabled = True    ' Next
            
            ' search navigation command buttons
            Main.cmdNavSearch(0).Enabled = True
            Main.cmdNavSearch(1).Enabled = True
            
            ' search and jump text fields
            Main.txtSearch(0).Enabled = True
            Main.txtSearch(1).Enabled = True
            Main.txtJump.Enabled = True
            Main.lblJump.Enabled = True
            Main.lblFind(0).Enabled = True
            Main.lblFind(1).Enabled = True
            Main.lblMatches.Enabled = True
            
            ' section captions
            For N = 1 To 7
                Main.lblCaption(N).Enabled = True
            Next N
            Main.lblCaption(5).Enabled = False      ' brænde
            Main.lblCaption(6).Enabled = False      ' pileflet
        
            ' text fields
            For N = 1 To 26
                Main.txtField(N).Enabled = True
            Next N
                        
            ' keywords
            For N = 1 To 10
                Main.chkKeyWord(N).Enabled = False
            Next N
                                    
            ' list and validate items
            Main.ListItem(0).Enabled = True
            
            ' options menu
            For N = 1 To 5
                Main.ToolsItem(N).Visible = False
            Next N
            
            Main.lblLocked.ForeColor = SDU_STATUS_COLOR_RED
            Main.lblLocked.Caption = SDU_STATUS_EDIT_OFF
            Main.lblLocked.ToolTipText = " Click to set EDIT to ON "
            
            MasterUser_G = False
            
            If WebServerConnectionOK_G Then         ' webserver connection OK
                Main.ToolsItem(1).Visible = True
                Main.HelpItem(1).Visible = True
                Main.HelpItem(2).Visible = True
                Main.HelpItem(3).Visible = True
                Main.HelpItem(4).Visible = True
            Else                                    ' webserver connection OFF
                Main.ToolsItem(1).Visible = False
                Main.HelpItem(1).Visible = False
                Main.HelpItem(2).Visible = False
                Main.HelpItem(3).Visible = False
                Main.HelpItem(4).Visible = False
            End If
    
    End Select
            
    Main.Refresh
    
errhandler:
    Exit Sub
End Sub


Public Sub Record_Find(ByVal qString As String, ByVal InComments As Boolean)

On Error GoTo errhandler

Dim N                       As Long
Dim P                       As Long
Dim sCount                  As Long
Dim tmpString               As String
    
    ' global array to hold search result
    ReDim sr(1 To 2, 1 To UBound(nr))
    sCount = 0
    
    Select Case InComments
        Case False                      ' in RECORDS
            For N = 1 To UBound(nr)
                For P = 1 To 26
                    tmpString = nr(N).txtField(P)
                    If P > 9 And P < 18 Then
                        tmpString = Replace(tmpString, " ", vbNullString)
                    End If
                    If Len(Trim$(Main.lblField(P))) > 0 Then
                        If InStr(tmpString, qString) > 0 Then               ' get FIRST occurrence of qString
                            sCount = sCount + 1
                            sr(1, sCount) = N
                            sr(2, sCount) = P
                            Exit For
                        End If
                    End If
                Next P
            Next N
            
        Case True                       ' in COMMENTS
            For N = 1 To UBound(nr)
                If InStr(nr(N).Comments, qString) > 0 Then                  ' get FIRST occurrence of qString
                    sCount = sCount + 1
                    sr(1, sCount) = N
                    sr(2, sCount) = 1
                    Exit For
                End If
            Next N
            
    End Select
            
    ' redim sr() to fit number of matches
    If sCount = 0 Then
        ReDim sr(1 To 2, 1 To 1)
    Else
        ReDim Preserve sr(1 To 2, 1 To sCount)
    End If
    
errhandler:
    Exit Sub
End Sub




Public Function Notes_LoadBackupA() As Boolean

On Error GoTo errhandler

Dim tDir                    As String
                
    ' delete files in notes folder, NOTES_DIR_G
    tDir = Dir(NOTES_DIR_G, vbNormal)
    Do While Len(tDir) > 0
        If FileExist(NOTES_DIR_G & tDir) And Right$(tDir, 4) = "note" Then
            Kill NOTES_DIR_G & tDir
        End If
        tDir = Dir
        DoEvents
    Loop
    
    ' copy all note files from NOTES_BACKUP to NOTES_DIR_G
    tDir = Dir(NOTES_DIR_G & "BACKUP\", vbNormal)
    Do While Len(tDir) > 0
        If FileExist(NOTES_DIR_G & "BACKUP\" & tDir) Then
            FileCopy NOTES_DIR_G & "BACKUP\" & tDir, NOTES_DIR_G & tDir
        End If
        tDir = Dir
        DoEvents
    Loop
    
    Notes_LoadBackupA = True
    
    Exit Function
    
errhandler:
    Notes_LoadBackupA = False
    Exit Function
End Function


Public Sub PaintTextFields()

On Error GoTo errhandler

Dim N                       As Long

    For N = 1 To 26
        If Len(Main.txtField(N).Text) > 0 Then
            Main.txtField(N).BackColor = &HFFF9DB
            Main.txtField(N).Refresh
        Else
            Main.txtField(N).BackColor = &HF6F6F6
            Main.txtField(N).Refresh
        End If
        DoEvents
    Next N
        
errhandler:
    Exit Sub
End Sub

Public Function LastRecordIsEmpty(ByVal RecordNumber As Long) As Boolean

On Error GoTo errhandler

Dim N                       As Long
    
    LastRecordIsEmpty = True
    
    ' check if all text fields are empty
    For N = 1 To 26
        If Len(nr(RecordNumber).txtField(N)) > 0 Then
            LastRecordIsEmpty = False
            Exit For
        End If
    Next N
    
errhandler:
    Exit Function
End Function


'------------------------------------------------------------------------------
' Displays selected record
'------------------------------------------------------------------------------
Public Function Record_ShowSingleR(ByVal RecordNumber As Long)
    
On Error GoTo errhandler
    
Dim N                       As Long
    
    ' show text fields
    For N = 1 To 26
        Record.txtField(N).Text = Trim$(nr(RecordNumber).txtField(N))
    Next N
    
    ' show keywords
    For N = 1 To 10
         Record.chkKeyWord(N).Value = Val(nr(RecordNumber).chkKeyWord(N))
    Next N
    
    ' show comments
    Record.txtComments = StringDecode(Trim$(nr(RecordNumber).Comments))
           
    Record_ShowSingleR = True
    
    Exit Function
    
errhandler:
    Record_ShowSingleR = False
    Exit Function
End Function
    
'------------------------------------------------------------------------------
' Displays selected record
'------------------------------------------------------------------------------
Public Function Record_ShowSingleA(ByVal RecordNumber As Long)
    
On Error GoTo errhandler
    
Dim N                       As Long
    
    ' show text fields
    For N = 1 To 26
        Main.txtField(N).Text = Trim$(nr(RecordNumber).txtField(N))
    Next N
    
    ' show keywords
    For N = 1 To 10
         Main.chkKeyWord(N).Value = Val(nr(RecordNumber).chkKeyWord(N))
    Next N
    
    ' show comments
    Main.txtComments = StringDecode(Trim$(nr(RecordNumber).Comments))
    
    ' show record ID number
    If Len(nr(RecordNumber).ID) = 0 Then
        Main.lblUniqueID.Caption = "RECORD IS EMPTY !"
    Else
        Main.lblUniqueID.Caption = nr(RecordNumber).ID
    End If
    
    Main.StatusBar1.Panels.Item(1).Text = " Record = " & RecordNumber
    
    ' refresh form
    Main.Refresh
    DoEvents
    
    ' paint active foelds cyan, rest grey
    Call PaintTextFields
       
    Record_ShowSingleA = True
    
    Exit Function
    
errhandler:
    Record_ShowSingleA = False
    Exit Function
    
End Function

'------------------------------------------------------------------------------
' See if the record is empty
'------------------------------------------------------------------------------
Public Function Record_IsEmptyA(ByVal RecordNumber As Long) As Boolean

On Error GoTo errhandler

Dim N                       As Long
Dim lenRec                  As Long
    
    ' see if all txtfields of the record are empty
    lenRec = 0
    For N = 1 To 26
        lenRec = lenRec + Len(nr(RecordNumber).txtField(N))
        If lenRec > 0 Then Exit For
    Next N
    
    If lenRec = 0 Then
        Record_IsEmptyA = True
    Else
        Record_IsEmptyA = False
    End If
        
errhandler:
    Exit Function
End Function




'------------------------------------------------------------------------------
' from Randi Birch VBnet's ressource page at :
' http://vbnet.mvps.org/code/browse/browsefolders.htm
'------------------------------------------------------------------------------
Public Function GetDirectoryDialog(myForm As Object) As String

On Error GoTo errhandler

Dim bi                      As BrowseInfo
Dim pidl                    As Long
Dim Path                    As String
Dim pos                     As Long
       
    With bi
        .hOwner = myForm.hWnd
        .lpszTitle = "Select Main Data Folder"
        .ulFlags = &H1
        .pIDLRoot = 0&
    End With
    
    pidl = SHBrowseForFolder(bi)
    Path = Space$(260)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
          pos = InStr(Path, Chr$(0))
    End If

    Call CoTaskMemFree(pidl)
    
    If Len(Trim$(Path)) < 2 Then GoTo errhandler
    
    GetDirectoryDialog = Left$(Path, pos - 1) & "\"
    
    Exit Function
    
errhandler:
    GetDirectoryDialog = vbNullString
    
End Function

'------------------------------------------------------------------------------
' Transfers data from text fields to nr() array.
' Parameters:
'            RecordNumber
'------------------------------------------------------------------------------
Public Function Record_StoreSingleA(ByVal RecordNumber As Long) As Boolean

On Error GoTo errhandler

Dim N                       As Long
Dim KeyWordSet              As Boolean
    
    ' set KeyWordSet to false
    KeyWordSet = False
    
    ' store text fields
    For N = 1 To 26
        nr(RecordNumber).txtField(N) = Trim$(Main.txtField(N).Text)
    Next N
        
    ' store keyword values
    For N = 1 To 10
        nr(RecordNumber).chkKeyWord(N) = Main.chkKeyWord(N).Value
        
        ' if at least one keyword is entered then set KeyWordSet to true
        If Main.chkKeyWord(N).Value = 1 And KeyWordSet = False Then
            KeyWordSet = True
        End If
    Next N
    
    ' store comments
    nr(RecordNumber).Comments = StringEncode(Trim$(Main.txtComments.Text))
            
    ' display record ID number
    Main.lblUniqueID.Caption = nr(RecordNumber).ID
        
    If Not KeyWordSet Then
        MsgBox "This record does not have a keyword!     ", vbInformation + vbOKOnly, " WARNING"
    End If
    
    Record_StoreSingleA = True
    
    Exit Function
    
errhandler:
    Record_StoreSingleA = False
    Exit Function
    
End Function


'--------------------------------------------------------------------------------------------------------
' Internet connection test. Seems to fail for unknown reasons
' hConnect = InternetOpenUrl(hOpen, "http://www.google.dk/", vbNullString, ByVal 0&, &H80000000, ByVal 0&)
' Using attempt to connect to webserver instead
' Created 21. August 2011, swr
'--------------------------------------------------------------------------------------------------------
Public Sub WebServerConnectionStatus_Refresh( _
                                              ByVal DatabaseFolderPath As String, _
                                              Optional ByVal ShowStatus_Seconds As Single = 10)
On Error GoTo errhandler

Dim hOpen                   As Long
Dim hConnect                As Long
   
Screen.MousePointer = 11: DoEvents

'USER MESSAGE 1

    ' display user message form while attempting to connect to webserver
    Call Get_Upload_Information(DatabaseFolderPath, False, True)    ' get connect parameters for current session
    Load InternetConnection
    InternetConnection.Height = 1665
    InternetConnection.lblStatus.Height = 900
    InternetConnection.lblTitle.Caption = "Server Connection...?"
    InternetConnection.lblStatus.Caption = "Checking the status of your connection to the webserver." & vbCrLf & vbCrLf & "Please wait, checking may last up to 30 sec."
    InternetConnection.Refresh
    InternetConnection.Visible = True
    WaitABit 0.5
    
'CONNECT TO WEBSERVER
    
    ' attempt to connect to the webserver specified in upload settings
    If InternetAttemptConnect(ByVal 0&) = 0 Then
        hOpen = InternetOpen("sdu_uk", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        
        hConnect = InternetConnect(hOpen, _
                                          uploaddb.RemoteServerIP, _
                                          INTERNET_INVALID_PORT_NUMBER, _
                                          StringDecode(uploaddb.UserName), _
                                          StringDecode(uploaddb.PassWord), _
                                          INTERNET_SERVICE_FTP, _
                                          INTERNET_FLAG_PASSIVE, 0)
        
        ' set global status flag for webserver connection
        If hConnect = 0 Then
            WebServerConnectionOK_G = False
        Else
            WebServerConnectionOK_G = True
        End If
    Else
        WebServerConnectionOK_G = False
    End If
    If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
    If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
    
'USER MESSAGE 2

    WaitABit 0.5
    Screen.MousePointer = 0
    If WebServerConnectionOK_G Then
        InternetConnection.lblTitle.Caption = "Connection is OK"
        InternetConnection.Refresh
        WaitABit 2
        Unload InternetConnection
    Else
        InternetConnection.Height = 2070
        InternetConnection.lblStatus.Height = 1260
        InternetConnection.lblTitle.Caption = "No Connection...!"
        InternetConnection.lblStatus.Caption = "There is no connection to the webserver at the moment. " & _
                                               "You will not be able to up- and download the database to/from the website." & vbCrLf & vbCrLf & _
                                               "This implies that editing the database ONLY affects the local copy on your own computer."
        InternetConnection.Refresh
        InternetConnection.Visible = True
        WaitABit ShowStatus_Seconds
        Unload InternetConnection
    End If
    
    Exit Sub
    
errhandler:
    If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
    If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
    Screen.MousePointer = 0
    WebServerConnectionOK_G = False
    Exit Sub
End Sub

Public Function WebsiteTransfer( _
                                 ByVal Put_Get As String, _
                                 ByVal LocalFilePath As String, _
                                 ByVal RemoteFileName As String) _
                                 As Boolean
On Error GoTo errhandler
    
If Not WebServerConnectionOK_G Then GoTo errhandler
    
Dim ConnectAttempts         As Long
Dim TransferAttempts        As Long
Dim hOpen                   As Long
Dim hConnect                As Long
Dim ret                     As Long
    
    Put_Get = UCase$(Put_Get)
    
    hOpen = InternetOpen("sdu_uk", _
                                   INTERNET_OPEN_TYPE_DIRECT, _
                                   vbNullString, _
                                   vbNullString, _
                                   0)
    ConnectAttempts = 0
RetryConnect:
    hConnect = InternetConnect(hOpen, _
                                      uploaddb.RemoteServerIP, _
                                      INTERNET_INVALID_PORT_NUMBER, _
                                      StringDecode(uploaddb.UserName), _
                                      StringDecode(uploaddb.PassWord), _
                                      INTERNET_SERVICE_FTP, _
                                      INTERNET_FLAG_PASSIVE, _
                                      0)
                                      
    ' repeat connect attenmpt 3 times on failure
    If hConnect = 0 And ConnectAttempts < 3 Then
        ConnectAttempts = ConnectAttempts + 1
        WaitABit 1
        GoTo RetryConnect
    ElseIf ConnectAttempts > 2 Then
        GoTo errhandler
    End If
                                  
    ret = FtpSetCurrentDirectory(hConnect, _
                                  uploaddb.RemoteFolderPath)
                                  
    TransferAttempts = 0
RetryTransfer:
    If Put_Get = "PUT" Then
        ret = FtpPutFile(hConnect, _
                                   LocalFilePath, _
                                   RemoteFileName, _
                                   FTP_TRANSFER_TYPE_BINARY, _
                                   0)
    ElseIf Put_Get = "GET" Then
        ret = FtpGetFile(hConnect, _
                                   RemoteFileName, _
                                   LocalFilePath, _
                                   False, _
                                   vbNormal, _
                                   FTP_TRANSFER_TYPE_BINARY, _
                                   0)
    Else
        GoTo errhandler
    End If
                     
    ' repeat transfer attenmpt 3 times on failure
    If ret = 0 And TransferAttempts < 3 Then
        TransferAttempts = TransferAttempts + 1
        WaitABit 1
        GoTo RetryTransfer
    ElseIf TransferAttempts > 2 Then
        GoTo errhandler
    End If
                     
    If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
    If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
    
    If ret = 0 Then
        GoTo errhandler
    Else
        WebsiteTransfer = -1
        Exit Function
    End If
    
errhandler:
    If hOpen <> 0 Then InternetCloseHandle hOpen: DoEvents
    If hConnect <> 0 Then InternetCloseHandle hConnect: DoEvents
    WebsiteTransfer = 0
    Exit Function
End Function


