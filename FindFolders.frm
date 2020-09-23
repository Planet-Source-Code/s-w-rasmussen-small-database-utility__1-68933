VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   8220
      TabIndex        =   6
      Top             =   1500
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   7455
   End
   Begin VB.CheckBox Check1 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6660
      TabIndex        =   2
      Top             =   300
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   3660
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2008 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function GetTickCount Lib "kernel32" () As Long



Private Sub Command1_Click()

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   
  'clear results textbox and list
   Text3.Text = ""
   
  'set up search params
   With FP
      .sFileRoot = Text1.Text       'start path
      .sFileNameExt = Text2.Text    'file type of interest
      .bRecurse = Check1.Value = 1  '1 = do recursive search
   End With

  'setting the list visibility to false
  'increases clear and load time
   List1.Visible = False
   List1.Clear
   
  'get start time, folders, and finish time
   tstart = GetTickCount()
   Call SearchForFolders(FP)
   tend = GetTickCount()
   
   List1.Visible = True
   
  'show the results
   Text3.Text = Format$(FP.Count, "###,###,###,##0") & _
                        " found (" & _
                        FP.sFileNameExt & ")"
                   
   Text4.Text = FormatNumber((tend - tstart) / 1000, 2) & "  seconds"

End Sub


Private Sub SearchForFolders(FP As FILE_PARAMS)

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sRoot As String
   Dim spath As String
   Dim sTmp As String
   
   sRoot = QualifyPath(FP.sFileRoot)
   spath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first match
   hFile = FindFirstFile(spath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
         
      Do
         
        'Only folders are wanted, so discard files
        'or parent/root DOS folders.
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) And _
             Asc(WFD.cFileName) <> vbDot Then
            
           'must be a folder, so remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
                       
           'This is where you add code to store
           'or display the returned file listing.
           '
           'if you want the folder name only, save 'sTmp'.
           'if you want the full path, save 'sRoot & sTmp'
            FP.Count = FP.Count + 1
            List1.AddItem sRoot & sTmp
            
           'if a recursive search was selected, call
           'this method again with a modified root
            If FP.bRecurse Then
            
               FP.sFileRoot = sRoot & sTmp
               Call SearchForFolders(FP)
            
            End If

         End If
         
      Loop While FindNextFile(hFile, WFD)
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub


Private Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function


Private Function QualifyPath(spath As String) As String

  'assures that a passed path ends in a slash
   If Right$(spath, 1) <> "\" Then
      QualifyPath = spath & "\"
   Else
      QualifyPath = spath
   End If
      
End Function


