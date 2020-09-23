VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form UploadInfo 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Upload Settings"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UploadInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5715
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   1
      Left            =   2260
      TabIndex        =   3
      Top             =   780
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   180
      Picture         =   "UploadInfo.frx":0941
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3000
      Width           =   510
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3730
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3200
      Width           =   900
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   4690
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3200
      Width           =   900
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   0
      Left            =   2260
      TabIndex        =   2
      ToolTipText     =   " The IP address of the webserver "
      Top             =   480
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":0E8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   2
      Left            =   2260
      TabIndex        =   4
      Top             =   1080
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":0F03
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   3
      Left            =   2260
      TabIndex        =   5
      Top             =   1380
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":0F7A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   4
      Left            =   2260
      TabIndex        =   6
      ToolTipText     =   " The path to the remote database on the webserver "
      Top             =   1680
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":0FF1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   5
      Left            =   2260
      TabIndex        =   7
      ToolTipText     =   " The local filename of the database "
      Top             =   1980
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   0   'False
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":1068
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   6
      Left            =   2260
      TabIndex        =   1
      ToolTipText     =   " The website where the database is stored "
      Top             =   180
      Width           =   3340
      _ExtentX        =   5900
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":10E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   " The URL to the SDU webpage  "
      Top             =   2280
      Width           =   3800
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":115F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUpload 
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   " The URL for download of setup file "
      Top             =   2580
      Width           =   3800
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"UploadInfo.frx":11D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Update URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   2610
      Width           =   1500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Version Info URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   2310
      Width           =   1500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Website URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   220
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Local filename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2010
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Remote folderpath:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1710
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Remote filename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   " The file name of the of the remote copy of the database "
      Top             =   1412
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1114
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   816
      Width           =   2000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Remote server IP adr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   518
      Width           =   2000
   End
End
Attribute VB_Name = "UploadInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)

On Error GoTo errhandler

    If Index = 0 Then
        With uploaddb
            .RemoteServerIP = Trim$(UploadInfo.txtUpload(0).Text)
            .UserName = StringEncode(Trim$(UploadInfo.txtUpload(1).Text))   ' the username is encrypted at this point
            .PassWord = StringEncode(Trim$(UploadInfo.txtUpload(2).Text))   ' the password is encrypted at this point
            .RemoteFileName = Trim$(UploadInfo.txtUpload(3).Text)
            .RemoteFolderPath = Trim$(UploadInfo.txtUpload(4).Text)
            .LocalFileName = Trim$(UploadInfo.txtUpload(5).Text)
            .WebsiteURL = Trim$(UploadInfo.txtUpload(6).Text)
            .ProgramInfoURL = Trim$(UploadInfo.txtUpload(7).Text)
            .UpdateExeURL = Trim$(UploadInfo.txtUpload(8).Text)
        End With
        
        ' remove leading "/" on serverpath if it is present
        If Left$(uploaddb.RemoteFolderPath, 1) = "/" Then
            uploaddb.RemoteFolderPath = Mid$(uploaddb.RemoteFolderPath, 2)
        End If
        
        ' make sure that url ends with "/"
        If Right$(uploaddb.WebsiteURL, 1) <> "/" Then
            uploaddb.WebsiteURL = uploaddb.WebsiteURL & "/"
        End If
        
        ' make sure that url ends with "/"
        If Right$(uploaddb.UpdateExeURL, 1) <> "/" Then
            uploaddb.UpdateExeURL = uploaddb.UpdateExeURL & "/"
        End If
        
        ' Replace æ, ø, å in remote and local filenames
        uploaddb.RemoteFileName = GetRemoteFileName(uploaddb.RemoteFileName)
        uploaddb.LocalFileName = "database.zlb"
        UploadInfo.txtUpload(5).Text = uploaddb.LocalFileName
                
    End If
    
    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler

    Call form_StayOnTop(UploadInfo, True, "C")
                
    txtUpload(0).Text = uploaddb.RemoteServerIP
    txtUpload(1).Text = StringDecode(uploaddb.UserName) ' decode username before displaying it
    txtUpload(2).Text = StringDecode(uploaddb.PassWord) ' decode password before displaying it
    txtUpload(3).Text = uploaddb.RemoteFileName
    txtUpload(4).Text = uploaddb.RemoteFolderPath
    txtUpload(5).Text = "database.zlb" 'uploaddb.LocalFileName
    txtUpload(6).Text = uploaddb.WebsiteURL
    txtUpload(7).Text = uploaddb.ProgramInfoURL
    txtUpload(8).Text = uploaddb.UpdateExeURL
    
errhandler:
    Exit Sub
End Sub

