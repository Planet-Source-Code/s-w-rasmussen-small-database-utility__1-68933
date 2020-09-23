VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form MasterUsers 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Administrators"
   ClientHeight    =   2985
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
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5715
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   180
      Picture         =   "Users.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2340
      Width           =   510
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Cancel"
      Height          =   300
      Index           =   1
      Left            =   3730
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2535
      Width           =   900
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Accept"
      Height          =   300
      Index           =   0
      Left            =   4690
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2535
      Width           =   900
   End
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   1
      Left            =   4300
      TabIndex        =   1
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":0E15
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":0E8C
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":0F03
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
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   2
      Left            =   4300
      TabIndex        =   3
      Top             =   660
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":0F7A
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":0FF1
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
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   3
      Left            =   4300
      TabIndex        =   5
      Top             =   960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":1068
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   1260
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":10DF
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
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   4
      Left            =   4300
      TabIndex        =   7
      Top             =   1260
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":1156
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   8
      Top             =   1560
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":11CD
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
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   5
      Left            =   4300
      TabIndex        =   10
      Top             =   1560
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":1244
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
   Begin RichTextLib.RichTextBox txtFullName 
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   11
      Top             =   1860
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":12BB
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
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   255
      Index           =   6
      Left            =   4300
      TabIndex        =   12
      Top             =   1860
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Users.frx":1332
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4300
      TabIndex        =   15
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "MasterUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdAction_Click(Index As Integer)

On Error GoTo errhandler

Dim N                       As Long
    
    If Index = 0 Then
        For N = 1 To 6
            users(1, N) = Trim$(txtFullName(N).Text)
            users(2, N) = Trim$(txtPassword(N).Text)
        Next N
    End If
    
    Unload Me
    
errhandler:
    Exit Sub
End Sub




Private Sub Form_Load()

On Error GoTo errhandler

Call form_StayOnTop(MasterUsers, True, "C")

Dim N                       As Long

    For N = 1 To 6
        txtFullName(N).Text = users(1, N)
        txtPassword(N).Text = users(2, N)
    Next N
       
errhandler:
    Exit Sub
End Sub



