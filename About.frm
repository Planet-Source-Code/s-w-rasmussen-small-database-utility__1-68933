VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4020
      Picture         =   "About.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4160
      TabIndex        =   6
      Top             =   2850
      Width           =   460
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   2400
      Width           =   4155
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copyright 2011 by S.W. Rasmussen. All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Use it to keep track of your friends and relatives - and to easily retrieve personal information about them."
      Height          =   555
      Left            =   300
      TabIndex        =   2
      Top             =   1380
      Width           =   3600
   End
   Begin VB.Label lblAboutTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Small Database Utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label lblAboutTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "fully configurable..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   4155
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

On Error GoTo errhandler
        
    form_StayOnTop About, True, "C"
        
    lblVersion.Caption = "version " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000")
    lblAboutTitle(1).Caption = "Small Database Utility"
    lblAboutTitle(2).Caption = "fully configurable..."
    
errhandler:
    Exit Sub
End Sub


Private Sub Label5_Click()

On Error GoTo errhandler

    Unload Me
    
errhandler:
    Exit Sub
End Sub


