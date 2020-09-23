VERSION 5.00
Begin VB.Form Passwrd 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Password"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Passwrd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "Passwrd.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   480
      Width           =   510
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   915
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   780
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "txtPassword"
      Top             =   690
      Width           =   1800
   End
   Begin VB.Label lblPassWordMsg 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "Enter password to clear Lock and enable Edit"
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   3285
   End
End
Attribute VB_Name = "Passwrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click(Index As Integer)

On Error GoTo errhandler
    
    Select Case Index
        Case 0: Password_G = txtPassword.Text
        Case 1: Password_G = vbNullString
    End Select
    
    Me.Hide
    DoEvents
    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Activate()
    
On Error GoTo errhandler
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    Me.Refresh
    txtPassword.SetFocus
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler
    
    form_StayOnTop Passwrd, True, "C"
    Password_G = vbNullString
    
errhandler:
    Exit Sub
End Sub


    
