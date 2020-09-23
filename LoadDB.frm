VERSION 5.00
Begin VB.Form LoadDB 
   BackColor       =   &H00A38983&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading database,  please wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ECE6E5&
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   3400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Small Database Utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ECE6E5&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3400
   End
End
Attribute VB_Name = "LoadDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim Success                 As Boolean
    
    ' declare arrays
    ReDim nr(1 To 1) As RECORD_DATA
    ReDim sr(1 To 2, 1 To 1)
    ReDim captions(1 To 80)
    ReDim users(1 To 2, 1 To 6)
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    Me.Show
            
    MAIN_DIR_G = GetSetting("SDU_UK", "User", "DataPath", vbNullString)
    Success = Compressed_Database_Read(MAIN_DIR_G & "database.zlb") '????!!!!
    
    Me.Hide
    Welcome.Hide
    'Unload Me
    
End Sub


