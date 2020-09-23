VERSION 5.00
Begin VB.Form DefaultKeyWord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Set Default KeyWord"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
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
   Icon            =   "DefaultKeyWord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Accept"
      Height          =   300
      Left            =   2430
      TabIndex        =   2
      Top             =   1440
      Width           =   720
   End
   Begin VB.ComboBox cmbKeyWords 
      Height          =   330
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label lblLegend 
      Caption         =   "Select default KeyWord for new records. The default KeyWord is valid for this session only."
      Height          =   735
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   3015
   End
End
Attribute VB_Name = "DefaultKeyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbKeyWords_Change()

On Error GoTo errhandler
    
    KeyWordLabelIndex_G = cmbKeyWords.ListIndex + 28
    
errhandler:
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    
On Error GoTo errhandler

    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()
   
On Error Resume Next

    ' load keyword labels from Main form into dropdownlist: cmdKeyWords
    cmbKeyWords.AddItem Main.lblField(28).Caption   ' index = 0
    cmbKeyWords.AddItem Main.lblField(29).Caption   ' index = 1
    cmbKeyWords.AddItem Main.lblField(30).Caption   ' index = 2
    cmbKeyWords.AddItem Main.lblField(31).Caption   ' index = 3
    cmbKeyWords.AddItem Main.lblField(32).Caption   ' index = 4
    cmbKeyWords.AddItem Main.lblField(33).Caption   ' index = 5
    cmbKeyWords.AddItem Main.lblField(34).Caption   ' index = 6
    cmbKeyWords.AddItem Main.lblField(35).Caption   ' index = 7
    cmbKeyWords.AddItem Main.lblField(36).Caption   ' index = 8
    cmbKeyWords.AddItem Main.lblField(37).Caption   ' index = 9
    
    ' select default KeyWord
    cmbKeyWords.ListIndex = KeyWordLabelIndex_G - 28
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errhandler

    KeyWordLabelIndex_G = cmbKeyWords.ListIndex + 28
    
    ' display default keyword om main form
    Main.lblDefaultKeyWord.Caption = "default = " & Trim$(cmbKeyWords.Text)
    
errhandler:
    Exit Sub
End Sub


