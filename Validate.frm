VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Validate 
   BackColor       =   &H00EEE8E6&
   Caption         =   " Validaion Result"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Validate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
      Height          =   300
      Left            =   5340
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3020
      Width           =   900
   End
   Begin MSComctlLib.ListView lvwValidate 
      Height          =   2800
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   " Left-Click: First Record, Right-Click:Second Record  "
      Top             =   120
      Width           =   6140
      _ExtentX        =   10821
      _ExtentY        =   4948
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "Validate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub


Private Sub Form_Load()

On Error GoTo errhandler

    form_StayOnTop Validate, True, "TR"
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Resize()

On Error GoTo errhandler

    If Me.Height < 4000 Then Me.Height = 4000
    If Me.Width < 6000 Then Me.Width = 6000
    
    lvwValidate.Top = 90
    lvwValidate.Left = 90
    lvwValidate.Height = Me.Height - 1020
    lvwValidate.Width = Me.Width - 300
    cmdClose.Left = Me.Width - 1130
    cmdClose.Top = Me.Height - 820
    
errhandler:
Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errhandler
    
errhandler:
    Exit Sub
End Sub

Private Sub lvwValidate_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
On Error GoTo errhandler
        
    ' If the ListView is already sorted by the clicked column, just reverse the order. Otherwise, sort the clicked column ascending.
    If lvwValidate.Sorted = True And ColumnHeader.SubItemIndex = lvwValidate.SortKey Then
        If lvwValidate.SortOrder = lvwAscending Then
            lvwValidate.SortOrder = lvwDescending
        Else
            lvwValidate.SortOrder = lvwAscending
        End If
    Else
        lvwValidate.Sorted = True
        lvwValidate.SortKey = ColumnHeader.SubItemIndex
        lvwValidate.SortOrder = lvwAscending
    End If
        
errhandler:
    Exit Sub
End Sub


Private Sub lvwValidate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler
    
    CurrRecord_G = lvwValidate.ListItems(lvwValidate.SelectedItem.Index).ListSubItems(Button)
        
    Call Record_ShowSingleA(CurrRecord_G)
    
errhandler:
    Exit Sub
End Sub


