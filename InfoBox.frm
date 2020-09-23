VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form InfoBox 
   BackColor       =   &H00E0FFFF&
   Caption         =   "My Personal Notes"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InfoBox.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7230
   Begin RichTextLib.RichTextBox txtNotes 
      Height          =   6435
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11351
      _Version        =   393217
      BackColor       =   14745599
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"InfoBox.frx":08CA
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
End
Attribute VB_Name = "InfoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

On Error GoTo errhandler
    
    form_StayOnTop InfoBox, True, "C"
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Resize()
    
    txtNotes.Left = 90
    txtNotes.Top = 90
    txtNotes.Height = Me.Height - 980
    txtNotes.Width = Me.Width - 300
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errhandler

Dim fil                     As Long

Const CRLF1 = "\par " & vbCrLf & "\par " & vbCrLf & "\par }"
Const CRLF2 = "\par " & vbCrLf & "\par }"
    
    If Len(InfoBox.txtNotes.Text) > 0 Then
    
        ' clean rtf string of exces crlf's
        Do While InStr(InfoBox.txtNotes.TextRTF, CRLF1)
            InfoBox.txtNotes.TextRTF = Replace(InfoBox.txtNotes.TextRTF, CRLF1, CRLF2)
        Loop
        
        ' save personal note for current record
        fil = FreeFile
        Open NOTES_DIR_G & nr(CurrRecord_G).ID & ".note" For Output As #fil
            Print #fil, InfoBox.txtNotes.TextRTF
        Close #fil
    Else
        InfoBox.txtNotes.Text = vbNullString
        If FileExist(NOTES_DIR_G & nr(CurrRecord_G).ID & ".note") Then
            Kill NOTES_DIR_G & nr(CurrRecord_G).ID & ".note"
        End If
    End If
    
errhandler:
    Exit Sub
End Sub


