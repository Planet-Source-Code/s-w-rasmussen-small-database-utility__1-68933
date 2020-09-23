VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Record 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   0  'None
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Exit Function"
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
      Left            =   11220
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   4980
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4980
      Width           =   900
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   13185
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   13185
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   615
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   13185
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   885
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   13185
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1140
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   13185
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1410
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   13185
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1665
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   13185
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1935
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   13185
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2190
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   13185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2460
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   13185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2730
      Width           =   195
   End
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0000
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   11
      Top             =   870
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0075
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Top             =   1350
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Record.frx":00EA
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   13
      Top             =   1620
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":015F
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   14
      Top             =   1620
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":01D4
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   15
      Top             =   2160
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0249
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   16
      Top             =   2430
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":02BE
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   12
      Left            =   2040
      TabIndex        =   17
      Top             =   3450
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0333
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   16
      Left            =   2040
      TabIndex        =   18
      ToolTipText     =   "Shift-Click to launch"
      Top             =   4740
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":03A9
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   17
      Left            =   2040
      TabIndex        =   19
      ToolTipText     =   "Shift-Click to launch"
      Top             =   5010
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":041F
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
   Begin RichTextLib.RichTextBox txtComments 
      Height          =   2070
      Left            =   5970
      TabIndex        =   20
      Top             =   3195
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3651
      _Version        =   393217
      BackColor       =   12972786
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Record.frx":0495
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   18
      Left            =   7800
      TabIndex        =   21
      Top             =   330
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0514
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   19
      Left            =   7800
      TabIndex        =   22
      Top             =   600
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":058A
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   20
      Left            =   7800
      TabIndex        =   23
      Top             =   870
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0600
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   23
      Left            =   7815
      TabIndex        =   24
      Top             =   1890
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0676
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   24
      Left            =   7815
      TabIndex        =   25
      Top             =   2160
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":06EC
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   25
      Left            =   7815
      TabIndex        =   26
      Top             =   2430
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0762
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   3720
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":07D8
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   21
      Left            =   7800
      TabIndex        =   28
      Top             =   1140
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":084E
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   22
      Left            =   7800
      TabIndex        =   29
      Top             =   1410
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":08C4
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   26
      Left            =   7815
      TabIndex        =   30
      Top             =   2700
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":093A
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4470
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":09B0
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   14
      Left            =   2040
      TabIndex        =   32
      Top             =   4200
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0A26
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   33
      Top             =   3180
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0A9C
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   34
      Top             =   2910
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0B12
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   35
      Top             =   1890
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0B88
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
   Begin RichTextLib.RichTextBox txtField 
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   81
      Top             =   330
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16250356
      Enabled         =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Record.frx":0BFD
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
   Begin VB.Line LineR 
      BorderColor     =   &H00800000&
      X1              =   13480
      X2              =   13480
      Y1              =   0
      Y2              =   5415
   End
   Begin VB.Line LineB 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   13515
      Y1              =   5380
      Y2              =   5380
   End
   Begin VB.Line LineL 
      BorderColor     =   &H00800000&
      X1              =   20
      X2              =   20
      Y1              =   0
      Y2              =   5415
   End
   Begin VB.Line LineT 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   13515
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Fornavn / Kontaktpers."
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   79
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Efternavn"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   78
      Top             =   630
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Stilling, Titel"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   77
      Top             =   900
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Vejnavn"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   76
      Top             =   1380
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Husnummer"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   75
      Top             =   1650
      Width           =   1800
   End
   Begin VB.Label lblField 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Etage (tv, th...)"
      Height          =   195
      Index           =   6
      Left            =   3960
      TabIndex        =   74
      Top             =   1650
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postdistrikt"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   73
      Top             =   2190
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postnummer"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   72
      Top             =   2460
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kirke, Sogn"
      Height          =   195
      Index           =   12
      Left            =   180
      TabIndex        =   71
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "E-mail"
      Height          =   195
      Index           =   16
      Left            =   180
      TabIndex        =   70
      Top             =   4770
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hjemmeside"
      Height          =   195
      Index           =   17
      Left            =   180
      TabIndex        =   69
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Label lblField 
      BackColor       =   &H009BE8E8&
      Caption         =   " Bemærkninger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   5895
      TabIndex        =   68
      Top             =   2970
      Width           =   5220
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FDD6C6&
      Caption         =   " Navn:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   120
      Width           =   5520
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FDD6C6&
      Caption         =   " Adresse:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   66
      Top             =   1140
      Width           =   5520
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FDD6C6&
      Caption         =   " Virksomhed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   65
      Top             =   2700
      Width           =   5520
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   18
      Left            =   5940
      TabIndex        =   64
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Brænde (sort, længde)"
      Enabled         =   0   'False
      Height          =   195
      Index           =   19
      Left            =   5940
      TabIndex        =   63
      Top             =   630
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kvantum, stablede m3:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   20
      Left            =   5940
      TabIndex        =   62
      Top             =   900
      Width           =   1800
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      Caption         =   " Brænde:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   5895
      TabIndex        =   61
      ToolTipText     =   " To enable: Checkmark first keyword "
      Top             =   120
      Width           =   5220
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   23
      Left            =   5955
      TabIndex        =   60
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hegntype, kvantum"
      Enabled         =   0   'False
      Height          =   195
      Index           =   24
      Left            =   5955
      TabIndex        =   59
      Top             =   2190
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   25
      Left            =   5955
      TabIndex        =   58
      Top             =   2460
      Width           =   1800
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H009FBF9F&
      Caption         =   " Pileflet:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5895
      TabIndex        =   57
      ToolTipText     =   " To enable: Checkmark second keyword "
      Top             =   1680
      Width           =   5220
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Donationer"
      Height          =   195
      Index           =   13
      Left            =   180
      TabIndex        =   56
      Top             =   3750
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   21
      Left            =   5940
      TabIndex        =   55
      Top             =   1170
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   22
      Left            =   5940
      TabIndex        =   54
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   26
      Left            =   5955
      TabIndex        =   53
      Top             =   2730
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, mobil"
      Height          =   195
      Index           =   15
      Left            =   180
      TabIndex        =   52
      Top             =   4500
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, fastnet"
      Height          =   195
      Index           =   14
      Left            =   180
      TabIndex        =   51
      Top             =   4230
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Branche, Type"
      Height          =   195
      Index           =   11
      Left            =   180
      TabIndex        =   50
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Virksomhed"
      Height          =   195
      Index           =   10
      Left            =   180
      TabIndex        =   49
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bynavn"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   48
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FDD6C6&
      Caption         =   " Kontakt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   47
      Top             =   3990
      Width           =   5520
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H000040C0&
      Caption         =   " Nøgleord:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   11325
      TabIndex        =   46
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Frivillig hjælper"
      Height          =   195
      Index           =   37
      Left            =   11370
      TabIndex        =   45
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bidragyder"
      Height          =   195
      Index           =   36
      Left            =   11370
      TabIndex        =   44
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "SH-Grafik kunde"
      Height          =   195
      Index           =   35
      Left            =   11370
      TabIndex        =   43
      Top             =   2190
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Forretningsforb."
      Height          =   195
      Index           =   34
      Left            =   11370
      TabIndex        =   42
      Top             =   1935
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Årsberetning"
      Height          =   195
      Index           =   33
      Left            =   11370
      TabIndex        =   41
      Top             =   1665
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Julehilsen"
      Height          =   195
      Index           =   32
      Left            =   11370
      TabIndex        =   40
      Top             =   1410
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Høstfest"
      Height          =   195
      Index           =   31
      Left            =   11370
      TabIndex        =   39
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Sommerfest"
      Height          =   195
      Index           =   30
      Left            =   11370
      TabIndex        =   38
      Top             =   885
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009FBF9F&
      Caption         =   "Pilefletkunde"
      Height          =   195
      Index           =   29
      Left            =   11370
      TabIndex        =   37
      Top             =   615
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0098C3C3&
      Caption         =   "Brændekunde"
      Height          =   195
      Index           =   28
      Left            =   11370
      TabIndex        =   36
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload RecordList
    Me.Visible = False
    
End Sub

Private Sub cmdExit_Click()

    Unload RecordList
    If ComposeList.Visible Then Unload ComposeList
    Me.Visible = False
    
End Sub


Private Sub Form_Activate()
    
    Me.Top = RecordList.Height + 100
    
End Sub

Private Sub Form_Load()

Dim N As Long
        
    ' field labels
    For N = 1 To 37
        lblField(N).Caption = Main.lblField(N).Caption
    Next N
    
    For N = 1 To 7
        lblCaption(N).Caption = Main.lblCaption(N).Caption
    Next N
    
    Me.Hide
     
End Sub


Private Sub Form_Resize()
    
    Me.Top = RecordList.Height + 100
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
            
    Main.Visible = True
    
End Sub


