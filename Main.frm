VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "vlmnuplus.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Small Database Utility"
   ClientHeight    =   7230
   ClientLeft      =   135
   ClientTop       =   735
   ClientWidth     =   13650
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
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   910
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2520
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08CA
            Key             =   "DATABASEOPEN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0E64
            Key             =   "UNLOCKED"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":13FE
            Key             =   "RESIDENT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1998
            Key             =   "RESIDENT_OFF"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1F32
            Key             =   "HIDE_FORM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24CC
            Key             =   "DATABASECLOSE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2A66
            Key             =   "REC"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3000
            Key             =   "OPENREC"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":359A
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":36F4
            Key             =   "ABOUT"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3C8E
            Key             =   "FORMS"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4228
            Key             =   "DOWNLOADSV"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4382
            Key             =   "WEBOPEN"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44DC
            Key             =   "EXIST"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4A76
            Key             =   "EDIT"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5010
            Key             =   "PERSONLISTE"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":55AA
            Key             =   "BLANK"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5B44
            Key             =   "NOTES"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":60DE
            Key             =   "EXCEL"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6678
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6C12
            Key             =   "COMPRESS"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":71AC
            Key             =   "TOOLS"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7746
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7CE0
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":827A
            Key             =   "SELCOLUMN"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8814
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":896E
            Key             =   "SAVE1"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8F08
            Key             =   "SAVE2"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":94A2
            Key             =   "CLOSE1"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":95FC
            Key             =   "EMPTY"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9B96
            Key             =   "DATABASE"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A130
            Key             =   "LOCKED"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAction 
      DisabledPicture =   "Main.frx":A6CA
      DownPicture     =   "Main.frx":AB23
      Height          =   375
      Index           =   9
      Left            =   165
      Picture         =   "Main.frx":AF7C
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   " Parks form in tray without closing program "
      Top             =   6450
      Width           =   375
   End
   Begin VB.CommandButton cmdNotes 
      BackColor       =   &H00AEFFFF&
      Caption         =   "My private notes..."
      Height          =   300
      Left            =   11520
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   6090
      UseMaskColor    =   -1  'True
      Width           =   1890
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   2880
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2614
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   2351
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   2088
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1825
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1562
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   1299
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1036
      Width           =   195
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   773
      Width           =   195
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00FDD6C6&
      Caption         =   "Show Hit List"
      Height          =   300
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   5625
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   78
      Top             =   6930
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
         EndProperty
      EndProperty
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
   Begin RichTextLib.RichTextBox txtJump 
      Height          =   270
      Left            =   10755
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   6480
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   476
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":B3D5
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
   Begin VB.CommandButton cmdAction 
      Caption         =   "Bookmark"
      Height          =   300
      Index           =   6
      Left            =   11500
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   " Shift+Click to set bookmark, Click to jump to bookmark "
      Top             =   6480
      Width           =   1200
   End
   Begin RichTextLib.RichTextBox txtSearch 
      Height          =   270
      Index           =   0
      Left            =   7800
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   " Query, case-insensitive "
      Top             =   5625
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   476
      _Version        =   393217
      BackColor       =   16775643
      Enabled         =   0   'False
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B44C
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
   Begin VB.CommandButton cmdNavSearch 
      BackColor       =   &H00FDD6C6&
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   13100
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   " Next, Right-Click: Last "
      Top             =   5625
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.CommandButton cmdNavSearch 
      BackColor       =   &H00FDD6C6&
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   12760
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   " Previous, Right-Click: First "
      Top             =   5625
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   210
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "New"
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   7380
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   " Create new Record "
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0000C000&
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   13275
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   " Next, Right-Click: Last "
      Top             =   6585
      Width           =   300
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0000C000&
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   12945
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   " Previous, Right-Click: First "
      Top             =   6585
      Width           =   300
   End
   Begin VB.CheckBox chkKeyWord 
      Alignment       =   1  'Right Justify
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
      Left            =   13250
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   510
      Width           =   195
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   6660
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   " Save current Record "
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   5940
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   " Delete current Record "
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Compose List..."
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   8100
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1350
   End
   Begin RichTextLib.RichTextBox txtField 
      Height          =   260
      Index           =   1
      Left            =   2100
      TabIndex        =   0
      Top             =   480
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B4C3
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
      Height          =   260
      Index           =   2
      Left            =   2100
      TabIndex        =   1
      Top             =   750
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B538
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
      Height          =   260
      Index           =   3
      Left            =   2100
      TabIndex        =   2
      Top             =   1020
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B5AD
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
      Left            =   2100
      TabIndex        =   3
      Top             =   1500
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B622
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
      Left            =   2100
      TabIndex        =   4
      Top             =   1770
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B697
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
      Left            =   3060
      TabIndex        =   5
      Top             =   1770
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B70C
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
      Left            =   2100
      TabIndex        =   7
      Top             =   2310
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B781
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
      Left            =   2100
      TabIndex        =   8
      Top             =   2580
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B7F6
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
      Left            =   2100
      TabIndex        =   11
      Top             =   3600
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B86B
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
      Left            =   2100
      TabIndex        =   15
      ToolTipText     =   "Shift-Click to launch"
      Top             =   4890
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B8E1
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
      Left            =   2100
      TabIndex        =   16
      ToolTipText     =   "Shift-Click to launch"
      Top             =   5160
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":B957
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
      Left            =   6025
      TabIndex        =   26
      Top             =   3345
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3651
      _Version        =   393217
      BackColor       =   12972786
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Main.frx":B9CD
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
      Height          =   260
      Index           =   18
      Left            =   7860
      TabIndex        =   17
      Top             =   480
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BA4C
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
      Height          =   260
      Index           =   19
      Left            =   7860
      TabIndex        =   18
      Top             =   750
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BAC2
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
      Height          =   260
      Index           =   20
      Left            =   7860
      TabIndex        =   19
      Top             =   1020
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BB38
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
      Left            =   7875
      TabIndex        =   22
      Top             =   2040
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BBAE
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
      Left            =   7875
      TabIndex        =   23
      Top             =   2310
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BC24
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
      Left            =   7875
      TabIndex        =   24
      Top             =   2580
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BC9A
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
      Left            =   2100
      TabIndex        =   12
      Top             =   3870
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":BD10
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
      Height          =   260
      Index           =   21
      Left            =   7860
      TabIndex        =   20
      Top             =   1290
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BD86
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
      Height          =   260
      Index           =   22
      Left            =   7860
      TabIndex        =   21
      Top             =   1560
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BDFC
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
      Left            =   7875
      TabIndex        =   25
      Top             =   2850
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Main.frx":BE72
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
      Left            =   2100
      TabIndex        =   14
      Top             =   4620
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":BEE8
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
      Left            =   2100
      TabIndex        =   13
      Top             =   4350
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":BF5E
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
      Left            =   2100
      TabIndex        =   10
      Top             =   3330
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":BFD4
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
      Left            =   2100
      TabIndex        =   9
      Top             =   3060
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":C04A
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
      Left            =   2100
      TabIndex        =   6
      Top             =   2040
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":C0C0
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
   Begin RichTextLib.RichTextBox txtSearch 
      Height          =   270
      Index           =   1
      Left            =   7800
      TabIndex        =   101
      TabStop         =   0   'False
      ToolTipText     =   " Query, case-insensitive "
      Top             =   5880
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   476
      _Version        =   393217
      BackColor       =   12972786
      Enabled         =   0   'False
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Main.frx":C135
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
   Begin VB.Label lblStar 
      BackColor       =   &H00EEE8E6&
      Caption         =   " *    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   5715
      TabIndex        =   118
      ToolTipText     =   " Opens web browser on click "
      Top             =   5190
      Width           =   285
   End
   Begin VB.Label lblStar 
      BackColor       =   &H00EEE8E6&
      Caption         =   " *    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5715
      TabIndex        =   117
      ToolTipText     =   " Opens web browser on click "
      Top             =   4925
      Width           =   285
   End
   Begin VB.Label lblStar 
      BackColor       =   &H00EEE8E6&
      Caption         =   " *    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5715
      TabIndex        =   116
      ToolTipText     =   " Opens e-mail program on click "
      Top             =   4660
      Width           =   285
   End
   Begin VB.Label lblStar 
      BackColor       =   &H00EEE8E6&
      Caption         =   " *    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5715
      TabIndex        =   115
      ToolTipText     =   " Opens e-mail program on click "
      Top             =   4395
      Width           =   285
   End
   Begin VB.Label lblDefaultKeyWord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   11460
      TabIndex        =   114
      Top             =   3135
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "Edit:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4215
      TabIndex        =   113
      ToolTipText     =   " Local and Web database is different "
      Top             =   6525
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "WebLock:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2640
      TabIndex        =   112
      ToolTipText     =   " Local and Web database is different "
      Top             =   6525
      Width           =   855
   End
   Begin VB.Label lblWebLockStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   3645
      TabIndex        =   111
      ToolTipText     =   "Edit status of databasefile on the website is unknown..."
      Top             =   6525
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "Source:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   900
      TabIndex        =   110
      ToolTipText     =   " Local and Web database is different "
      Top             =   6525
      Width           =   660
   End
   Begin VB.Label lblLastSaved 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "file date"
      Height          =   210
      Left            =   240
      TabIndex        =   109
      Top             =   6120
      Width           =   570
   End
   Begin VB.Label lblResident 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      Caption         =   "running resident"
      Height          =   210
      Left            =   240
      TabIndex        =   108
      Top             =   5880
      Width           =   1170
   End
   Begin VB.Label lblWebDBVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1650
      TabIndex        =   107
      ToolTipText     =   " Database source "
      Top             =   6525
      Width           =   45
   End
   Begin VLMnuPlus.VLMenuPlus VLMenuPlus1 
      Left            =   4380
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   40777.3413425926
      AutoShowHelp    =   0   'False
      UseCustomColors =   -1  'True
      TextColor       =   0
      DisabledTextColor=   14737632
      HighlightedTextColor=   2322544
      MenuHighlight   =   12972786
      MenuHighlightBorder=   14737632
      TootipBackground=   12648447
      TootipTextColor =   0
      Language        =   0
      ShowTooltip     =   0   'False
   End
   Begin VB.Label lblApplicationTitle 
      BackColor       =   &H00EEE8E6&
      Caption         =   "undertitel..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   11325
      TabIndex        =   105
      Top             =   5265
      Width           =   2175
   End
   Begin VB.Label lblApplicationTitle 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Person database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   11325
      TabIndex        =   104
      Top             =   3360
      Width           =   2160
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Left            =   11325
      Stretch         =   -1  'True
      Top             =   3630
      Width           =   2175
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFC0C0&
      X1              =   899.333
      X2              =   899.333
      Y1              =   400
      Y2              =   368
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFC0C0&
      X1              =   636
      X2              =   636
      Y1              =   400
      Y2              =   416
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFC0C0&
      X1              =   636
      X2              =   900
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFC0C0&
      X1              =   395
      X2              =   395
      Y1              =   368
      Y2              =   416
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFC0C0&
      X1              =   395
      X2              =   636
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFC0C0&
      X1              =   395
      X2              =   900
      Y1              =   368
      Y2              =   368
   End
   Begin VB.Line Line6 
      BorderColor     =   &H009FBF9F&
      X1              =   873.333
      X2              =   752
      Y1              =   65.667
      Y2              =   65.667
   End
   Begin VB.Line Line5 
      BorderColor     =   &H009FBF9F&
      X1              =   752
      X2              =   752
      Y1              =   129.333
      Y2              =   65.333
   End
   Begin VB.Line Line4 
      BorderColor     =   &H009FBF9F&
      X1              =   744
      X2              =   752
      Y1              =   129.333
      Y2              =   129.333
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0098C3C3&
      X1              =   752
      X2              =   752
      Y1              =   25.333
      Y2              =   49
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0098C3C3&
      X1              =   873.333
      X2              =   752
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0098C3C3&
      X1              =   744
      X2              =   752
      Y1              =   25.333
      Y2              =   25.333
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Search Comments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   6015
      TabIndex        =   102
      Top             =   5910
      Width           =   1590
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0098C3C3&
      Caption         =   "Brændekunde"
      Height          =   195
      Index           =   28
      Left            =   11425
      TabIndex        =   100
      Top             =   510
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009FBF9F&
      Caption         =   "Pilefletkunde"
      Height          =   195
      Index           =   29
      Left            =   11430
      TabIndex        =   99
      Top             =   773
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Sommerfest"
      Height          =   195
      Index           =   30
      Left            =   11430
      TabIndex        =   98
      Top             =   1036
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Høstfest"
      Height          =   195
      Index           =   31
      Left            =   11430
      TabIndex        =   97
      Top             =   1299
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Julehilsen"
      Height          =   195
      Index           =   32
      Left            =   11430
      TabIndex        =   96
      Top             =   1562
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Årsberetning"
      Height          =   195
      Index           =   33
      Left            =   11430
      TabIndex        =   95
      Top             =   1825
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Forretningsforb."
      Height          =   195
      Index           =   34
      Left            =   11430
      TabIndex        =   94
      Top             =   2088
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "SH-Grafik kunde"
      Height          =   195
      Index           =   35
      Left            =   11430
      TabIndex        =   93
      Top             =   2351
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bidragyder"
      Height          =   195
      Index           =   36
      Left            =   11430
      TabIndex        =   92
      Top             =   2614
      Width           =   1695
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Frivillig hjælper"
      Height          =   195
      Index           =   37
      Left            =   11430
      TabIndex        =   91
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblUniqueID 
      BackColor       =   &H00EEE8E6&
      Caption         =   "089787698769876986"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   90
      Top             =   5640
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEE8E6&
      Caption         =   "ID of current record:"
      Height          =   195
      Left            =   240
      TabIndex        =   89
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblLocked 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEE8E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   4665
      TabIndex        =   77
      ToolTipText     =   """ Click to set EDIT to ON """
      Top             =   6525
      Width           =   45
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
      Left            =   11385
      TabIndex        =   74
      Top             =   275
      Width           =   2050
   End
   Begin VB.Label lblNumMatches 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   10755
      TabIndex        =   72
      Top             =   5625
      Width           =   480
   End
   Begin VB.Label lblMatches 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Num. Hits:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9690
      TabIndex        =   71
      Top             =   5655
      Width           =   900
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
      Left            =   180
      TabIndex        =   70
      Top             =   4140
      Width           =   5520
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bynavn"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   69
      Top             =   2070
      Width           =   1800
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Find Records:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   6015
      TabIndex        =   66
      Top             =   5655
      Width           =   1590
   End
   Begin VB.Label lblJump 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Goto Record:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9690
      TabIndex        =   65
      Top             =   6525
      Width           =   1200
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Virksomhed"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   64
      Top             =   3090
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Branche, Type"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   63
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, fastnet"
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   62
      Top             =   4380
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, mobil"
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   61
      Top             =   4650
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   26
      Left            =   6015
      TabIndex        =   57
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   22
      Left            =   6000
      TabIndex        =   56
      Top             =   1590
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   21
      Left            =   6000
      TabIndex        =   55
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Donationer"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   54
      Top             =   3900
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
      Left            =   5955
      TabIndex        =   53
      ToolTipText     =   " To enable: Checkmark second keyword "
      Top             =   1830
      Width           =   5220
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   25
      Left            =   6015
      TabIndex        =   52
      Top             =   2610
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hegntype, kvantum"
      Enabled         =   0   'False
      Height          =   195
      Index           =   24
      Left            =   6015
      TabIndex        =   51
      Top             =   2340
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   23
      Left            =   6015
      TabIndex        =   50
      Top             =   2070
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
      Left            =   5955
      TabIndex        =   49
      ToolTipText     =   " To enable: Checkmark first keyword "
      Top             =   275
      Width           =   5225
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kvantum, stablede m3:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   20
      Left            =   6000
      TabIndex        =   48
      Top             =   1050
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Brænde (sort, længde)"
      Enabled         =   0   'False
      Height          =   195
      Index           =   19
      Left            =   6000
      TabIndex        =   47
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Enabled         =   0   'False
      Height          =   195
      Index           =   18
      Left            =   6000
      TabIndex        =   46
      Top             =   510
      Width           =   1800
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
      Left            =   180
      TabIndex        =   45
      Top             =   2850
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
      Left            =   180
      TabIndex        =   44
      Top             =   1290
      Width           =   5520
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
      Left            =   180
      TabIndex        =   43
      Top             =   275
      Width           =   5525
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
      Left            =   5955
      TabIndex        =   42
      Top             =   3120
      Width           =   5220
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hjemmeside"
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   41
      Top             =   5190
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "E-mail"
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   40
      Top             =   4920
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kirke, Sogn"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   39
      Top             =   3630
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postnummer"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   38
      Top             =   2610
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postdistrikt"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   37
      Top             =   2340
      Width           =   1800
   End
   Begin VB.Label lblField 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Etage (tv, th...)"
      Height          =   195
      Index           =   6
      Left            =   4020
      TabIndex        =   36
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Husnummer"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   35
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Vejnavn"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   34
      Top             =   1530
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Stilling, Titel"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   33
      Top             =   1050
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Efternavn"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Fornavn / Kontaktpers."
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   510
      Width           =   1800
   End
   Begin VB.Menu m_File 
      Caption         =   "File"
      Begin VB.Menu FileItem 
         Caption         =   "Open Database"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu FileItem 
         Caption         =   "Save Database"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu FileItem 
         Caption         =   "Create New Database"
         Index           =   3
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu FileItem 
         Caption         =   "Database - Save Backup"
         Index           =   5
         Shortcut        =   ^B
      End
      Begin VB.Menu FileItem 
         Caption         =   "Database - Load Backup"
         Index           =   6
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu FileItem 
         Caption         =   "Private Notes - Save Backup"
         Index           =   8
         Shortcut        =   ^N
      End
      Begin VB.Menu FileItem 
         Caption         =   "Private Notes - Load Backup"
         Index           =   9
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu FileItem 
         Caption         =   "Upload Database to Website (FTP)"
         Index           =   11
         Shortcut        =   ^U
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu FileItem 
         Caption         =   "Hide Program"
         Index           =   13
         Shortcut        =   ^H
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu FileItem 
         Caption         =   "Exit Program"
         Index           =   15
      End
   End
   Begin VB.Menu m_list 
      Caption         =   "List"
      Begin VB.Menu ListItem 
         Caption         =   "Choose Columns For Record List..."
         Index           =   0
      End
      Begin VB.Menu ListItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu ListItem 
         Caption         =   "Open Record List in Excel Format..."
         Index           =   2
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu m_Tools 
      Caption         =   "Tools"
      Begin VB.Menu ToolsItem 
         Caption         =   "Enable Editing Database..."
         Index           =   0
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "Clear The Edit WebLock..."
         Index           =   1
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "Edit Configuration..."
         Index           =   3
         Shortcut        =   ^Q
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "Advanced Functions..."
         Index           =   5
         Shortcut        =   ^R
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu ToolsItem 
         Caption         =   "Run Resident"
         Index           =   7
      End
   End
   Begin VB.Menu m_DataBases 
      Caption         =   "DataBases"
      Begin VB.Menu DataBaseItem 
         Caption         =   "Existing Databases:"
         Index           =   0
      End
   End
   Begin VB.Menu m_Help 
      Caption         =   "Help"
      Begin VB.Menu HelpItem 
         Caption         =   "Instructions"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu HelpItem 
         Caption         =   "Screenshots..."
         Index           =   2
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu HelpItem 
         Caption         =   "New Version Check..."
         Index           =   4
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu HelpItem 
         Caption         =   "Webserver Connection Check"
         Index           =   6
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu HelpItem 
         Caption         =   "About"
         Index           =   8
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' See:
' http://msdn2.microsoft.com/en-us/library/aa384106.aspx
' for details of WinHttpRequest
'------------------------------------------------------------------------------
Option Explicit

Private txtJumpValue_M      As String
Private navIndex_M          As Long
Private WithEvents mydl     As VicsDL
Attribute mydl.VB_VarHelpID = -1


Private Sub BuildListA(Optional ByVal ShowAll As Boolean = False)

On Error GoTo errhandler

Dim P                       As Long
Dim N                       As Long
Dim rowCount                As Long
Dim itmX                    As Object
            
    If ShowAll Then
        ' clear lvwmain
        RecordList.lvwMain.ColumnHeaders.Clear
        RecordList.lvwMain.ListItems.Clear
        RecordList.lvwMain.MultiSelect = False
        RecordList.lvwMain.SmallIcons = Main.ImageList1
        RecordList.lvwMain.Visible = False
            
        ' add column headers: main.lblfield().caption
        RecordList.lvwMain.ColumnHeaders.add , , "Nr", 700, lvwColumnLeft
        For N = 1 To 26
            If Len(Trim$(lblField(N).Caption)) > 0 Then
                RecordList.lvwMain.ColumnHeaders.add , , lblField(N).Caption, 1500, lvwColumnLeft
            End If
        Next N
                
        ' set View property to Report.
        RecordList.lvwMain.View = lvwReport
                
        ' add list items from url_curr() or url_impt() array
        For N = 1 To UBound(sr, 2)
            Set itmX = RecordList.lvwMain.ListItems.add()
            itmX.Text = Format$(sr(1, N), "0000")
            itmX.SmallIcon = 3
            
            rowCount = 0
            For P = 1 To 26
                If Len(Trim$(lblField(P).Caption)) > 0 Then
                    rowCount = rowCount + 1
                    itmX.SubItems(rowCount) = nr(sr(1, N)).txtField(P)
                End If
            Next P
        Next N
    
    Else
        ' clear lvwmain
        RecordList.lvwMain.ColumnHeaders.Clear
        RecordList.lvwMain.ListItems.Clear
        RecordList.lvwMain.MultiSelect = False
        RecordList.lvwMain.SmallIcons = Main.ImageList1
        RecordList.lvwMain.Visible = False
            
        ' add column headers: main.lblfield().caption
        RecordList.lvwMain.ColumnHeaders.add , , "Nr", 800, lvwColumnLeft
        For N = 1 To 17
            If Len(Trim$(lblField(N).Caption)) > 0 Then
                RecordList.lvwMain.ColumnHeaders.add , , lblField(N).Caption, 1500, lvwColumnLeft
            End If
        Next N
                
        ' set View property to Report.
        RecordList.lvwMain.View = lvwReport
                
        ' add list items from url_curr() or url_impt() array
        For N = 1 To UBound(sr, 2)
            Set itmX = RecordList.lvwMain.ListItems.add()
            
            If Record_GetNotes(sr(1, N), False, True) Then
                itmX.Text = Format$(sr(1, N), "0000+")
            Else
                itmX.Text = Format$(sr(1, N), "0000")
            End If
            
            itmX.SmallIcon = 3
            
            rowCount = 0
            For P = 1 To 17
                If Len(Trim$(lblField(P).Caption)) > 0 Then
                    rowCount = rowCount + 1
                    itmX.SubItems(rowCount) = nr(sr(1, N)).txtField(P)
                End If
            Next P
        Next N
    End If
            
' GREEN, search matches
    Call SetListViewLedger(RecordList.lvwMain, 3)
    RecordList.StatusBar1.Panels.Item(3).Text = UBound(nr) & " Hitlist for current search."
    RecordList.Caption = " Hitlist for current search."
    RecordList.FileItem(0).Visible = True   ' export to Excel
    RecordList.FileItem(1).Visible = True   ' sep
    RecordList.FileItem(2).Visible = False  ' add Records to project
    RecordList.FileItem(3).Visible = False  ' sep
    RecordList.FileItem(4).Visible = True   ' exit
    RecordList.lvwMain.Visible = True
    RecordList.Refresh
    
    ' form behaviour: SHOW SEARCH RESULT LIST
    RecordList.Visible = True
    If Not MasterUser_G Then
        Record_ShowSingleR (CurrRecord_G)
        Record.Visible = True
    Else
        Main.Visible = True
    End If
    
errhandler:
    Exit Sub
End Sub




Private Sub ClearSearch()

    ' clear existing search results
    ReDim sr(1 To 2, 1 To 1)
    lblNumMatches.Caption = vbNullString
    txtSearch(0).Text = vbNullString
    txtSearch(1).Text = vbNullString
    Unload RecordList
            
End Sub

Private Sub chkKeyWord_Click(Index As Integer)
    
On Error GoTo errhandler

Dim N                       As Long

    Select Case Index
        Case 1      ' brændekunder
            If chkKeyWord(1).Value = 1 Then
                For N = 18 To 22
                    lblField(N).Enabled = True
                    txtField(N).Locked = False
                    txtField(N).BackColor = &HFFFFFF
                Next N
                lblCaption(5).Enabled = True
            Else
                For N = 18 To 22
                    lblField(N).Enabled = False
                    txtField(N).Locked = True
                    txtField(N).BackColor = &HEEE8E6
                Next N
                lblCaption(5).Enabled = False
            End If
            
        Case 2      ' pilefletkunder
            If chkKeyWord(2).Value = 1 Then
                For N = 23 To 26
                    lblField(N).Enabled = True
                    txtField(N).Locked = False
                    txtField(N).BackColor = &HFFFFFF
                Next N
                lblCaption(6).Enabled = True
            Else
                For N = 23 To 26
                    lblField(N).Enabled = False
                    txtField(N).Locked = True
                    txtField(N).BackColor = &HEEE8E6
                Next N
                lblCaption(6).Enabled = False
            End If
    End Select
        
errhandler:
    Exit Sub
End Sub

Public Sub cmdAction_Click(Index As Integer)
    
On Error GoTo errhandler

Dim N                       As Long
    
    Select Case Index

'RECORD - DELETE
        Case 1
            If Not MasterUser_G Then Exit Sub
            Call Record_DeleteSingleA(CurrRecord_G)
            StatusBar1.Panels.Item(3).Text = " Record number " & CurrRecord_G & " successfully deleted."
            
'RECORD - STORE
        Case 2
            If Not MasterUser_G Then Exit Sub
            Call Record_StoreSingleA(CurrRecord_G)
            
            ' save database locally
            FileItem_Click 1
            StatusBar1.Panels.Item(3).Text = " Record number " & CurrRecord_G & " saved on local disk."
            
            ' upload database to website
            FileItem_Click 11
                        
'RECORD - NEW
        Case 3
            If Not MasterUser_G Then Exit Sub
            Call Record_StoreSingleA(CurrRecord_G)
            
            ' if last record is empty then do not create another new empty
            If Record_IsEmptyA(NumRecords_G) Then
                StatusBar1.Panels.Item(3).Text = " A new empty record is already created as Record " & NumRecords_G
                CurrRecord_G = NumRecords_G
                nr(CurrRecord_G).ID = GetUniqueID
                Call Record_ShowSingleA(CurrRecord_G)
                Exit Sub
            End If
            
            ' increment array by one
            NumRecords_G = NumRecords_G + 1
            ReDim Preserve nr(1 To NumRecords_G) As RECORD_DATA
            CurrRecord_G = NumRecords_G
            
            ' create new unique ID value
            nr(CurrRecord_G).ID = GetUniqueID
                        
            Call Record_ShowSingleA(CurrRecord_G)
            
            ' set default keyword: last keyword field, index = 10
            chkKeyWord(KeyWordLabelIndex_G - 27).Value = 1
            MsgBox "The Default Keyword is set to: '" & lblField(KeyWordLabelIndex_G).Caption & " '" & Space$(5) & vbCrLf & vbCrLf & "Note that at least one Keyword must be checked.     ", vbInformation + vbOKOnly, " WARNING"
            
            StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
            StatusBar1.Panels.Item(3).Text = " New blank record successfully created as number " & CurrRecord_G & "  -  Default Keyword is: " & lblField(KeyWordLabelIndex_G).Caption
                                    
'LIST COMPOSE
        Case 5                                                              ' SHOW LIST COMPOSE
            ' store current record
            FromSearch_G = False
            FromIncomplete_G = False
            Call Record_StoreSingleA(CurrRecord_G)
            Main.Visible = False
            ComposeList.Show
            If ComposeList.WindowState = vbMinimized Then
                ComposeList.WindowState = vbNormal
            End If
            
' HIDE MAIN FORM
        Case 9
            If MasterUser_G Then
                Call Record_StoreSingleA(CurrRecord_G)
                Call ToolsItem_Click(0)
            End If
            
            For N = Forms.count - 1 To 0 Step -1
                If Forms(N).Name <> "Main" Then
                    Unload Forms(N)
                    WaitABit 0.1
                End If
            Next N
            Main.Visible = False
            
    End Select

errhandler:
    Exit Sub
End Sub


Private Sub cmdAction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
    
    Select Case Index
        Case 6                                                              ' bookmark
            If Shift Then
                SaveSetting "SDU_UK", "User", "Bookmark", CurrRecord_G
            Else
                ' retrieve bookmarked record
                Call Record_StoreSingleA(CurrRecord_G)
                CurrRecord_G = GetSetting("SDU_UK", "User", "Bookmark", "1")
                Main.StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
                Call Record_ShowSingleA(CurrRecord_G)
            End If
            
        Case 7                                                              ' previous / first
            If InfoBox.Visible Then Unload InfoBox
                                    
            If Button = 2 Then
                Call Record_StoreSingleA(CurrRecord_G)
                DoEvents
                CurrRecord_G = 1
                Call Record_ShowSingleA(CurrRecord_G)
                StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
                Call Record_GetNotes(CurrRecord_G, False, True)
            Else
                Call Record_StoreSingleA(CurrRecord_G)
                DoEvents
                CurrRecord_G = CurrRecord_G - 1
                If CurrRecord_G < 1 Then CurrRecord_G = 1
                Call Record_ShowSingleA(CurrRecord_G)
                StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
                Call Record_GetNotes(CurrRecord_G, False, True)
            End If
            
        Case 8                                                              ' next / last
            If InfoBox.Visible Then Unload InfoBox
                        
            If Button = 2 Then
                Call Record_StoreSingleA(CurrRecord_G)
                DoEvents
                CurrRecord_G = NumRecords_G
                Call Record_ShowSingleA(CurrRecord_G)
                StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
                Call Record_GetNotes(CurrRecord_G, False, True)
            Else
                Call Record_StoreSingleA(CurrRecord_G)
                DoEvents
                CurrRecord_G = CurrRecord_G + 1
                If CurrRecord_G > NumRecords_G Then CurrRecord_G = NumRecords_G
                Call Record_ShowSingleA(CurrRecord_G)
                StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
                Call Record_GetNotes(CurrRecord_G, False, True)
            End If
    
    End Select
    
errhandler:
    Exit Sub
End Sub





Private Sub cmdList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler
    
    If Len(txtSearch(0).Text & txtSearch(1).Text) = 0 Then
        MsgBox "There are no search results to display.     ", vbInformation, " NO SEARCH HITS"
        Exit Sub
        
    Else
        lblMatches.Caption = "Antal hits:"
        lblNumMatches.Caption = UBound(sr, 2)
        
        FromSearch_G = True
        FromIncomplete_G = False
        
        If Button = 2 Then              ' show all
            Call BuildListA(True)
        Else
            Call BuildListA            ' show limited
        End If
        
        If Len(txtSearch(0).Text) > 0 Then
            RecordList.StatusBar1.Panels.Item(3).Text = UBound(sr, 2) & " results from search with the query string: " & Main.txtSearch(0).Text
            RecordList.Caption = " Records    (Query = " & txtSearch(0).Text & ")    "
        Else
            RecordList.StatusBar1.Panels.Item(3).Text = UBound(sr, 2) & " results from search with the query string: " & Main.txtSearch(1).Text
            RecordList.Caption = " Records    (Query = " & txtSearch(1).Text & ")    "
        End If
    End If
    
errhandler:
    Exit Sub
End Sub


Private Sub cmdNavSearch_Click(Index As Integer)
    
On Error GoTo errhandler

Dim navRecord               As Long
Dim navField                As Long
Static LastFieldNumber      As Long
    
    lblMatches.Caption = "Record no:"
    
    If sr(1, 1) = 0 Then Exit Sub
    
    If InfoBox.Visible Then Unload InfoBox
    
    Select Case Index
        Case 0                                  ' previous
            If navIndex_M > 1 Then
                navIndex_M = navIndex_M - 1
            Else
                navIndex_M = 1
            End If
        Case 1                                  ' next
            If navIndex_M < UBound(sr, 2) Then
                navIndex_M = navIndex_M + 1
            Else
                navIndex_M = UBound(sr, 2)
            End If
    End Select
    
    lblNumMatches.Caption = navIndex_M
    
    navRecord = sr(1, navIndex_M)
    navField = sr(2, navIndex_M)
    
    CurrRecord_G = navRecord
        
    Call Record_GetNotes(CurrRecord_G, False, True)
                
    Record_ShowSingleA (navRecord)
    
    lblUniqueID.Caption = nr(navRecord).ID
    
    StatusBar1.Panels(1).Text = " Record = " & navRecord
    
    If navField > 0 Then
        LastFieldNumber = navField
        txtField.Item(navField).BackColor = &HC1FFFE
        DoEvents
    Else
        txtField.Item(LastFieldNumber).BackColor = &HFFF9DB
        DoEvents
    End If
    
errhandler:
    Exit Sub
End Sub



Private Sub cmdNotes_Click()

On Error GoTo errhandler
    
    If InfoBox.Visible Then
        Unload InfoBox
    Else
        Call Record_GetNotes(CurrRecord_G, True, False)
    End If
    
errhandler:
    Exit Sub
End Sub


Public Sub DataBaseItem_Click(Index As Integer)
    
On Error GoTo errhandler

Dim Response                As Long
Dim ReturnValue             As Long
Dim msg                     As String
Dim databases               As DATABASES_INFO
    
    If Index = 0 Then GoTo errhandler
    
    'Call Installed_Databases
    
    ' get folder name for database to be loaded
    MAIN_DIR_G = app_LastOpen(3, Index)
        
'--------------------------------------------------------------------------------------------------
    ' NB! Note that ALL up-/download information is cleared at this point.
    ' Hence it is NOT possible to download as the URL's are not available
'--------------------------------------------------------------------------------------------------
        
    ' if it is not possible to retrieve upload information - or internet connection down
    If Get_Upload_Information(MAIN_DIR_G, True) = False Then
        
        msg = "Unable to retrieve up- and download information!     " & vbCrLf & vbCrLf & _
              "Loading the local copy of the database instead...     "
              
        Response = MsgBox(msg, vbDefaultButton1 + vbInformation + vbOKOnly, " UPLOAD INFORMATION MISSING!")
        ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)                                           ' Load database from local folder
        Call WebEditLock_OFF(MAIN_DIR_G)
        
    Else
        Response = MsgBox("Do you wish to download the database from the website ?     ", _
               vbDefaultButton1 + vbQuestion + vbYesNo, _
               " DOWNLOAD DATABASE?")
               
        If Response = vbNo Then
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)                                       ' Load database from local folder
            Call WebEditLock_OFF(MAIN_DIR_G)
            
        ElseIf Response = vbYes Then
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, True)                                        ' Load database from website
            Call WebEditLock_ON(MAIN_DIR_G)
            
        End If
    
    End If
                
    Select Case ReturnValue
        Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded - unidentified error"
        Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
        Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed - database loaded from local folder  " & MAIN_DIR_G
        Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
        Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
        Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
    End Select
        
    If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
    Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
    lblApplicationTitle(0).Caption = captions(45)
    lblApplicationTitle(1).Caption = captions(46)
    Caption = captions(45) & " - " & captions(46)
        
    ' lock new database
    Call WebEditLock_ON(MAIN_DIR_G)
    
    StatusBar1.Panels(2).Text = MAIN_DIR_G
        
errhandler:
    Exit Sub
End Sub

Public Sub FileItem_Click(Index As Integer)

On Error GoTo errhandler

Dim tmpDir                  As String
Dim Success                 As Boolean
Dim tmpResponse             As Long
    
    Select Case Index
    
'OPEN DATABASE
        Case 0
        
            ' get new database folder and subfolder names
            tmpDir = GetDirectoryDialog(Main)
            If InStr(tmpDir, "SDU_") = 0 Then
                Main.StatusBar1.Panels.Item(3).Text = " This is not a valid database folder."
                Exit Sub
            Else
                tmpDir = QualifyPath(tmpDir)
            End If
            MAIN_DIR_G = tmpDir
            
            ' create Main and sub folders on c-drive
            Call CreateSystemFoldersA(MAIN_DIR_G)
            
            ' load selected database
            Success = Compressed_Database_Read(MAIN_DIR_G)
            
            ' set image, caption and database title and subtitle
            If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
            Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
            lblApplicationTitle(0).Caption = captions(45)
            lblApplicationTitle(1).Caption = captions(46)
            Caption = captions(45) & " - " & captions(46)
            
            ' lock new database
            ToolsItem_Click 1
            
            StatusBar1.Panels(2).Text = MAIN_DIR_G
            
            Call Installed_Databases
            
'SAVE DATABASE
        Case 1
            Main.StatusBar1.Panels.Item(3).Text = " Saving backup..."
            Success = Compressed_Database_Write(MAIN_DIR_G)
            Main.StatusBar1.Panels.Item(3).Text = " Database successfully saved."
            
'CREATE NEW DATABASE
        Case 3
            Main.StatusBar1.Panels.Item(3).Text = " Creating new empty database..."
            NewDatabase.SSTab1.Tab = 0
            NewDatabase.Show 1
            Main.StatusBar1.Panels.Item(3).Text = " New empty database successfully created."
            Call Installed_Databases
            
'DATABASE - SAVE BACKUP
        Case 5
            Main.StatusBar1.Panels.Item(3).Text = " Performing backup of all Records,  please wait..."
            Call Database_SaveBackupA(MAIN_DIR_G)
            Call LogFile_WriteA(3)
            Main.StatusBar1.Panels.Item(3).Text = " Backup of database successfully completed."
            
'DATABASE - LOAD BACKUP
        Case 6
            Main.cdlg.CancelError = True
            Main.cdlg.DialogTitle = "Load Database Backup File"
            Main.cdlg.InitDir = MAIN_DIR_G & "BACKUP"
            Main.cdlg.Filename = "*.*"
            Main.cdlg.Filter = "All files (*.*)|*.*"
            Main.cdlg.flags = &H2 Or &H800
            Main.cdlg.ShowOpen
            
            Main.StatusBar1.Panels.Item(3).Text = " Loading database backup..."
            Call DataBase_LoadBackupA(Main.cdlg.Filename)
                        
'PRIVATE NOTES - SAVE BACKUP
        Case 8
            Main.StatusBar1.Panels.Item(3).Text = " Backing up private Notes..."
            Call Notes_SaveBackupA
            Main.StatusBar1.Panels.Item(3).Text = " Backup of private Notes completed."

'PRIVATE NOTES - LOAD BACKUP
        Case 9
            Main.StatusBar1.Panels.Item(3).Text = " Loading backup of Private Notes..."
            Call Notes_LoadBackupA
            Main.StatusBar1.Panels.Item(3).Text = " Backup of Private Notes successfully completed."
            
'UPLOAD DATABASE (see: http://www.15seconds.com/issue/981203.htm for details)
        Case 11
            Main.StatusBar1.Panels.Item(3).Text = " Uploading Database to website. Please wait..."
            tmpResponse = Compressed_Database_Write(MAIN_DIR_G, True)
            Select Case tmpResponse
                Case 99: Main.StatusBar1.Panels.Item(3).Text = " Upload information is missing. Database was not uploaded."
                Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Uploading Database to website failed."
                Case -1: Main.StatusBar1.Panels.Item(3).Text = " Uploading Database to website successfully completed.": Call LogFile_WriteA(1)
            End Select
                        
'HIDE
        Case 13
            Call cmdAction_Click(9)
            
'EXIT
        Case 15
            Me.Caption = " Small Database Utility is closing..."
            Main.StatusBar1.Panels.Item(3).Text = " Closing Small Database Utility. Please wait..."
            DoEvents
            Unload Me

    End Select
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Load()

On Error Resume Next

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    Me.Height = 8010
    
    Set mydl = New VicsDL
    
    ' get:
    '    total number of databases on c-drive        Databases.TotalNumber
    '    last used database                          Databases.LastUsedDB
    '    foldername of first database on the list    Databases.FirstDB
    databases = Installed_Databases(GetSetting("SDU_UK", "User", "DataPath", "C:\"))
    DoEvents
    
    StatusBar1.Panels.Item(2).Text = MAIN_DIR_G
    
    NumRecords_G = 1
    CurrRecord_G = 1
        
    ' set constrols and menu items
    Call SetActiveControls(2)
            
    ' setup tray application parameters
    With IconData
        .cbSize = Len(IconData)                                 ' Length of the NOTIFYICONDATA type
        .hIcon = Me.Icon                                        ' Reference to Main icon
        .hWnd = Me.hWnd                                         ' hWnd of the Main
        .szTip = "Small Database Utility" & Chr(0)              ' Tooltip message
        .uCallbackMessage = WM_MOUSEMOVE                        ' Target for messages
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE            ' Flags
        .uID = vbNull
    End With
    
    ' set menu captions for Tools menu item
    If GetRunSettingsFromRegistry(0) Then
        ToolsItem(7).Caption = "Run Once"
        lblResident.Caption = "Program is running resident"
    Else
        ToolsItem(7).Caption = "Run Resident"
        lblResident.Caption = "Program is running once"
    End If
    
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler
    
Dim N                       As Long

    ' catch mouse double-click on tray icon
    If x = WM_LBUTTONCLK And y = 0 Then
    
        ' unload all loaded forms - except Main
        If Forms.count > 1 Then
            For N = Forms.count - 1 To 0 Step -1
                If Forms(N).Name <> "Main" And Forms(N).Name <> "Record" Then
                    Unload Forms(N)
                    WaitABit 0.1
                End If
            Next N
        End If
        
        ' Show Main if hidden - else hide Main
        If Main.Visible Then
            Main.Visible = False
        Else
            Main.Visible = True
            Main.WindowState = 0
        End If
    End If

errhandler:
    Exit Sub
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo errhandler

Dim N                       As Long
Dim msg                     As String
Dim Response                As Long
Dim Success                 As Boolean
    
    ' Do not change this caption! - it is used by the updater to monitor the closing down of sdu
    Me.Caption = " Small Database Utility is closing..."
    
    If MasterUser_G Then
        If Configuration_Is_Dirty_G Then
            msg = "Changes to the configuration have not been saved!   " & vbCrLf & vbCrLf & _
                  "Do you want to save them now?"
            Response = MsgBox(msg, vbExclamation + vbYesNo, " WARNING")
        
            If Response = vbYes Then
              AppConfig.cmdAccept_Click
              Call LogFile_WriteA(4)
              DoEvents
            End If
            
        End If
        
        ' make sure that current record and database is saved
        Call Record_StoreSingleA(CurrRecord_G)
        Success = Compressed_Database_Write(MAIN_DIR_G)
        
        ' tell user if local and remote database files are different
        Call Compare_Local_And_Remote_DB(True, False, False)
                
    End If
    
    ' store local  folder path to database and current bookmark
    If Len(MAIN_DIR_G) > 0 Then
        SaveSetting "SDU_UK", "User", "DataPath", MAIN_DIR_G
        SaveSetting "SDU_UK", "User", "Bookmark", CurrRecord_G
    End If
    
    ' clear weblock if the lock is ON
    If lblLocked.Caption = "IS ON" Then
        Call WebEditLock_OFF(MAIN_DIR_G)
    End If
    
    ' remove icons from tray
    Call Shell_NotifyIcon(NIM_DELETE, IconData)
    DoEvents
    
    ' scan forms collection and close all loaded forms - except Main - which is in the process of closing anyway
    If Forms.count > 0 Then
        For N = Forms.count - 1 To 0 Step -1
            If Forms(N).Name <> "Main" Then
            Unload Forms(N)
                WaitABit 1
            End If
        Next N
    End If
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

Dim Success                 As Boolean
    
    ' if user has decided to update...
    If UpdateToNewVersion_G Then
        Success = Shell(App.Path & "\" & "updater.exe", vbNormalFocus)
        DoEvents
    End If
    
    WaitABit 1
    
End Sub

Public Sub HelpItem_Click(Index As Integer)

On Error Resume Next

Dim msg                     As String
Dim Response                As Long
Dim fil                     As Long
Dim tmp                     As String

    Select Case Index
        Case 0
            If FileExist(App.Path & "\hlp\vejledning.rtf") Then
                fil = FreeFile
                Open App.Path & "\hlp\vejledning.rtf" For Input As #fil
                    tmp = Input(LOF(fil), fil)
                Close #fil
                InfoBox.Caption = " Vejledning"
                InfoBox.txtNotes.TextRTF = tmp
                InfoBox.txtNotes.Locked = True
                InfoBox.Left = Screen.Width - InfoBox.Width - 1200
                InfoBox.Show
            End If
            
        Case 2
            ' open html page in Internet Explorer
            Set oIE7 = Nothing
            Set oIE7 = CreateObject("InternetExplorer.application")
            oIE7.Visible = True
            oIE7.navigate "http://www.swr.dk/software/_slideshows/sdu_uk/_jas.htm"
            Set oIE7 = Nothing
            
        Case 4
            ' read log file on update website to get latest update version
            Call GetVersionFromWeb
            
            ' ask user what to do
            If NewestVersionOnWebRaw_G > Val(App.Major & App.Minor & Format$(App.Revision, "000")) Then
                
                msg = "An update to Small Database Utility is available.     " & vbCrLf & vbCrLf & _
                "YES" & Chr(9) & "to update now" & vbCrLf & _
                "NO" & Chr(9) & "to read about the update     " & vbCrLf & _
                "CANCEL" & Chr(9) & "to continue without updating     "
                
                Response = MsgBox(msg, vbInformation + vbYesNoCancel, " NEW VERSION CHECK")
                If Response = vbYes Then                                                            ' UPDATE
                    UpdateToNewVersion_G = True
                    Unload Me
                ElseIf Response = vbNo Then                                                         ' VISIT DOWNLOAD PAGE
                    UpdateToNewVersion_G = False
                    Set oIE7 = Nothing
                    Set oIE7 = CreateObject("InternetExplorer.application")
                    oIE7.Visible = True
                    oIE7.navigate uploaddb.ProgramInfoURL
                    Set oIE7 = Nothing
                Else                                                                                ' DO NOTHING
                    UpdateToNewVersion_G = False
                    Exit Sub
                End If
            Else
                msg = "There are currently no updates or new versions available.     "
                Response = MsgBox(msg, vbInformation + vbOKOnly, " NEW VERSION CHECK")
                UpdateToNewVersion_G = False
                Exit Sub
            End If
            
        Case 6
            ' check connection to webserver and set active controls according to result
            Call WebServerConnectionStatus_Refresh(MAIN_DIR_G, 5)
            If MasterUser_G Then
                Call SetActiveControls(1)
            Else
                Call SetActiveControls(2)
            End If
        Case 8
            About.Show 1
            
    End Select
    
errhandler:
    'oIE7.TheaterMode = False
    'Set oIE7 = Nothing
    Exit Sub
End Sub













Private Sub lblLocked_Click()

On Error GoTo errhandler

    Call ToolsItem_Click(0)
            
errhandler:
    Exit Sub
End Sub

Private Sub lblWebLockStatus_Click()
    
On Error GoTo errhandler
    
    If Main.lblWebLockStatus.Caption = "ON" And lblLocked.Caption = "IS OFF" Then
        Call ToolsItem_Click(1)
        If ToolsItem(1).Caption = "WebLock Is Cleared" Then
            Call ToolsItem_Click(0)
        End If
    End If

errhandler:
    Exit Sub
End Sub

Private Sub ListItem_Click(Index As Integer)

On Error GoTo errhandler

Dim Success                 As Long
Dim sTopic                  As String
Dim hWndDesk                As Long

    Select Case Index
        Case 0
            ComposeList.Show
            
        Case 2
            Main.cdlg.CancelError = True
            Main.cdlg.DialogTitle = "Open Excel File"
            Main.cdlg.InitDir = EXCEL_DIR_G
            Main.cdlg.Filename = "*.xls"
            Main.cdlg.Filter = "All files (*.*)|*.*|Excel file (*.xls)|*.xls"
            Main.cdlg.FilterIndex = 2
            Main.cdlg.flags = &H800 Or &H1000
            Main.cdlg.ShowOpen
            
            hWndDesk = GetDesktopWindow()
            
            Success = ShellExecute(hWndDesk, sTopic, "excel.exe", Main.cdlg.Filename, EXCEL_DIR_G, 1)
            
    End Select
    
errhandler:
    Exit Sub
End Sub





Public Sub m_DataBases_Click()

End Sub

Public Sub ToolsItem_Click(Index As Integer)

On Error GoTo errhandler

Dim msg                     As String
Dim ReturnValue             As Long
Dim CurrentRecord           As Long

    Select Case Index
    
'ENABLE EDIT DATABASE
        Case 0
            If ToolsItem(0).Caption = "Enable Editing Database..." Or lblLocked.Caption = "IS OFF" Then
            
                If lblLocked.Caption = "IS ON" Then
                    GoTo DisableEdit
                    ToolsItem(0).Caption = "Disable Editing Database"
                    Exit Sub
                End If
                
                ' Read locked status: Exit if database is being edited by another user
                If WebEditLock_READ(MAIN_DIR_G) = "ON" Then
                    msg = "The database " & MAIN_DIR_G & " is being edited by another user!     " & vbCrLf & vbCrLf & _
                          "Either wait until the other user completes editing and unlocks the database." & vbCrLf & vbCrLf & _
                          "Or - if you are sure that nobody is editing the database - clear the Web Lock     " & vbCrLf & _
                          "to enable editing the database."
                    MsgBox msg, vbExclamation + vbOKOnly, " WARNING"
                    Exit Sub
                End If
                DoEvents
                
                If Not ToolsItem(1).Caption = "WebLock Is Cleared" Then
                    Passwrd.lblPassWordMsg.Caption = "Enter password to enable Edit:"
                    Passwrd.Show 1
                    DoEvents
                End If
                            
                ' see if password is valid
                If InStr(ValidPasswords_G, Password_G) > 0 And Len(Password_G) >= 3 Then
                    
                    ' store current record number before activating Edit function
                    CurrentRecord = CurrRecord_G
                    
                    ' Lock database so other users cannot edit
                    Call WebEditLock_ON(MAIN_DIR_G)
                    
                    ' get database from website ensuring that editing is performed on the most recent copy
                    ReturnValue = Compressed_Database_Read(MAIN_DIR_G, True)
                    
                    Select Case ReturnValue
                        Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded, unidentified error"
                        Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
                        Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Retrieving the database from  " & uploaddb.WebsiteURL & "  failed, loaded then local copy in  " & MAIN_DIR_G
                        Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from the local folder  " & MAIN_DIR_G
                        Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
                        Case -2: Main.StatusBar1.Panels.Item(3).Text = " Retrieving the database from  " & uploaddb.WebsiteURL & "  failed and a local copy does not exist in  " & MAIN_DIR_G
                        Case Else: Main.StatusBar1.Panels.Item(3).Text = " An unknown error occurred, the database could not be loaded..."
                    End Select
                                                        
                    ' set main caption
                    If InStr(Me.Caption, "(read only)") > 0 Then
                        Me.Caption = Space$(2) & Trim$(Replace(Me.Caption, "(read only)", vbNullString))
                    End If
                            
                    Call SetActiveControls(1)
                                    
                    ' display new record
                    CurrRecord_G = CurrentRecord
                    Call Record_ShowSingleA(CurrRecord_G)
                    Call LogFile_WriteA(0)
                    DoEvents
                    
                    ' clear existing search results
                    Call ClearSearch
                    
                Else
                    If InStr(Me.Caption, "(read only)") = 0 Then
                        Me.Caption = Me.Caption & "     (read only)"
                    End If
                    
                    Call SetActiveControls(2)
                    DoEvents
                    
                End If
                                                
                lblCaption(7).ToolTipText = " Click to set default KeyWord "
                If KeyWordLabelIndex_G = 0 Then KeyWordLabelIndex_G = 29
                lblDefaultKeyword.Caption = "default = " & Trim$(lblField(KeyWordLabelIndex_G))
                
                ToolsItem(0).Caption = "Disable Editing Database"
                
            Else
            
DisableEdit:
'DISABLE EDIT DATABASE

            If lblLocked.Caption = "IS ON" Then
                ToolsItem(0).Caption = "Enable Editing Database... "
            End If
            
            ' Unlock database for editing by other users
            If MasterUser_G Then
                Call WebEditLock_OFF(MAIN_DIR_G)                                   ' Clear Locked flag on website
            End If
                        
            Call SetActiveControls(2)
            DoEvents
            
            lblCaption(7).ToolTipText = vbNullString
            lblDefaultKeyword.Caption = vbNullString
            
            ' clear existing search results
            Call ClearSearch
            
            ' if form Compose list is loaded then unload it, hide form Record and show form Main
            If IsFormLoaded("ComposeList") Then
                Unload ComposeList
            End If
            
            If IsFormLoaded("RecordList") Then
                Unload RecordList
            End If
            
        End If
        
'CLEAR WEB EDIT LOCK
        Case 1
            Passwrd.lblPassWordMsg.Caption = "Enter password to clear Lock and enable Edit:"
            Passwrd.Show 1
            DoEvents
            
            ' see if password is valid
            If InStr(ValidPasswords_G, Password_G) > 0 And Len(Password_G) >= 3 Then
                                
                ' Clear Locked flag on website
                Call WebEditLock_OFF(MAIN_DIR_G)
                StatusBar1.Panels.Item(3).Text = " The database is successfully unlocked and can now be edited."
                
                If ToolsItem(1).Caption = "Clear The Edit WebLock..." Then
                    ToolsItem(1).Caption = "WebLock Is Cleared"
                End If
            End If
                        
'SHOW CONFIGURATION FORM
        Case 3
            Call LogFile_ReadA
            AppConfig.SSTab1.Tab = 11
            AppConfig.Show
            
'ADVANCED TOOLS
        Case 5
            Advanced.Show
            Unload AppConfig
            DoEvents
        
        Case 7
            If ToolsItem(7).Caption = "Run Resident" Then
                ToolsItem(7).Caption = "Run Once"           '
                Call GetRunSettingsFromRegistry(1)                                             ' set Registry RunKey
                lblResident.Caption = "Program is running resident"
            Else
                ToolsItem(7).Caption = "Run Resident"
                Call GetRunSettingsFromRegistry(2)                                             ' delete Registry RunKey
                lblResident.Caption = "Program is running once"
            End If

    End Select
    
errhandler:
    Exit Sub
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
On Error GoTo errhandler
    
    If Panel.Index = 2 Then
        Call ToolsItem_Click(0)
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)

On Error GoTo errhandler
     
    ' when editing is off...
    If lblLocked.Caption = SDU_STATUS_EDIT_OFF Then
        ' offer to enable editing when user start typing
        Select Case KeyAscii
            Case 32 To 127
                If lblLocked.Caption = SDU_STATUS_EDIT_OFF Then
                    KeyAscii = 0
                    Call WebEditLock_OFF(MAIN_DIR_G)
                End If
        End Select
    End If
    
errhandler:
    Exit Sub
End Sub


Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo errhandler
    
    ' when editing is off...
    If lblLocked.Caption = SDU_STATUS_EDIT_OFF Then
        ' offer to enable editing when user start typing
        Select Case KeyAscii
            Case 32 To 127
                If lblLocked.Caption = SDU_STATUS_EDIT_OFF Then
                    KeyAscii = 0
                    Call WebEditLock_OFF(MAIN_DIR_G)
                End If
        End Select
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub txtField_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler

Dim N                       As Long
Dim URL                     As String
Dim EMAIL                   As String
Dim oIE                     As Object
    
    If MasterUser_G Then
        For N = 14 To 17                                                ' MasterUser_G = true
            txtField(N).ToolTipText = " Shift-Click to launch "
        Next N
        If Shift = 0 Then Exit Sub
    Else
        For N = 14 To 17                                                ' MasterUser_G = false
            txtField(N).ToolTipText = " Click to launch "
        Next N
    End If
    
    txtField(Index).BackColor = &HFFF9DB
    
    ' if website URL exists in field 16 and/or 17 then open the homepage
    Select Case Index
        Case 14
            EMAIL = LCase$(Trim$(txtField(14).Text))
        Case 15
            EMAIL = LCase$(Trim$(txtField(15).Text))
        Case 16
            URL = LCase$(Trim$(txtField(16).Text))
        Case 17
            URL = LCase$(Trim$(txtField(17).Text))
    End Select
    
    Select Case Index
    
'E-MAIL
        Case 14, 15
            If Len(EMAIL) = 0 Or InStr(EMAIL, "@") = 0 Or InStr(EMAIL, ".") = 0 Then
                Exit Sub
            Else
                Set oIE = CreateObject("InternetExplorer.application")
                oIE.navigate2 "mailto:" & EMAIL & "?subject=Mail from " & GetUserName
                Set oIE = Nothing
            End If
            
'WEBSITE
        Case 16, 17
            ' if url prefix is missing then add it
            If Left$(URL, 4) = "www." Then URL = "http://" & URL
            
            ' if field does not contain an url then exit
            If Len(URL) = 0 Or InStr(URL, "http") = 0 Then
                Exit Sub
            Else
                Set oIE = CreateObject("InternetExplorer.application")
                oIE.Visible = True
                oIE.navigate URL
                Set oIE = Nothing
            End If
    End Select
            
errhandler:
    Exit Sub
End Sub


Private Sub txtJump_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' store value before change is implemented
    txtJumpValue_M = txtJump.Text
    
End Sub

Private Sub txtJump_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo errhandler
    
    ' exit if text in txtJump is not numeric
    If Not IsNumeric(txtJump.Text) Then
        txtJump.Text = vbNullString
        Exit Sub
    End If
    
    ' exit if value in txtJump is larger that total number of records
    If Val(txtJump.Text) >= NumRecords_G Then
        txtJump.Text = txtJumpValue_M
        Exit Sub
    ElseIf Val(txtJump.Text) < 1 Then
        Exit Sub
    End If
        
    ' save current record before jumping to next
    Call Record_StoreSingleA(CurrRecord_G)
    
    ' set new record number
    CurrRecord_G = CLng(txtJump.Text)
    
    ' display new record
    Call Record_ShowSingleA(CurrRecord_G)
    
    StatusBar1.Panels.Item(1).Text = " Record = " & CurrRecord_G
    
errhandler:
    Exit Sub
End Sub


Private Sub txtJump_LostFocus()
    
    txtJump.Text = vbNullString

End Sub

Private Sub txtSearch_Click(Index As Integer)

On Error GoTo errhandler
    
    lblMatches.Caption = "No. hits:"
    
    If Index = 0 Then
        txtSearch(1).Text = vbNullString
    Else
        txtSearch(0).Text = vbNullString
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If txtSearch(Index).SelLength > 0 Then txtSearch(Index).Text = vbNullString
    
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    
    txtSearch(Index).SelStart = 0
    txtSearch(Index).SelLength = Len(txtSearch(Index).Text)
    txtSearch(Index).SelAlignment = 2
    txtSearch(Index).SelLength = 0
    txtSearch(Index).SelStart = Len(txtSearch(Index).Text)
    
End Sub

Private Sub txtSearch_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
On Error GoTo errhandler

Static LastFieldNumber As Long
    
    If Index = 0 Then
        Call Record_Find(txtSearch(Index).Text, False)
    Else
        Call Record_Find(txtSearch(Index).Text, True)
    End If
    
    lblNumMatches.Caption = UBound(sr, 2)
    
    If sr(1, 1) > 0 Then
        CurrRecord_G = sr(1, 1)
        Record_ShowSingleA (CurrRecord_G)
        lblUniqueID.Caption = nr(CurrRecord_G).ID
        
        StatusBar1.Panels(1).Text = " Record = " & sr(1, 1)
        If sr(2, 1) > 0 Then
            LastFieldNumber = sr(2, 1)
            txtField.Item(sr(2, 1)).BackColor = &HC1FFFE
            DoEvents
        Else
            txtField.Item(LastFieldNumber).BackColor = &HFFF9DB
            DoEvents
        End If
    Else
        txtField.Item(LastFieldNumber).BackColor = &HFFF9DB
        lblNumMatches.Caption = 0
        DoEvents
    End If
    
    navIndex_M = 1
    
errhandler:
    Exit Sub
End Sub

Private Sub VLMenuPlus1_SetMenuItemAttributes(ByVal aMenuItem As VLMnuPlus.CMenuItem)

On Error Resume Next

Dim sCaption                As String
            
    VLMenuPlus1.SetImageList ImageList1
    VLMenuPlus1.HighlightStyle = 1
    VLMenuPlus1.HighlightAppearance = 1
    VLMenuPlus1.BitmapBackground = &HE5E1DC
    
    sCaption = VLMenuPlus1.GetCleanCaption(aMenuItem.Caption)
                 
    Select Case sCaption
        
        Case "RUN RESIDENT"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("RESIDENT_OFF").Picture
            
        Case "RUN ONCE"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("RESIDENT").Picture
            
        Case "HIDE PROGRAM"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("HIDE_FORM").Picture
            
        Case "ABOUT"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("ABOUT").Picture
            
        Case "SCREENSHOTS"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("FORMS").Picture
            
        Case "INSTRUCTIONS"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("HELP").Picture
                   
        Case "OPEN DATABASE"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DATABASEOPEN").Picture
             
        Case "SAVE DATABASE"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("SAVE2").Picture
             
        Case "DATABASE - SAVE BACKUP"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("SAVE2").Picture
             
        Case "DATABASE - LOAD BACKUP"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DATABASEOPEN").Picture
             
        Case "PRIVATE NOTES - SAVE BACKUP"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("SAVE2").Picture
             
        Case "PRIVATE NOTES - LOAD BACKUP"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DATABASEOPEN").Picture
             
        Case "UPLOAD DATABASE TO WEBSITE (FTP)"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DOWNLOADSV").Picture
                                                       
        Case "CREATE NEW DATABASE"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("NEW").Picture
             
        Case "CLEAR THE EDIT WEBLOCK"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("LOCKED").Picture
        
        Case "WEBLOCK IS CLEARED"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("UNLOCKED").Picture
                                    
        Case "ENABLE EDITING DATABASE"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("LOCKED").Picture
             
        Case "DISABLE EDITING DATABASE"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("UNLOCKED").Picture
             
        Case "EDIT CONFIGURATION"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("EDIT").Picture
             
        Case "ADVANCED FUNCTIONS"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("TOOLS").Picture
                         
        Case "EXIT PROGRAM"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("CLOSE").Picture
        
        Case "CHOOSE COLUMNS FOR RECORD LIST"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("SELCOLUMN").Picture
             
        Case "OPEN RECORD LIST IN EXCEL FORMAT", _
             "EXPORT RECORD LIST IN EXCEL FORMAT"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("EXCEL").Picture
             
        Case "ADD RECORDS TO CURRENT RECORD SET"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DOWNLOAD").Picture
        
        Case "CLOSE WINDOW"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("CLOSE1").Picture
            
        Case "NEW VERSION CHECK"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("WEBOPEN").Picture

        Case "WEBSERVER CONNECTION CHECK"
             Set aMenuItem.Picture = ImageList1.ListImages.Item("DOWNLOADSV").Picture
             
        Case Is <> "EXISTING DATABASES:"
            Set aMenuItem.Picture = ImageList1.ListImages.Item("DATABASEOPEN").Picture
            
    End Select
        
    If aMenuItem.IsTopLevel = True Then
        Set aMenuItem.Picture = Nothing
    End If
    
End Sub


