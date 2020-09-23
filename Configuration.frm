VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AppConfig 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Application Configuration"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Configuration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdClearLog 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Clear Log"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Accept"
      Height          =   300
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   " Save and Close "
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
      Height          =   300
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   4560
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4305
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7594
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   3528
      BackColor       =   15657190
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Section 1"
      TabPicture(0)   =   "Configuration.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLegend(1)"
      Tab(0).Control(1)=   "txtFieldCaption(3)"
      Tab(0).Control(2)=   "txtFieldCaption(2)"
      Tab(0).Control(3)=   "txtGroupCaption(1)"
      Tab(0).Control(4)=   "txtFieldCaption(1)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Section 2"
      TabPicture(1)   =   "Configuration.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLegend(2)"
      Tab(1).Control(1)=   "txtFieldCaption(9)"
      Tab(1).Control(2)=   "txtFieldCaption(8)"
      Tab(1).Control(3)=   "txtFieldCaption(7)"
      Tab(1).Control(4)=   "txtFieldCaption(6)"
      Tab(1).Control(5)=   "txtFieldCaption(5)"
      Tab(1).Control(6)=   "txtFieldCaption(4)"
      Tab(1).Control(7)=   "txtGroupCaption(2)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Section 3"
      TabPicture(2)   =   "Configuration.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLegend(3)"
      Tab(2).Control(1)=   "txtFieldCaption(13)"
      Tab(2).Control(2)=   "txtFieldCaption(12)"
      Tab(2).Control(3)=   "txtFieldCaption(11)"
      Tab(2).Control(4)=   "txtFieldCaption(10)"
      Tab(2).Control(5)=   "txtGroupCaption(3)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Section 4"
      TabPicture(3)   =   "Configuration.frx":091E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblLegend(4)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblEmail"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblEmailWork"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblWebsitePrivate"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblWebbrowser"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtFieldCaption(17)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtFieldCaption(16)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtFieldCaption(15)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtFieldCaption(14)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtGroupCaption(4)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Section 5"
      TabPicture(4)   =   "Configuration.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtGroupCaption(5)"
      Tab(4).Control(1)=   "txtFieldCaption(18)"
      Tab(4).Control(2)=   "txtFieldCaption(19)"
      Tab(4).Control(3)=   "txtFieldCaption(20)"
      Tab(4).Control(4)=   "txtFieldCaption(21)"
      Tab(4).Control(5)=   "txtFieldCaption(22)"
      Tab(4).Control(6)=   "lblSection5(4)"
      Tab(4).Control(7)=   "lblSection5(3)"
      Tab(4).Control(8)=   "lblSection5(2)"
      Tab(4).Control(9)=   "lblSection5(1)"
      Tab(4).Control(10)=   "lblSection5(0)"
      Tab(4).Control(11)=   "lblLegend(5)"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "Section 6"
      TabPicture(5)   =   "Configuration.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblLegend(6)"
      Tab(5).Control(1)=   "lblKeyword2(0)"
      Tab(5).Control(2)=   "lblKeyword2(1)"
      Tab(5).Control(3)=   "lblKeyword2(2)"
      Tab(5).Control(4)=   "lblKeyword2(3)"
      Tab(5).Control(5)=   "txtFieldCaption(26)"
      Tab(5).Control(6)=   "txtFieldCaption(25)"
      Tab(5).Control(7)=   "txtFieldCaption(24)"
      Tab(5).Control(8)=   "txtFieldCaption(23)"
      Tab(5).Control(9)=   "txtGroupCaption(6)"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Comments:"
      TabPicture(6)   =   "Configuration.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblLegend(7)"
      Tab(6).Control(1)=   "txtFieldCaption(27)"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Key Words:"
      TabPicture(7)   =   "Configuration.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblLegend(8)"
      Tab(7).Control(1)=   "lblKeyWord(0)"
      Tab(7).Control(2)=   "lblKeyWord(1)"
      Tab(7).Control(3)=   "lblDefaultKeyword"
      Tab(7).Control(4)=   "lblKeywordNum(0)"
      Tab(7).Control(5)=   "lblKeywordNum(1)"
      Tab(7).Control(6)=   "lblKeywordNum(2)"
      Tab(7).Control(7)=   "lblKeywordNum(3)"
      Tab(7).Control(8)=   "lblKeywordNum(4)"
      Tab(7).Control(9)=   "lblKeywordNum(5)"
      Tab(7).Control(10)=   "lblKeywordNum(6)"
      Tab(7).Control(11)=   "lblKeywordNum(7)"
      Tab(7).Control(12)=   "txtFieldCaption(37)"
      Tab(7).Control(13)=   "txtFieldCaption(36)"
      Tab(7).Control(14)=   "txtFieldCaption(35)"
      Tab(7).Control(15)=   "txtFieldCaption(34)"
      Tab(7).Control(16)=   "txtFieldCaption(33)"
      Tab(7).Control(17)=   "txtFieldCaption(32)"
      Tab(7).Control(18)=   "txtFieldCaption(31)"
      Tab(7).Control(19)=   "txtFieldCaption(30)"
      Tab(7).Control(20)=   "txtFieldCaption(29)"
      Tab(7).Control(21)=   "txtFieldCaption(28)"
      Tab(7).Control(22)=   "txtGroupCaption(7)"
      Tab(7).Control(23)=   "txtDefaultKeyword"
      Tab(7).ControlCount=   24
      TabCaption(8)   =   "About Keywords:"
      TabPicture(8)   =   "Configuration.frx":09AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblLegend(0)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Welcome Image:"
      TabPicture(9)   =   "Configuration.frx":09C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Image1"
      Tab(9).Control(1)=   "lblAppTitle(1)"
      Tab(9).Control(2)=   "lblAppTitle(2)"
      Tab(9).Control(3)=   "Label5"
      Tab(9).Control(4)=   "Label6"
      Tab(9).Control(5)=   "Option1(1)"
      Tab(9).Control(6)=   "Option1(2)"
      Tab(9).Control(7)=   "Option1(3)"
      Tab(9).Control(8)=   "Option1(4)"
      Tab(9).Control(9)=   "Option1(5)"
      Tab(9).Control(10)=   "Option1(6)"
      Tab(9).Control(11)=   "Option1(7)"
      Tab(9).Control(12)=   "Option1(8)"
      Tab(9).Control(13)=   "Option1(9)"
      Tab(9).ControlCount=   14
      TabCaption(10)  =   "Application Title:"
      TabPicture(10)  =   "Configuration.frx":09E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "cmdTextTitle(1)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "cmdTextTitle(2)"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).Control(2)=   "cmdTextTitle(4)"
      Tab(10).Control(2).Enabled=   0   'False
      Tab(10).Control(3)=   "cmdTextTitle(3)"
      Tab(10).Control(3).Enabled=   0   'False
      Tab(10).Control(4)=   "txtApplicationTitle(2)"
      Tab(10).Control(4).Enabled=   0   'False
      Tab(10).Control(5)=   "txtApplicationTitle(1)"
      Tab(10).Control(5).Enabled=   0   'False
      Tab(10).Control(6)=   "Label3"
      Tab(10).Control(6).Enabled=   0   'False
      Tab(10).Control(7)=   "Label2"
      Tab(10).Control(7).Enabled=   0   'False
      Tab(10).Control(8)=   "lblLegend(9)"
      Tab(10).Control(8).Enabled=   0   'False
      Tab(10).ControlCount=   9
      TabCaption(11)  =   "View Logfile"
      TabPicture(11)  =   "Configuration.frx":09FE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "lblCurrentUser"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "Label4"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).Control(2)=   "Label7"
      Tab(11).Control(2).Enabled=   0   'False
      Tab(11).Control(3)=   "lstLog"
      Tab(11).Control(3).Enabled=   0   'False
      Tab(11).Control(4)=   "Picture1"
      Tab(11).Control(4).Enabled=   0   'False
      Tab(11).ControlCount=   5
      Begin VB.TextBox txtDefaultKeyword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -70020
         MaxLength       =   2
         TabIndex        =   100
         Top             =   1100
         Width           =   245
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   -69740
         Picture         =   "Configuration.frx":0A1A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   78
         Top             =   1000
         Width           =   510
      End
      Begin VB.ListBox lstLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F6F6&
         ForeColor       =   &H00800000&
         Height          =   1920
         Left            =   -74760
         TabIndex        =   74
         Top             =   2160
         Width           =   5440
      End
      Begin VB.CommandButton cmdTextTitle 
         Caption         =   "Transparent"
         Height          =   300
         Index           =   1
         Left            =   -74760
         TabIndex        =   52
         Top             =   2880
         Width           =   1080
      End
      Begin VB.CommandButton cmdTextTitle 
         Caption         =   "Color"
         Height          =   300
         Index           =   2
         Left            =   -73620
         TabIndex        =   53
         Top             =   2880
         Width           =   720
      End
      Begin VB.CommandButton cmdTextTitle 
         Caption         =   "Color"
         Height          =   300
         Index           =   4
         Left            =   -70500
         TabIndex        =   55
         Top             =   2880
         Width           =   720
      End
      Begin VB.CommandButton cmdTextTitle 
         Caption         =   "Font"
         Height          =   300
         Index           =   3
         Left            =   -71280
         TabIndex        =   54
         Top             =   2880
         Width           =   720
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 9"
         Height          =   210
         Index           =   9
         Left            =   -70900
         TabIndex        =   49
         Top             =   3900
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 8"
         Height          =   210
         Index           =   8
         Left            =   -70900
         TabIndex        =   48
         Top             =   3594
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 7"
         Height          =   210
         Index           =   7
         Left            =   -70900
         TabIndex        =   47
         Top             =   3292
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 6"
         Height          =   210
         Index           =   6
         Left            =   -70900
         TabIndex        =   46
         Top             =   2990
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 5"
         Height          =   210
         Index           =   5
         Left            =   -70900
         TabIndex        =   45
         Top             =   2688
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 4"
         Height          =   210
         Index           =   4
         Left            =   -70900
         TabIndex        =   44
         Top             =   2386
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 3"
         Height          =   210
         Index           =   3
         Left            =   -70900
         TabIndex        =   43
         Top             =   2084
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 2"
         Height          =   210
         Index           =   2
         Left            =   -70900
         TabIndex        =   42
         Top             =   1782
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture 1"
         Height          =   210
         Index           =   1
         Left            =   -70900
         TabIndex        =   41
         Top             =   1480
         Width           =   1600
      End
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   2
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":0F65
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   255
         Index           =   27
         Left            =   -74760
         TabIndex        =   28
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   10217704
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":0FDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   1
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16635590
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1059
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   5
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16635590
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":10D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   12
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16635590
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1153
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16635590
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":11D0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   22
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   10011587
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":124D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   40
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   10469279
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":12CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGroupCaption 
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   29
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16576
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1347
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   3
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":13C4
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   4
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":143B
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   4
         Left            =   -74640
         TabIndex        =   6
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":14B2
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   5
         Left            =   -74640
         TabIndex        =   7
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1529
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   6
         Left            =   -72705
         TabIndex        =   8
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":15A0
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   7
         Left            =   -74640
         TabIndex        =   9
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1617
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   8
         Left            =   -74640
         TabIndex        =   10
         Top             =   2760
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":168E
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   9
         Left            =   -74640
         TabIndex        =   11
         Top             =   3060
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1705
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   10
         Left            =   -74640
         TabIndex        =   13
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":177C
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   11
         Left            =   -74640
         TabIndex        =   14
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":17F3
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   12
         Left            =   -74640
         TabIndex        =   15
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":186A
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   13
         Left            =   -74640
         TabIndex        =   16
         Top             =   2760
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":18E1
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   18
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1958
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":19CF
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   20
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1A46
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   17
         Left            =   360
         TabIndex        =   21
         Top             =   2760
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1ABD
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   18
         Left            =   -74640
         TabIndex        =   23
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1B34
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   19
         Left            =   -74640
         TabIndex        =   24
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1BAB
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   20
         Left            =   -74640
         TabIndex        =   25
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1C22
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   21
         Left            =   -74640
         TabIndex        =   26
         Top             =   2760
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1C99
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   22
         Left            =   -74640
         TabIndex        =   27
         Top             =   3060
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1D10
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   23
         Left            =   -74640
         TabIndex        =   56
         Top             =   1860
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1D87
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   24
         Left            =   -74640
         TabIndex        =   57
         Top             =   2160
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1DFE
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   25
         Left            =   -74640
         TabIndex        =   58
         Top             =   2460
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1E75
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   26
         Left            =   -74640
         TabIndex        =   59
         Top             =   2760
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1EEC
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   28
         Left            =   -74640
         TabIndex        =   30
         Top             =   1860
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1F63
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   29
         Left            =   -74640
         TabIndex        =   31
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":1FDA
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   30
         Left            =   -74640
         TabIndex        =   32
         Top             =   2460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":2051
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   31
         Left            =   -74640
         TabIndex        =   33
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":20C8
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   32
         Left            =   -74640
         TabIndex        =   34
         Top             =   3060
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":213F
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   33
         Left            =   -71800
         TabIndex        =   35
         Top             =   1860
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":21B6
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   34
         Left            =   -71800
         TabIndex        =   36
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":222D
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   35
         Left            =   -71800
         TabIndex        =   37
         Top             =   2460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":22A4
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   36
         Left            =   -71800
         TabIndex        =   38
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":231B
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
      Begin RichTextLib.RichTextBox txtFieldCaption 
         Height          =   195
         Index           =   37
         Left            =   -71800
         TabIndex        =   39
         Top             =   3060
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   344
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":2392
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
      Begin RichTextLib.RichTextBox txtApplicationTitle 
         Height          =   420
         Index           =   2
         Left            =   -74760
         TabIndex        =   51
         Top             =   2040
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   741
         _Version        =   393217
         BackColor       =   -2147483633
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":2409
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtApplicationTitle 
         Height          =   585
         Index           =   1
         Left            =   -74760
         TabIndex        =   50
         Top             =   1380
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1032
         _Version        =   393217
         BackColor       =   -2147483633
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"Configuration.frx":2480
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "10"
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
         Index           =   7
         Left            =   -70020
         TabIndex        =   108
         Top             =   3060
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "9"
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
         Index           =   6
         Left            =   -70020
         TabIndex        =   107
         Top             =   2760
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "8"
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
         Index           =   5
         Left            =   -70020
         TabIndex        =   106
         Top             =   2460
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "7"
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
         Index           =   4
         Left            =   -70020
         TabIndex        =   105
         Top             =   2160
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "6"
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
         Index           =   3
         Left            =   -70020
         TabIndex        =   104
         Top             =   1860
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "5"
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
         Index           =   2
         Left            =   -72860
         TabIndex        =   103
         Top             =   3060
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "4"
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
         Index           =   1
         Left            =   -72860
         TabIndex        =   102
         Top             =   2760
         Width           =   245
      End
      Begin VB.Label lblKeywordNum 
         Alignment       =   2  'Center
         BackColor       =   &H00EEE8E6&
         Caption         =   "3"
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
         Index           =   0
         Left            =   -72860
         TabIndex        =   101
         Top             =   2460
         Width           =   245
      End
      Begin VB.Label lblDefaultKeyword 
         Alignment       =   1  'Right Justify
         Caption         =   "Default keyword number"
         Height          =   195
         Left            =   -72200
         TabIndex        =   99
         Top             =   1100
         Width           =   2100
      End
      Begin VB.Label lblKeyWord 
         BackColor       =   &H009FBF9F&
         Caption         =   " keyword 2"
         Height          =   200
         Index           =   1
         Left            =   -72930
         TabIndex        =   98
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label lblKeyWord 
         BackColor       =   &H0098C3C3&
         Caption         =   " keyword 1"
         Height          =   200
         Index           =   0
         Left            =   -72930
         TabIndex        =   97
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label lblKeyword2 
         Caption         =   "(Enabled when keyword 2 is checked)"
         Height          =   195
         Index           =   3
         Left            =   -72700
         TabIndex        =   96
         Top             =   2760
         Width           =   3300
      End
      Begin VB.Label lblKeyword2 
         Caption         =   "(Enabled when keyword 2 is checked)"
         Height          =   195
         Index           =   2
         Left            =   -72700
         TabIndex        =   95
         Top             =   2460
         Width           =   3300
      End
      Begin VB.Label lblKeyword2 
         Caption         =   "(Enabled when keyword 2 is checked)"
         Height          =   195
         Index           =   1
         Left            =   -72700
         TabIndex        =   94
         Top             =   2160
         Width           =   3300
      End
      Begin VB.Label lblKeyword2 
         Caption         =   "(Enabled when keyword 2 is checked)"
         Height          =   195
         Index           =   0
         Left            =   -72700
         TabIndex        =   93
         Top             =   1860
         Width           =   3300
      End
      Begin VB.Label lblSection5 
         Caption         =   "(Enabled when keyword 1 is checked)"
         Height          =   195
         Index           =   4
         Left            =   -72700
         TabIndex        =   92
         Top             =   3060
         Width           =   3300
      End
      Begin VB.Label lblSection5 
         Caption         =   "(Enabled when keyword 1 is checked)"
         Height          =   195
         Index           =   3
         Left            =   -72700
         TabIndex        =   91
         Top             =   2760
         Width           =   3300
      End
      Begin VB.Label lblSection5 
         Caption         =   "(Enabled when keyword 1 is checked)"
         Height          =   195
         Index           =   2
         Left            =   -72700
         TabIndex        =   90
         Top             =   2460
         Width           =   3300
      End
      Begin VB.Label lblSection5 
         Caption         =   "(Enabled when keyword 1 is checked)"
         Height          =   195
         Index           =   1
         Left            =   -72700
         TabIndex        =   89
         Top             =   2160
         Width           =   3300
      End
      Begin VB.Label lblSection5 
         Caption         =   "(Enabled when keyword 1 is checked)"
         Height          =   195
         Index           =   0
         Left            =   -72700
         TabIndex        =   88
         Top             =   1860
         Width           =   3300
      End
      Begin VB.Label lblWebbrowser 
         Caption         =   "*    (Opens Internet Explorer on click)"
         Height          =   255
         Left            =   2200
         TabIndex        =   87
         Top             =   2760
         Width           =   3300
      End
      Begin VB.Label lblWebsitePrivate 
         Caption         =   "*    (Opens Internet Explorer on click)"
         Height          =   255
         Left            =   2200
         TabIndex        =   86
         Top             =   2460
         Width           =   3300
      End
      Begin VB.Label lblEmailWork 
         Caption         =   "*    (Opens Microsoft mail program on click)"
         Height          =   255
         Left            =   2200
         TabIndex        =   85
         Top             =   2160
         Width           =   3300
      End
      Begin VB.Label lblEmail 
         Caption         =   "*    (Opens Microsoft mail program on click)"
         Height          =   195
         Left            =   2200
         TabIndex        =   84
         Top             =   1860
         Width           =   3300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "User information..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -74820
         TabIndex        =   83
         Top             =   1080
         Width           =   1710
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "img: 360 x 270 px"
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
         Left            =   -70440
         TabIndex        =   82
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Welcome form:"
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
         Left            =   -74820
         TabIndex        =   77
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "User Name:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   76
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   -73800
         TabIndex        =   75
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label lblAppTitle 
         Caption         =   "Application subtitle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Index           =   2
         Left            =   -74780
         TabIndex        =   73
         Top             =   3780
         Width           =   3440
      End
      Begin VB.Label lblAppTitle 
         Caption         =   "Application title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   385
         Index           =   1
         Left            =   -74780
         TabIndex        =   72
         Top             =   1530
         Width           =   3440
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2595
         Index           =   0
         Left            =   -73920
         TabIndex        =   71
         Top             =   1490
         Width           =   4425
      End
      Begin VB.Label Label3 
         Caption         =   "Text:"
         Height          =   195
         Left            =   -71250
         TabIndex        =   70
         Top             =   2600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Background:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   69
         Top             =   2600
         Width           =   975
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   9
         Left            =   -73920
         TabIndex        =   68
         Top             =   3480
         Width           =   4425
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2700
         Left            =   -74850
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   3600
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   6
         Left            =   -74760
         TabIndex        =   67
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   5
         Left            =   -74760
         TabIndex        =   66
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   4
         Left            =   240
         TabIndex        =   65
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   3
         Left            =   -74760
         TabIndex        =   64
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   2
         Left            =   -74760
         TabIndex        =   63
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   1
         Left            =   -74760
         TabIndex        =   62
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   7
         Left            =   -74760
         TabIndex        =   61
         Top             =   3480
         Width           =   5400
      End
      Begin VB.Label lblLegend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   8
         Left            =   -74760
         TabIndex        =   60
         Top             =   3480
         Width           =   5400
      End
   End
End
Attribute VB_Name = "AppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetColor() As Long

On Error GoTo errhandler

    Main.cdlg.CancelError = True
    Main.cdlg.Color = TitleColor_G
    Main.cdlg.flags = cdlCCRGBInit
    Main.cdlg.ShowColor
    GetColor = Main.cdlg.Color
    
errhandler:
    Exit Function
End Function

Private Sub UpdateConfig()

On Error GoTo errhandler

    lblAppTitle(1).Caption = txtApplicationTitle(1).Text
    lblAppTitle(2).Caption = txtApplicationTitle(2).Text
    Main.lblApplicationTitle(0).Caption = txtApplicationTitle(1).Text
    Main.lblApplicationTitle(1).Caption = txtApplicationTitle(2).Text
        
errhandler:
    Exit Sub
End Sub


Public Sub cmdAccept_Click()

On Error GoTo errhandler
        
    Call UpdateCaptionsArray
    Call FieldLabels_CopyFromArrayToMain
    Call Installed_Databases(MAIN_DIR_G)
    Call Installed_Databases
                  
    If Len(txtDefaultKeyword.Text) = 0 Then
        KeyWordLabelIndex_G = 28 + 1                            ' first default keyword is 1
    Else
        KeyWordLabelIndex_G = 28 + Val(txtDefaultKeyword.Text)  ' first keyword text field is 28
    End If
    
    captions(57) = KeyWordLabelIndex_G
    
    Unload Me
    
errhandler:
    Screen.MousePointer = 0
    Exit Sub
End Sub

'------------------------------------------------------------------------------
' Delete logfile, clear list, and sdu() array. 23. september 2007, swr
'------------------------------------------------------------------------------
Private Sub cmdClearLog_Click()

On Error GoTo errhandler

Dim Response                As Long
Dim msg                     As String
    
    msg = "You are about to clear the logfile." & vbCrLf & vbCrLf & _
          "Are you sure know what you are doing?     "
          
    Response = MsgBox(msg, vbCritical + vbOKCancel, " WARNING")
    
    If Response = vbOK Then
        If FileExist(MAIN_DIR_G & "program.log") Then
            FileCopy MAIN_DIR_G & "program.log", MAIN_DIR_G & "program.log_" & "bak"
            Kill MAIN_DIR_G & "program.log"
            lstLog.Clear
            ReDim SDU(0 To 0)
        End If
    End If
    
errhandler:
    Exit Sub
End Sub




Private Sub cmdClose_Click()

On Error GoTo errhandler
    
    Unload Me
    
errhandler:
    Exit Sub
End Sub

'------------------------------------------------------------------------------
' Reloads captions() array with configuration data. Note that values for
' global font settings are entered when editing database title and setting the
' image. Created 23. september 2007, swr
'------------------------------------------------------------------------------
Private Sub UpdateCaptionsArray()

On Error GoTo errhandler

Dim N                       As Long
        
        ' clear captions()
        ReDim captions(1 To 80) As String
        
        ' set the current main folder
        'captions(0) = MAIN_DIR_G                                '  0
        
        ' field captions
        For N = 1 To 37
            captions(N) = Trim$(txtFieldCaption(N).Text)        '  1 - 37
        Next N
        
        ' group captions
        For N = 1 To 7
            captions(N + 37) = Trim$(txtGroupCaption(N).Text)   ' 38 - 44
        Next N
        
        ' titles, image and font properties
        captions(45) = txtApplicationTitle(1).Text              ' 45 - 56
        captions(46) = txtApplicationTitle(2).Text
        captions(47) = ImgIndex_G
        captions(48) = TitleFontName_G
        captions(49) = TitleFontSize_G
        captions(50) = TitleColor_G
        captions(51) = TitleOpaque_G
        captions(52) = TitleBackcolor_G
        captions(53) = TitleBold_G
        captions(54) = TitleItalic_G
        captions(55) = StringEncode(Password_G)
        captions(56) = StringEncode(ValidPasswords_G)
        
        ' not used
        For N = 1 To 13                                         ' 57 - 69, not used
            captions(N + 56) = vbNullString
        Next N
        
errhandler:
    Exit Sub
End Sub

Private Sub CopyFieldLabelsFromArray()

On Error GoTo errhandler
        
Dim N                       As Long
    
    ' Fields 1 to 37
    For N = 1 To 37
        txtFieldCaption(N).Text = captions(N)
    Next N
           
    ' Groups + TabCaptions (1 to 6)
    For N = 1 To 6
        txtGroupCaption(N).Text = captions(N + 37)
        SSTab1.TabCaption(N - 1) = captions(N + 37)
    Next N
        
    ' Comments TabCaption
    SSTab1.TabCaption(6) = captions(27)
    
    ' Keywords TabCaption + GroupCaption
    txtGroupCaption(7).Text = captions(44)
    SSTab1.TabCaption(7) = captions(44)
    
    ' Keywords - set forecolor for GroupCaption to White
    txtGroupCaption(7).SelStart = 0
    txtGroupCaption(7).SelLength = Len(txtGroupCaption(7).Text)
    txtGroupCaption(7).SelColor = &HFFFFFF
    txtGroupCaption(7).SelStart = 0
            
    ' Application title + subtitle
    For N = 1 To 2
        txtApplicationTitle(N).SelStart = 0
        txtApplicationTitle(N).SelLength = Len(txtApplicationTitle(N).Text)
        txtApplicationTitle(N).SelColor = &H800000
        txtApplicationTitle(N).SelStart = 0
    Next N
        
errhandler:
    Exit Sub
End Sub



Private Sub cmdTextTitle_Click(Index As Integer)
    
On Error GoTo errhandler
    
    Select Case Index
    
' FONT TRANSPARENT
        Case 1          ' Background transparent
            txtApplicationTitle(1).BackColor = &H8000000F
            txtApplicationTitle(2).BackColor = &H8000000F
            lblAppTitle(1).BackStyle = 0
            lblAppTitle(2).BackStyle = 0
            TitleOpaque_G = 0
            
' FONT OPAQUE
        Case 2          ' Background color
            TitleBackcolor_G = GetColor
            txtApplicationTitle(1).BackColor = TitleBackcolor_G
            txtApplicationTitle(2).BackColor = TitleBackcolor_G
            lblAppTitle(1).BackStyle = 1
            lblAppTitle(2).BackStyle = 1
            lblAppTitle(1).BackColor = TitleBackcolor_G
            lblAppTitle(2).BackColor = TitleBackcolor_G
            TitleOpaque_G = 1
            
' FONT FAMILY ETC.
        Case 3
            Main.cdlg.CancelError = True
            Main.cdlg.FontName = TitleFontName_G
            Main.cdlg.FontSize = TitleFontSize_G
            Main.cdlg.FontBold = TitleBold_G
            Main.cdlg.FontBold = TitleItalic_G
            Main.cdlg.flags = cdlCFScreenFonts Or cdlCFTTOnly
            Main.cdlg.ShowFont
            
            ' store values in global variables
            TitleFontName_G = Main.cdlg.FontName
            TitleFontSize_G = Main.cdlg.FontSize
            TitleBold_G = Main.cdlg.FontBold
            TitleItalic_G = Main.cdlg.FontItalic
            
            txtApplicationTitle(1).Font.Name = TitleFontName_G
            txtApplicationTitle(2).Font.Name = TitleFontName_G
            
            txtApplicationTitle(1).Font.Size = TitleFontSize_G
            txtApplicationTitle(2).Font.Size = TitleFontSize_G * 0.7
            
            txtApplicationTitle(1).Font.Bold = TitleBold_G
            txtApplicationTitle(2).Font.Bold = TitleBold_G
            
            txtApplicationTitle(1).Font.Italic = TitleItalic_G
            txtApplicationTitle(2).Font.Italic = TitleItalic_G
            
            lblAppTitle(1).FontName = TitleFontName_G
            lblAppTitle(2).FontName = TitleFontName_G
            
            lblAppTitle(1).FontSize = TitleFontSize_G * 0.7
            lblAppTitle(2).FontSize = TitleFontSize_G * 0.5
            
            lblAppTitle(1).FontBold = TitleBold_G
            lblAppTitle(2).FontBold = TitleBold_G
            
            lblAppTitle(1).FontItalic = TitleItalic_G
            lblAppTitle(2).FontItalic = TitleItalic_G
            
' FONT COLOR
         Case 4
            TitleColor_G = GetColor
            
            txtApplicationTitle(1).SelStart = 0
            txtApplicationTitle(1).SelLength = Len(txtApplicationTitle(1).Text)
            txtApplicationTitle(1).SelColor = TitleColor_G
            txtApplicationTitle(1).SelStart = 0
            txtApplicationTitle(2).SelStart = 0
            txtApplicationTitle(2).SelLength = Len(txtApplicationTitle(2).Text)
            txtApplicationTitle(2).SelColor = TitleColor_G
            txtApplicationTitle(2).SelStart = 0
            
            lblAppTitle(1).ForeColor = TitleColor_G
            lblAppTitle(2).ForeColor = TitleColor_G
    End Select
        
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler

Dim msg As String

Const wrd1 = "Enter Title and Field Captions for "
Const wrd2 = vbCrLf & "Note that the visible text corresponds to the available space for Field Captions in the Main form."
    
    form_StayOnTop AppConfig, True, "C"
    
    SSTab1.Tab = 0
            
    Call CopyFieldLabelsFromArray
        
    msg = "IMPORTANT information about Keywords:" & vbCrLf & vbCrLf & _
          "Each Record must have at least one of its 10 Keywords checked to make the Record accessible for searching and listing." & vbCrLf & vbCrLf & _
          "Additional Keywords can be checked to include the Record in more than one data category." & vbCrLf & vbCrLf & _
          "The Keyword Caption can be chosen  - and later changed - without affecting the existing classification of the Record."
    
    lblLegend(0).Caption = msg
    lblLegend(1).Caption = wrd1 & " - " & SSTab1.TabCaption(0) & " - data group." & wrd2    ' Group 1
    lblLegend(2).Caption = wrd1 & " - " & SSTab1.TabCaption(1) & " - data group." & wrd2    ' Group 2
    lblLegend(3).Caption = wrd1 & " - " & SSTab1.TabCaption(2) & " - data group." & wrd2    ' Group 3
    lblLegend(4).Caption = wrd1 & " - " & SSTab1.TabCaption(3) & " - data group." & wrd2    ' Group 4
    lblLegend(5).Caption = wrd1 & " - " & SSTab1.TabCaption(4) & " - data group." & wrd2    ' Group 5
    lblLegend(6).Caption = wrd1 & " - " & SSTab1.TabCaption(5) & " - data group." & wrd2    ' Group 6
    lblLegend(7).Caption = "Enter Title for the Comments Field on the Main form."           ' Comments
    lblLegend(8).Caption = "Enter Captions for the 10 Keywords." & wrd2                     ' Keywords
    lblLegend(9).Caption = "Enter application Title and Subtitle." & wrd2                   ' Application title
    
    Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
    
    txtApplicationTitle(1).Text = captions(45)
    txtApplicationTitle(2).Text = captions(46)
    
    txtApplicationTitle(1).Font.Name = TitleFontName_G
    txtApplicationTitle(2).Font.Name = TitleFontName_G
    
    txtApplicationTitle(1).Font.Size = TitleFontSize_G          ' fontsize for title on welcome form
    txtApplicationTitle(2).Font.Size = TitleFontSize_G * 0.5    ' fontsize for subtitle on welcome form
    
    If TitleOpaque_G Then
        txtApplicationTitle(1).BackColor = TitleBackcolor_G
        txtApplicationTitle(2).BackColor = TitleBackcolor_G
    Else
        txtApplicationTitle(1).BackColor = &H8000000F
        txtApplicationTitle(2).BackColor = &H8000000F
    End If
    
    ' paint text
    txtApplicationTitle(1).SelStart = 0
    txtApplicationTitle(1).SelLength = Len(txtApplicationTitle(1).Text)
    txtApplicationTitle(1).SelColor = TitleColor_G
    txtApplicationTitle(1).SelStart = 0
    txtApplicationTitle(2).SelStart = 0
    txtApplicationTitle(2).SelLength = Len(txtApplicationTitle(2).Text)
    txtApplicationTitle(2).SelColor = TitleColor_G
    txtApplicationTitle(2).SelStart = 0
    
    lblAppTitle(1).Font.Name = TitleFontName_G
    lblAppTitle(2).Font.Name = TitleFontName_G
    
    lblAppTitle(1).FontSize = TitleFontSize_G * 0.7             ' fontsize on config copy - title
    lblAppTitle(2).FontSize = lblAppTitle(1).FontSize * 0.6     ' fontsize on config copy - subtitle
    
    lblAppTitle(1).ForeColor = TitleColor_G
    lblAppTitle(2).ForeColor = TitleColor_G
    
    lblAppTitle(1).Caption = captions(45)
    lblAppTitle(2).Caption = captions(46)
    
    lblAppTitle(1).BackStyle = TitleOpaque_G
    lblAppTitle(2).BackStyle = TitleOpaque_G
    lblAppTitle(1).BackColor = TitleBackcolor_G
    lblAppTitle(2).BackColor = TitleBackcolor_G
       
    lblCurrentUser.Caption = GetUserName
    
    Option1(Val(ImgIndex_G)).Value = True
    
    Call LogFile_ReadA
    
    txtDefaultKeyword.Text = KeyWordLabelIndex_G - 28
    txtFieldCaption(KeyWordLabelIndex_G - 1).BackColor = &HC0FFFF
    
    Configuration_Is_Dirty_G = False
    
errhandler:
    Exit Sub
End Sub


Private Sub lblKeyWord_Click(Index As Integer)
    
    Select Case Index
        Case 0
            AppConfig.SSTab1.Tab = 4
        Case 1
            AppConfig.SSTab1.Tab = 5
    End Select
    
End Sub

Private Sub Option1_Click(Index As Integer)
    
    ImgIndex_G = Format$(Index, "00")
    
    Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
    Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
            
End Sub

Private Sub txtApplicationTitle_Change(Index As Integer)

On Error GoTo errhandler
   
    Configuration_Is_Dirty_G = True
    Call UpdateConfig
    
errhandler:
    Exit Sub
End Sub

Private Sub txtDefaultKeyword_Change()
    
Dim N                       As Long

    If Val(txtDefaultKeyword.Text) > 10 Then txtDefaultKeyword.Text = 10
    If Val(txtDefaultKeyword.Text) < 1 Then txtDefaultKeyword.Text = vbNullString
    
    For N = 28 To 37
        txtFieldCaption(N).BackColor = &HFFFFFF
    Next N
    
    KeyWordLabelIndex_G = Val(txtDefaultKeyword.Text) + 28
    txtFieldCaption(KeyWordLabelIndex_G - 1).BackColor = &HC0FFFF
    
    Configuration_Is_Dirty_G = True
    
End Sub

Private Sub txtGroupCaption_Change(Index As Integer)
        
    Configuration_Is_Dirty_G = True

End Sub


