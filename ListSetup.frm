VERSION 5.00
Begin VB.Form ComposeList 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Compose Record List"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8835
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Show All Records"
      Height          =   300
      Index           =   4
      Left            =   7110
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   96
      TabStop         =   0   'False
      ToolTipText     =   " Show all Records, no keywords "
      Top             =   4920
      Width           =   1600
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Save Selection"
      Height          =   300
      Index           =   3
      Left            =   6120
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1320
   End
   Begin VB.ListBox lstdummy 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
      Height          =   300
      Index           =   2
      Left            =   4600
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4920
      Width           =   840
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Show Selected List"
      Height          =   300
      Index           =   1
      Left            =   5480
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1600
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Clear All"
      Height          =   300
      Index           =   0
      Left            =   7500
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1200
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Frivillig hjælper"
      Height          =   220
      Index           =   37
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bidragyder"
      Height          =   220
      Index           =   36
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bemærkninger"
      Height          =   220
      Index           =   27
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   26
      ToolTipText     =   " Checkmark for 'YES' "
      Top             =   3000
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Virksomhed"
      Height          =   220
      Index           =   10
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   9
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   3000
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Height          =   220
      Index           =   26
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   25
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2520
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Height          =   220
      Index           =   25
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   24
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2280
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hegntype, kvantum"
      Height          =   220
      Index           =   24
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   23
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Height          =   220
      Index           =   23
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   22
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1800
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pris ialt Kr.:"
      Height          =   220
      Index           =   22
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   21
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1320
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - næste levering"
      Height          =   220
      Index           =   21
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   20
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1080
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kvantum, stablede m3:"
      Height          =   220
      Index           =   20
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   19
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   840
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Brænde (sort, længde)"
      Height          =   220
      Index           =   19
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   18
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   600
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Dato - sidste levering"
      Height          =   220
      Index           =   18
      Left            =   3180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   17
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   360
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Hjemmeside"
      Height          =   220
      Index           =   17
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   16
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   4920
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "E-mail"
      Height          =   220
      Index           =   16
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   15
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   4680
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, mobil"
      Height          =   220
      Index           =   15
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   14
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   4440
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Telefon, fastnet"
      Height          =   220
      Index           =   14
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   13
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   4200
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Donationer til hjemmet"
      Height          =   220
      Index           =   13
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   12
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   3720
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Branche, Type"
      Height          =   220
      Index           =   11
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   10
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   3240
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Kirke, Sogn"
      Height          =   220
      Index           =   12
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   11
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   3480
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postnummer"
      Height          =   220
      Index           =   9
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   8
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2520
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Postdistrikt"
      Height          =   220
      Index           =   8
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   7
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2280
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Bynavn"
      Height          =   220
      Index           =   7
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   6
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Etage (tv, mf, th)"
      Height          =   220
      Index           =   6
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   5
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1800
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Husnummer"
      Height          =   220
      Index           =   5
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   4
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1560
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Vejnavn"
      Height          =   220
      Index           =   4
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   3
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   1320
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Stilling, Titel"
      Height          =   220
      Index           =   3
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   2
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   840
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Efternavn"
      Height          =   220
      Index           =   2
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   1
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   600
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Fornavn / Kontaktpers."
      Height          =   220
      Index           =   1
      Left            =   180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   0
      ToolTipText     =   " Check to include in Recordlist "
      Top             =   360
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Julehilsen"
      Height          =   220
      Index           =   32
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "SH-Grafik kunder"
      Height          =   220
      Index           =   35
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Brændekunde"
      Height          =   220
      Index           =   28
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   360
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Pilefletkunde"
      Height          =   220
      Index           =   29
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Sommerfest"
      Height          =   220
      Index           =   30
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   840
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Høstfest"
      Height          =   220
      Index           =   31
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Årsberetning"
      Height          =   220
      Index           =   33
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2160
   End
   Begin VB.CheckBox chktxtList 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEE8E6&
      Caption         =   "Forretningsforb."
      Height          =   220
      Index           =   34
      Left            =   6180
      MaskColor       =   &H00EEE8E6&
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2160
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   8510
      TabIndex        =   95
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   5510
      TabIndex        =   94
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   5510
      TabIndex        =   93
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   5510
      TabIndex        =   92
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2510
      TabIndex        =   91
      Top             =   3960
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2510
      TabIndex        =   90
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2510
      TabIndex        =   89
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2510
      TabIndex        =   88
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "37"
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
      Height          =   225
      Index           =   37
      Left            =   8400
      TabIndex        =   87
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "36"
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
      Height          =   225
      Index           =   36
      Left            =   8400
      TabIndex        =   86
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "35"
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
      Height          =   225
      Index           =   35
      Left            =   8400
      TabIndex        =   85
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "34"
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
      Height          =   225
      Index           =   34
      Left            =   8400
      TabIndex        =   84
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "33"
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
      Height          =   225
      Index           =   33
      Left            =   8400
      TabIndex        =   83
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "32"
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
      Height          =   225
      Index           =   32
      Left            =   8400
      TabIndex        =   82
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
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
      Height          =   225
      Index           =   31
      Left            =   8400
      TabIndex        =   81
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30"
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
      Height          =   225
      Index           =   30
      Left            =   8400
      TabIndex        =   80
      Top             =   840
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "29"
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
      Height          =   225
      Index           =   29
      Left            =   8400
      TabIndex        =   79
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004040D0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "28"
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
      Height          =   225
      Index           =   28
      Left            =   8400
      TabIndex        =   78
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009BE8E8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "27"
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
      Height          =   225
      Index           =   27
      Left            =   5400
      TabIndex        =   77
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009FBF9F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "26"
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
      Height          =   225
      Index           =   26
      Left            =   5400
      TabIndex        =   76
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009FBF9F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "25"
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
      Height          =   225
      Index           =   25
      Left            =   5400
      TabIndex        =   75
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009FBF9F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "24"
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
      Height          =   225
      Index           =   24
      Left            =   5400
      TabIndex        =   74
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009FBF9F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "23"
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
      Height          =   225
      Index           =   23
      Left            =   5400
      TabIndex        =   73
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "22"
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
      Height          =   225
      Index           =   22
      Left            =   5400
      TabIndex        =   72
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21"
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
      Height          =   225
      Index           =   21
      Left            =   5400
      TabIndex        =   71
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
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
      Height          =   225
      Index           =   20
      Left            =   5400
      TabIndex        =   70
      Top             =   840
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "19"
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
      Height          =   225
      Index           =   19
      Left            =   5400
      TabIndex        =   69
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098C3C3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "18"
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
      Height          =   225
      Index           =   18
      Left            =   5400
      TabIndex        =   68
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17"
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
      Height          =   225
      Index           =   17
      Left            =   2400
      TabIndex        =   67
      Top             =   4920
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "16"
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
      Height          =   225
      Index           =   16
      Left            =   2400
      TabIndex        =   66
      Top             =   4680
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
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
      Height          =   225
      Index           =   15
      Left            =   2400
      TabIndex        =   65
      Top             =   4440
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
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
      Height          =   225
      Index           =   14
      Left            =   2400
      TabIndex        =   64
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
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
      Height          =   225
      Index           =   13
      Left            =   2400
      TabIndex        =   63
      Top             =   3720
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
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
      Height          =   225
      Index           =   12
      Left            =   2400
      TabIndex        =   62
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
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
      Height          =   225
      Index           =   11
      Left            =   2400
      TabIndex        =   61
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
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
      Height          =   225
      Index           =   10
      Left            =   2400
      TabIndex        =   60
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
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
      Height          =   225
      Index           =   9
      Left            =   2400
      TabIndex        =   59
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
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
      Height          =   225
      Index           =   8
      Left            =   2400
      TabIndex        =   58
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
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
      Height          =   225
      Index           =   7
      Left            =   2400
      TabIndex        =   57
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
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
      Height          =   225
      Index           =   6
      Left            =   2400
      TabIndex        =   56
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
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
      Height          =   225
      Index           =   5
      Left            =   2400
      TabIndex        =   55
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
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
      Height          =   225
      Index           =   4
      Left            =   2400
      TabIndex        =   54
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
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
      Height          =   225
      Index           =   2
      Left            =   2400
      TabIndex        =   53
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
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
      Height          =   225
      Index           =   3
      Left            =   2400
      TabIndex        =   52
      Top             =   840
      Width           =   300
   End
   Begin VB.Label lblColumnNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
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
      Height          =   225
      Index           =   1
      Left            =   2400
      TabIndex        =   51
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblComposeInfo 
      BackColor       =   &H00EEE8E6&
      ForeColor       =   &H00800000&
      Height          =   1635
      Left            =   3135
      TabIndex        =   48
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H009BE8E8&
      Caption         =   " Bemærkninger:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   3135
      TabIndex        =   44
      Top             =   2760
      Width           =   2560
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
      Height          =   225
      Index           =   1
      Left            =   135
      TabIndex        =   43
      Top             =   120
      Width           =   2560
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
      Height          =   225
      Index           =   2
      Left            =   135
      TabIndex        =   42
      Top             =   1080
      Width           =   2560
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
      Height          =   225
      Index           =   3
      Left            =   135
      TabIndex        =   41
      Top             =   2760
      Width           =   2560
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H0098C3C3&
      Caption         =   " Brænde:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   3135
      TabIndex        =   40
      Top             =   120
      Width           =   2560
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H009FBF9F&
      Caption         =   " Pileflet:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   3135
      TabIndex        =   39
      Top             =   1560
      Width           =   2560
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
      Height          =   225
      Index           =   4
      Left            =   135
      TabIndex        =   38
      Top             =   3960
      Width           =   2560
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H000000C0&
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
      Height          =   225
      Index           =   8
      Left            =   6135
      TabIndex        =   37
      Top             =   120
      Width           =   2560
   End
End
Attribute VB_Name = "ComposeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn()                As Long

Private Sub UpdateColumnNumbers()
    
On Error GoTo errhandler

Dim N                       As Long
Dim cNum                    As Long
Dim x()                     As String
    
    lstdummy.Clear
     
    ' create sorted list of colum active indices + and field indices
    For N = 1 To 37
        If Len(lblColumnNumber(N).Caption) > 0 Then
            lstdummy.AddItem Format$(Val(lblColumnNumber(N).Caption), "00") & ";" & N
        End If
    Next N
    
    ' clear all column numbers
    For N = 1 To 37
        lblColumnNumber(N).Caption = vbNullString
    Next N
    
    ' enter sorted, consequtive column numbers
    cNum = 1
    For N = 0 To lstdummy.ListCount - 1
        x() = Split(lstdummy.List(N), ";")
        lblColumnNumber(x(1)).Caption = cNum
        cNum = cNum + 1
    Next N
    
errhandler:
    Exit Sub
End Sub

'------------------------------------------------------------------------------
' Builds and sort 2d array cn() holding column number and record field numbers
'------------------------------------------------------------------------------
Private Function BuildColumnArray(ByVal ShowKeywords As Boolean, ByVal AllRecords As Boolean) As Boolean

On Error GoTo errhandler

Dim cCount                  As Long
Dim N                       As Long
Dim x()                     As String

    ' build d2 array including Record field number
    ' text: 1-26, comments: 27, keywords: 28-38) and table column number (1 to cCount)
    ReDim cn(1 To 2, 1 To 1)
    cCount = 0
    If Not ShowKeywords And AllRecords Then
        For N = 1 To 17
            If Len(Trim$(Main.lblField(N).Caption)) > 0 Then
                cCount = cCount + 1
                ReDim Preserve cn(1 To 2, 1 To cCount)
                cn(1, cCount) = N                                       ' text field number (1 to 26)
                cn(2, cCount) = N                                       ' column number
            End If
        Next N
        
    ElseIf Not ShowKeywords And Not AllRecords Then
        For N = 1 To 17
            If Len(Trim$(Main.lblField(N).Caption)) > 0 Then
                If chktxtList(N).Value = 1 Then
                    cCount = cCount + 1
                    ReDim Preserve cn(1 To 2, 1 To cCount)
                    cn(1, cCount) = N                                   ' text field number (1 to 26)
                    cn(2, cCount) = Val(lblColumnNumber(N).Caption)     ' column number
                End If
            End If
        Next N
        
    Else
        For N = 1 To 37
            If Len(Trim$(Main.lblField(N).Caption)) > 0 Then
                If chktxtList(N).Value = 1 Then
                    cCount = cCount + 1
                    ReDim Preserve cn(1 To 2, 1 To cCount)
                    cn(1, cCount) = N                                   ' text field number (1 to 26)
                    cn(2, cCount) = Val(lblColumnNumber(N).Caption)     ' column number
                End If
            End If
        Next N
    End If
    
    ' sort cn() array according to column number, cCount
    lstdummy.Clear
    For N = 1 To cCount
        lstdummy.AddItem Format$(cn(2, N), "00") & "§" & cn(1, N)
    Next N
    
    If cCount = 0 Then GoTo errhandler
    
    ReDim cn(1 To 2, 1 To cCount)
    For N = 0 To lstdummy.ListCount - 1
        x() = Split(lstdummy.List(N), "§")
         cn(1, N + 1) = x(0)
         cn(2, N + 1) = x(1)
    Next N
    
    BuildColumnArray = True
    
    Exit Function
    
errhandler:
    BuildColumnArray = False
    Exit Function
End Function

Private Sub BuildList(ByVal ShowKeywords As Boolean, ByVal AllRecords As Boolean)

On Error GoTo errhandler

Dim Q                       As Long
Dim P                       As Long
Dim N                       As Long
Dim x                       As Long
Dim itmX                    As Object
Dim sLine()                 As String
        
    ' close personal comments form before building list
    If InfoBox.Visible Then Unload InfoBox
    
    If Not ShowKeywords And AllRecords Then                     ' complete list, no keywords
        Call BuildColumnArray(False, True)
        
    ElseIf Not ShowKeywords And Not AllRecords Then             ' sellist, no keywords
        Call BuildColumnArray(False, False)
        
    Else                                                        ' sellist, selkeywords
        Call BuildColumnArray(True, False)
        
    End If
    
    ' clear lvwmain
    RecordList.lvwMain.ColumnHeaders.Clear
    RecordList.lvwMain.ListItems.Clear
    RecordList.lvwMain.MultiSelect = False
    RecordList.lvwMain.SmallIcons = Main.ImageList1
    RecordList.lvwMain.Visible = True
    
    ' add column headers: main.lblfield().caption
    RecordList.lvwMain.ColumnHeaders.add , , "Nr", 800, lvwColumnLeft
    For N = 1 To UBound(cn, 2)
        If Len(Trim$(Main.lblField(cn(2, N)).Caption)) > 0 Then
            RecordList.lvwMain.ColumnHeaders.add , , Main.lblField(cn(2, N)).Caption, 1500, lvwColumnLeft        ' link name
        End If
    Next N
               
    ' set View property to Report.
    RecordList.lvwMain.View = lvwReport
            
    ReDim sLine(1 To UBound(cn, 2))
    
    ' add list items from url_curr() or url_impt() array
    For N = 1 To UBound(nr)
    
        ' create empty array to hold column content, Ubound(cn,2) = number of columns
        ReDim sLine(1 To UBound(cn, 2))
        
        Select Case AllRecords
            Case True
                For Q = 1 To UBound(cn, 2)
                    If cn(2, Q) = 27 Then
                        If Len(nr(N).Comments) > 1 Then
                            sLine(Q) = "YES"
                        Else
                            sLine(Q) = vbNullString
                        End If
                    Else
                        sLine(Q) = nr(N).txtField(cn(2, Q))
                    End If
                Next Q
            Case False
                If ShowKeywords Then
                    For P = UBound(cn, 2) To 1 Step -1
                        If cn(2, P) > 27 Then
                            If nr(N).chkKeyWord(cn(2, P) - 27) = 1 Then
                                For Q = 1 To UBound(cn, 2)
                                    If cn(2, Q) > 27 Then
                                        If nr(N).chkKeyWord(cn(2, Q) - 27) = 1 Then
                                            sLine(Q) = "YES"
                                        Else
                                            sLine(Q) = vbNullString
                                        End If
                                    ElseIf cn(2, Q) = 27 Then
                                        If Len(nr(N).Comments) > 1 Then
                                            sLine(Q) = "YES"
                                        Else
                                            sLine(Q) = vbNullString
                                        End If
                                    Else
                                        sLine(Q) = nr(N).txtField(cn(2, Q))
                                    End If
                                Next Q
                                Exit For
                            End If
                        End If
                    Next P
                Else
                    For P = UBound(cn, 2) To 1 Step -1
                        For Q = 1 To UBound(cn, 2)
                            sLine(Q) = nr(N).txtField(cn(2, Q))
                        Next Q
                        Exit For
                    Next P
                End If
        End Select
        
        x = 0
        For P = 1 To UBound(sLine)
            x = x + Len(sLine(P))
        Next P
        If x = 0 Then GoTo NextN
        
        Set itmX = RecordList.lvwMain.ListItems.add()
            
        If Record_GetNotes(N, False, True) Then
            itmX.Text = Format$(N, "0000+")
        Else
            itmX.Text = Format$(N, "0000")
        End If
            
        itmX.SmallIcon = 3
        
        For Q = 1 To UBound(sLine)
            itmX.SubItems(Q) = sLine(Q)
        Next Q
            
NextN:
    Next N
    
If AllRecords Then
'GREY, all records in project
    Call SetListViewLedger(RecordList.lvwMain, 5)   ' grey, all records
    RecordList.StatusBar1.Panels.Item(3).Text = RecordList.lvwMain.ListItems.count & " Alle Records i projectet - uden nøgle information."
    RecordList.Caption = " All Records in project"
    RecordList.FileItem(0).Visible = True   ' export to Excel
    RecordList.FileItem(1).Visible = True   ' sep
    RecordList.FileItem(2).Visible = False  ' add Records to project
    RecordList.FileItem(3).Visible = False  ' sep
    RecordList.FileItem(4).Visible = True   ' exit
    RecordList.lvwMain.Visible = True
    RecordList.Refresh
    RecordList.Show
    
    ' manages forms behaviour
    If MasterUser_G Then
        Main.Visible = True
        Record.Visible = False
    Else
        Main.Visible = False
        Record_ShowSingleR (CurrRecord_G)
        Record.Visible = True
    End If
    
Else
'YELLOW, selected fields
    Call SetListViewLedger(RecordList.lvwMain, 4)   ' yellow, selected fields list
    RecordList.StatusBar1.Panels.Item(3).Text = RecordList.lvwMain.ListItems.count & " Records der indeholder den valgte information."
    RecordList.Caption = " Udvalgte Records i projektet."
    RecordList.FileItem(0).Visible = True   ' export to Excel
    RecordList.FileItem(1).Visible = True   ' sep
    RecordList.FileItem(2).Visible = False  ' add Records to project
    RecordList.FileItem(3).Visible = False  ' sep
    RecordList.FileItem(4).Visible = True   ' exit
    RecordList.lvwMain.Visible = True
    RecordList.Refresh
    RecordList.Show
    
    ' manages forms behaviour
    If MasterUser_G Then
        Main.Visible = True
        Record.Visible = False
    Else
        Main.Visible = False
        Record_ShowSingleR (CurrRecord_G)
        Record.Visible = True
    End If
End If
        
errhandler:
    Exit Sub
End Sub

Private Function GetKeywordStatus() As Boolean

On Error GoTo errhandler
    
Dim N                       As Long
Dim msg                     As String
Dim Response                As Long
Dim NoKeywords              As Boolean
Dim NoTextFields            As Boolean
    
    NoKeywords = True
    NoTextFields = True
                              
    For N = 1 To 27
        If chktxtList(N).Value = 1 Then
            NoTextFields = False
            Exit For
        End If
    Next N
    
    For N = 28 To 37
        If chktxtList(N).Value = 1 Then
            NoKeywords = False
            Exit For
        End If
    Next N
    
    AllKeyWordsChecked_G = True
    For N = 28 To 37
        If chktxtList(N).Enabled And chktxtList(N).Value = 0 Then
            AllKeyWordsChecked_G = False
            Exit For
        End If
    Next N
    
    If NoKeywords And NoTextFields Then
        msg = "You havn't checked any Fields or Keywords.     "
        
    ElseIf NoKeywords Then
        msg = "You havn't checked at least one Keyword.     "
        
    ElseIf NoTextFields Then
        msg = "You havn't checked at least one Text field.     "
        
    Else
        GetKeywordStatus = True
        Exit Function
        
    End If
    
    Response = MsgBox(msg, vbInformation + vbOKOnly, " MISSING INFORMATION")
        
errhandler:
    GetKeywordStatus = False
    Exit Function
End Function

Private Function GetNextColumnNumber() As Long

On Error GoTo errhandler
    
Dim N                       As Long
Dim maxNumber               As Long

    For N = 1 To 37
        If Val(lblColumnNumber(N).Caption) > maxNumber Then
            maxNumber = Val(lblColumnNumber(N).Caption)
        End If
    Next N
    
    GetNextColumnNumber = maxNumber + 1
        
    Exit Function
    
errhandler:
    Exit Function
End Function



Private Sub AnalyseRecords()

On Error GoTo errhandler

Dim N                       As Long
Dim P                       As Long
    
    ' set all checkboxe values to 0
    For N = 1 To 37
        chktxtList(N).Value = 0
    Next N
    
    ' disable all checkboxes
    For N = 1 To 37
        chktxtList(N).Caption = Main.lblField(N).Caption
        chktxtList(N).Enabled = False
        lblColumnNumber(N).BackColor = &H8000000F
        lblColumnNumber(N).BorderStyle = 0
    Next N
        
    For N = 1 To 6
        lblCaption(N).Caption = Main.lblCaption(N).Caption
    Next N
    lblCaption(7).Caption = Main.lblField(27).Caption
    lblCaption(8).Caption = Main.lblCaption(7).Caption
    
    ' enable chkmarks if a single record contains information
    For N = 1 To UBound(nr)
        
        ' read text fields
        For P = 1 To 17
            If Len(nr(N).txtField(P)) > 0 Then
                chktxtList(P).Enabled = True
                lblColumnNumber(P).BackColor = &HFDD6C6
                lblColumnNumber(P).BorderStyle = 1
            End If
        Next P
        
        ' read text fields
        For P = 18 To 22
            If Len(nr(N).txtField(P)) > 0 Then
                chktxtList(P).Enabled = True
                lblColumnNumber(P).BackColor = &H98C3C3
                lblColumnNumber(P).BorderStyle = 1
            End If
        Next P
        
        ' read text fields
        For P = 23 To 26
            If Len(nr(N).txtField(P)) > 0 Then
                chktxtList(P).Enabled = True
                lblColumnNumber(P).BackColor = &H9FBF9F
                lblColumnNumber(P).BorderStyle = 1
            End If
        Next P
        
        ' read comment field
        If Len(nr(N).Comments) > 10 Then
            chktxtList(27).Enabled = True
            lblColumnNumber(27).BackColor = &H9BE8E8
            lblColumnNumber(27).BorderStyle = 1
        End If
        
        ' read keywords
        For P = 1 To 10
            If nr(N).chkKeyWord(P) = 1 Then
                chktxtList(P + 27).Enabled = True
                lblColumnNumber(P + 27).BackColor = &H4040D0
                lblColumnNumber(P + 27).BorderStyle = 1
            End If
        Next P
        
    Next N
    
    ' clear all column numbers
    For N = 1 To 37
        lblColumnNumber(N).Caption = vbNullString
    Next N
        
errhandler:
    Exit Sub
End Sub

Public Sub chktxtList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
    
    If chktxtList(Index).Value = 1 Then
        lblColumnNumber(Index).Caption = GetNextColumnNumber
    Else
        lblColumnNumber(Index).Caption = vbNullString
    End If
    
    UpdateColumnNumbers
    
errhandler:
    Exit Sub
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
On Error GoTo errhandler

Dim N                       As Long

    Select Case Index
        Case 0                                  ' Reset settings
            Call AnalyseRecords
                        
        Case 1                                  ' Accept settings and show list
            If Not GetKeywordStatus Then Exit Sub
            FromSearch_G = False
            FromIncomplete_G = False
            
            If AllKeyWordsChecked_G Then
                Call BuildList(False, False)
            Else
                Call BuildList(True, False)
            End If
            
            ' manages forms behaviour
            If MasterUser_G Then
                Main.Visible = True
                Record.Visible = False
                ComposeList.Visible = False
                RecordList.Show
            Else
                Main.Visible = False
                Record_ShowSingleR (CurrRecord_G)
                Record.Visible = True
                ComposeList.Visible = False
            End If
                        
        Case 2                                  ' Close form without saving
            Main.Visible = True
            Unload ComposeList
        
        Case 3
            ReDim selected(1 To 2, 1 To 37)
            For N = 1 To 37
                If chktxtList(N).Enabled Then
                    selected(1, N) = chktxtList(N).Value
                    selected(2, N) = lblColumnNumber(N).Caption
                Else
                    selected(1, N) = 0
                    selected(2, N) = vbNullString
                End If
            Next N
            
            'Call IniFile_Write_Selection        ' Store selection
            
        Case 4
            FromSearch_G = False
            Me.Visible = False
            DoEvents
            Call BuildList(False, True)
            
    End Select
    
errhandler:
    Exit Sub
End Sub
Private Sub Form_Load()

On Error GoTo errhandler

Dim N                       As Long
Dim msg                     As String
    
    form_StayOnTop ComposeList, True, "C"
            
    'Me.Hide
    
    Call AnalyseRecords
    
    msg = "Place the checkmarks in the order you want to display the information. Only fields containing information in at least one Record are enabled. " & _
          "The column number is shown to the right of the checkbox. Note that only Records with at least one Keyword are shown." & vbCrLf & vbCrLf & _
          "Click 'Rebuild' to update the form. When you are done then click 'Show List.'"
        
    lblComposeInfo.Caption = msg
    
    'Call IniFile_Read_Selection
    For N = 1 To 37
        chktxtList(N).Value = swVal(selected(1, N))
        lblColumnNumber(N).Caption = selected(2, N)
    Next N
    
    'Me.Show
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Main.Visible = True
    Record.Visible = False
    
End Sub

Private Sub lblSel_Click(Index As Integer)

On Error GoTo errhandler

Dim N                       As Long
Dim cValue                  As Long
    
    If lblSel(Index).Caption = "-" Then lblSel(Index).Caption = "+": cValue = 0 Else lblSel(Index).Caption = "-": cValue = 1
    
    Select Case Index
        Case 1
            For N = 1 To 3
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 2
            For N = 4 To 9
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 3
            For N = 10 To 13
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 4
            For N = 14 To 17
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 5
            For N = 18 To 22
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 6
            For N = 23 To 26
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 7
            For N = 27 To 27
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
        Case 8                                                      ' keywords
            For N = 28 To 37
                If chktxtList(N).Enabled Then
                    chktxtList(N).Value = cValue
                    If cValue = 0 Then
                        lblColumnNumber(N).Caption = vbNullString
                    Else
                        lblColumnNumber(N).Caption = N
                    End If
                End If
            Next N
                        
    End Select
    
    Call UpdateColumnNumbers
    
errhandler:
    Exit Sub
End Sub
