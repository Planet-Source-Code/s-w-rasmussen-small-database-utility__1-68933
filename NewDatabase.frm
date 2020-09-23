VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form NewDatabase 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Create New Database"
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
   Icon            =   "NewDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4305
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7594
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   15657190
      TabCaption(0)   =   " How to Create a New Database"
      TabPicture(0)   =   "NewDatabase.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCreateNewDatabase"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Settings for the New Database"
      TabPicture(1)   =   "NewDatabase.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblFileName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblPassWord"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblUserName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblSettings(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtSubTitle"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtDatebaseTitle"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtPassword"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtUserName"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPrefix"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtNewDatabase"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Picture1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Templates"
      TabPicture(2)   =   "NewDatabase.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblSettings(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "optDatabaseTemplate(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "optDatabaseTemplate(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "optDatabaseTemplate(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "optDatabaseTemplate(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Picture1(1)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
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
         Index           =   1
         Left            =   180
         Picture         =   "NewDatabase.frx":091E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3610
         Width           =   510
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
         Index           =   0
         Left            =   -74820
         Picture         =   "NewDatabase.frx":0E69
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3610
         Width           =   510
      End
      Begin RichTextLib.RichTextBox txtCreateNewDatabase 
         Height          =   3495
         Left            =   -74820
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   6165
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"NewDatabase.frx":13B4
      End
      Begin VB.OptionButton optDatabaseTemplate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "A. Create new blank database without Field, Section and Keyword names."
         Height          =   400
         Index           =   0
         Left            =   480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Value           =   -1  'True
         Width           =   5200
      End
      Begin VB.OptionButton optDatabaseTemplate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C8D0D4&
         Caption         =   "C. Use Danish default names for Fields, Sections and Keywords."
         Height          =   400
         Index           =   1
         Left            =   480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5200
      End
      Begin VB.OptionButton optDatabaseTemplate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C8D0D4&
         Caption         =   "B. Use the same Field, Section and Keyword names as the currently loaded database."
         Height          =   400
         Index           =   2
         Left            =   480
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1500
         Visible         =   0   'False
         Width           =   5200
      End
      Begin VB.OptionButton optDatabaseTemplate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C8D0D4&
         Caption         =   "D. Use English default names for Fields, Sections and Keywords."
         Height          =   400
         Index           =   3
         Left            =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2820
         Visible         =   0   'False
         Width           =   5200
      End
      Begin VB.TextBox txtNewDatabase 
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
         Left            =   -71460
         TabIndex        =   1
         Top             =   840
         Width           =   2260
      End
      Begin VB.TextBox txtPrefix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EEE8E6&
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
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "SDU_"
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   -72000
         TabIndex        =   2
         Top             =   1335
         Width           =   2800
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
         Left            =   -72000
         TabIndex        =   3
         Top             =   1845
         Width           =   1200
      End
      Begin VB.TextBox txtDatebaseTitle 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   4
         Text            =   "My New Database"
         Top             =   2340
         Width           =   2800
      End
      Begin VB.TextBox txtSubTitle 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   5
         Text            =   "New Blank Database"
         Top             =   2835
         Width           =   2800
      End
      Begin VB.Label lblSettings 
         Caption         =   "Please choose which template you wish to use for your new database"
         ForeColor       =   &H00800000&
         Height          =   460
         Index           =   1
         Left            =   1140
         TabIndex        =   21
         Top             =   3610
         Width           =   4755
      End
      Begin VB.Label lblSettings 
         Caption         =   "Please note that you must enter valid information in all text fields before you can create the new database."
         ForeColor       =   &H00800000&
         Height          =   460
         Index           =   0
         Left            =   -73860
         TabIndex        =   20
         Top             =   3610
         Width           =   4755
      End
      Begin VB.Label lblUserName 
         BackColor       =   &H00C8D0D4&
         Caption         =   "The name of the administrator"
         Height          =   195
         Left            =   -74640
         TabIndex        =   18
         Top             =   1365
         Width           =   2490
      End
      Begin VB.Label lblPassWord 
         BackColor       =   &H00C8D0D4&
         Caption         =   "Password for editing records"
         Height          =   195
         Left            =   -74640
         TabIndex        =   17
         Top             =   1875
         Width           =   2490
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H00C8D0D4&
         Caption         =   "Name of the folder on the C-drive"
         Height          =   195
         Left            =   -74640
         TabIndex        =   16
         Top             =   870
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C8D0D4&
         Caption         =   "Main title of the database"
         Height          =   195
         Left            =   -74640
         TabIndex        =   15
         Top             =   2370
         Width           =   2490
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C8D0D4&
         Caption         =   "Descriptive subtitle"
         Height          =   195
         Left            =   -74640
         TabIndex        =   14
         Top             =   2865
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Accept"
      Height          =   315
      Index           =   0
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "NewDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NEW_DATABASE
    Filename            As String
    AdminName           As String
    AdminPsw            As String
    DatabaseTitle       As String
    SubTitle            As String
End Type
Private NewDB As NEW_DATABASE


Private Function SaveNewDatabase(ByVal DatabaseType As Long) As Boolean

On Error GoTo errhandler
    
Dim N                       As Long
Dim P                       As Long
Dim tmp                     As String
Dim Success                 As Boolean
Dim ReturnValue             As Long
Dim fil                     As Long
        
    Select Case DatabaseType
    
'NEW BLANK
        Case 0
                        
            ' compose and implement new database
            MAIN_DIR_G = "C:\" & NewDB.Filename
            
            ' create main folder for new database
            If Len(Dir(MAIN_DIR_G, vbDirectory)) = 0 Then
                MkDir MAIN_DIR_G
            End If
            MAIN_DIR_G = QualifyPath(MAIN_DIR_G)
            
            fil = FreeFile
            Open MAIN_DIR_G & "database.txt" For Output As #fil
                Print #fil, "[HEADER]" '...............................................[HEADER]
                Print #fil, "Small Database Utility"
                Print #fil, "Database saved on " & Format$(Date, "DD. MMM YYYY") & " - " & Format$(Time, "HH:MM:SS")
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                Print #fil, MAIN_DIR_G
                Print #fil,
                
                Print #fil, "[RECORDS - 1]" '..........................................[RECORDS]
                For N = 1 To 1
                    Print #fil, GetUniqueID                             ' record id             00
                    For P = 1 To 26
                        Print #fil, "text " & Format$(P, "00")          ' all fields            01 - 26
                    Next P
                    
                    Print #fil, 1                                       ' first keyword is checked
                    For P = 1 To 9
                        Print #fil, 0                                   ' all keyword values    27 - 37"
                    Next P
                    
                    Print #fil, ""                                      ' comments              38
                    Print #fil, "-//-"                                  ' end of record
                Next N
                Print #fil,
                
                Print #fil, "[CONFIGURATION]" '........................................[CONFIGURATION]
                
                ' field names
                For N = 1 To 26
                    Print #fil, "field " & Format$(N, "00")
                Next N
                
                ' comments content
                Print #fil, "comments"
                
                'keyword labels
                For N = 1 To 10
                    Print #fil, "keyword " & Format$(N, "00")
                Next N
                
                'section captions
                For N = 1 To 7
                    Print #fil, "Captions " & Format$(N, "00")
                Next N
                
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                Print #fil, "04"
                Print #fil, "Arial"
                Print #fil, "20"
                Print #fil, "16777215"
                Print #fil, "0"
                Print #fil, "8421440"
                Print #fil, "0"
                Print #fil, "0"
                Print #fil, "141F18"
                Print #fil, "1C28EA6E6D13280A5E1E592D181B184C78CD1669434E414669A87374091C189D2CEB31223E061345051A4B79BD72745D26B53D3D4973B4FFFF"
                
                ' blanks up to 80
                For N = 1 To 80 - (27 + 10 + 1 + 7 + 12)
                    Print #fil,
                Next N

                Print #fil,
                Print #fil, "[UPLOAD SETTINGS]" '......................................[UPLOAD SETTINGS]
                Print #fil, "my RemoteServer"
                Print #fil, "my Username"
                Print #fil, "my Password"
                Print #fil, "*.zlb"
                Print #fil, "my Remote Folder Path"
                Print #fil, "database.zlb"
                Print #fil, "my WebSite URL"
                Print #fil, "my Update URL"
                Print #fil, "my Download URL"
                Print #fil,
                
                Print #fil, "[ADMINISTRATOR]" '........................................[ADMINISTRATOR]
                Print #fil, StringEncode("Default Administrator") & "::" & StringEncode("sdu")
                Print #fil, StringEncode(NewDB.AdminName) & "::" & StringEncode(NewDB.AdminPsw)
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil,
                
                Print #fil, "[SELECTION]" '............................................[SELECTION]
                For N = 1 To 37
                    Print #fil, "0;"
                Next N
            Close #fil
                            
            Open MAIN_DIR_G & "database.txt" For Binary As #fil
                tmp = Space(LOF(fil))
                Get #fil, , tmp
            Close fil
            Success = String_Compress_Save(tmp, MAIN_DIR_G & "database.zlb")
                                                    
            Call CreateSystemFoldersA(MAIN_DIR_G)
            
            ' read new empty database
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G)
                            
            Select Case ReturnValue
                Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded, unidentified error"
                Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
                Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed, database loaded from local folder  " & MAIN_DIR_G
                Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
                Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
                Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
            End Select
                            
            Call FieldLabels_CopyFromArrayToMain
            
            If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
            Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
            Main.lblApplicationTitle(0).Caption = NewDB.DatabaseTitle
            Main.lblApplicationTitle(1).Caption = NewDB.SubTitle
            Main.Caption = NewDB.DatabaseTitle & " - " & NewDB.SubTitle
            
            Call Record_ShowSingleA(1)
            
            CurrRecord_G = 1
            NumRecords_G = 1
            
            Me.Hide
            MsgBox "The new database successfully created     ", vbInformation + vbOK, " NEW DATABASE"
            
'DANISH DEFAULT
        Case 1
                        
            ' compose and implement new database
            MAIN_DIR_G = "C:\" & NewDB.Filename
            
            ' create main folder for new database
            If Len(Dir(MAIN_DIR_G, vbDirectory)) = 0 Then
                MkDir MAIN_DIR_G
            End If
            MAIN_DIR_G = QualifyPath(MAIN_DIR_G)
            
            fil = FreeFile
            Open MAIN_DIR_G & "database.txt" For Output As #fil
                Print #fil, "[HEADER]" '...............................................[HEADER]
                Print #fil, "Small Database Utility"
                Print #fil, "Database saved on " & Format$(Date, "DD. MMM YYYY") & " - " & Format$(Time, "HH:MM:SS")
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                Print #fil, MAIN_DIR_G
                Print #fil,
                
                Print #fil, "[RECORDS - 1]" '..........................................[RECORDS]
                For N = 1 To 1
                    Print #fil, GetUniqueID                             ' record id             00
                    For P = 1 To 26
                        Print #fil, "text " & Format$(P, "00")          ' all fields            01 - 26
                    Next P
                    
                    Print #fil, 1                                       ' first keyword is checked
                    For P = 1 To 9
                        Print #fil, 0                                   ' all keyword values    27 - 37"
                    Next P
                    
                    Print #fil, ""                                      ' comments              38
                    Print #fil, "-//-"                                  ' end of record
                Next N
                Print #fil,
                
                Print #fil, "[CONFIGURATION]" '........................................[CONFIGURATION]
                
                Print #fil, "Fornavn(e)"                ' 26 items
                Print #fil, "Efternavn(e)"
                Print #fil, "Titel"
                Print #fil, "Adresse"
                Print #fil, "Husnummer"
                Print #fil, "Etage"
                Print #fil, "By"
                Print #fil, "Postnummer"
                Print #fil, "Postdistrikt"
                Print #fil, "Fastnet (privat)"
                Print #fil, "Fastnet (arbejde)"
                Print #fil, "Mobil (privat)"
                Print #fil, "Mobil (arbejde)"
                Print #fil, "Email (privat)"
                Print #fil, "Email (arbejde)"
                Print #fil, "Website (privat)"
                Print #fil, "Website (arbejde)"
                Print #fil, "Fødselsår"
                Print #fil, "Fødselsdag"
                Print #fil, "Bryllupsdag"
                Print #fil, "Jubilæum"
                Print #fil, "Dødsdag"
                Print #fil, "Hjemmeside"
                Print #fil, "Grafisk arbejde"
                Print #fil, "Fotoarbejde"
                Print #fil, "Andet"
                
                Print #fil, "Bemærkninger"              ' 1 item
                
                Print #fil, "Mærkedage"                 ' 10 items
                Print #fil, "Kunder"
                Print #fil, "WebShop"
                Print #fil, "Familie"
                Print #fil, "Venner"
                Print #fil, "Bekendte"
                Print #fil, "Kolleger"
                Print #fil, "Forretning"
                Print #fil, "Institution"
                Print #fil, "Virksomhed"
                
                Print #fil, "Navn"                      ' 7 items
                Print #fil, "Adresse"
                Print #fil, "Telefon"
                Print #fil, "Elektronisk"
                Print #fil, "Mærkedage"
                Print #fil, "Kunder"
                Print #fil, "Keywords"
                
                Print #fil, NewDB.DatabaseTitle         ' 12 items
                Print #fil, NewDB.SubTitle
                Print #fil, "04"
                Print #fil, "Arial"
                Print #fil, "20"
                Print #fil, "16777215"
                Print #fil, "0"
                Print #fil, "8421440"
                Print #fil, "0"
                Print #fil, "0"
                Print #fil, "141F18"                    ' default password: "sdu"
                Print #fil, "1C28EA6E6D13280A5E1E592D181B184C78CD1669434E414669A87374091C189D2CEB31223E061345051A4B79BD72745D26B53D3D4973B4FFFF"
                
                ' blanks up to 80
                For N = 1 To 80 - (26 + 1 + 10 + 7 + 12)
                    Print #fil,
                Next N

                Print #fil,
                Print #fil, "[UPLOAD SETTINGS]" '......................................[UPLOAD SETTINGS]
                Print #fil, "my RemoteServer"
                Print #fil, "my Username"
                Print #fil, "my Password"
                Print #fil, "*.zlb"
                Print #fil, "my Remote Folder Path"
                Print #fil, "database.zlb"
                Print #fil, "my WebSite URL"
                Print #fil, "my Update URL"
                Print #fil, "my Download URL"
                Print #fil,
                
                Print #fil, "[ADMINISTRATOR]" '........................................[ADMINISTRATOR]
                Print #fil, StringEncode("Default Administrator") & "::" & StringEncode("sdu")
                Print #fil, StringEncode(NewDB.AdminName) & "::" & StringEncode(NewDB.AdminPsw)
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil,
                
                Print #fil, "[SELECTION]" '............................................[SELECTION]
                For N = 1 To 37
                    Print #fil, "0;"
                Next N
            Close #fil
                            
            Open MAIN_DIR_G & "database.txt" For Binary As #fil
                tmp = Space(LOF(fil))
                Get #fil, , tmp
            Close fil
            Success = String_Compress_Save(tmp, MAIN_DIR_G & "database.zlb")
                                                    
            Call CreateSystemFoldersA(MAIN_DIR_G)
            
            ' check upload information
            Success = Get_Upload_Information(MAIN_DIR_G, True)
        
            ' read new empty database
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G)
                            
            Select Case ReturnValue
                Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded, unidentified error"
                Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
                Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed, database loaded from local folder  " & MAIN_DIR_G
                Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
                Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
                Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
            End Select
                            
            Call FieldLabels_CopyFromArrayToMain
            
            If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "02"
            Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
            Main.lblApplicationTitle(0).Caption = NewDB.DatabaseTitle
            Main.lblApplicationTitle(1).Caption = NewDB.SubTitle
            Main.Caption = NewDB.DatabaseTitle & " - " & NewDB.SubTitle
            
            Call Record_ShowSingleA(1)
            
            CurrRecord_G = 1
            NumRecords_G = 1
            
            Me.Hide
            MsgBox "The new database successfully created     ", vbInformation + vbOK, " NEW DATABASE"
            
'COPY OF LOADED
        Case 2
                        
            ' compose and implement new database
            MAIN_DIR_G = "C:\" & NewDB.Filename
            
            ' create main folder for new database
            If Len(Dir(MAIN_DIR_G, vbDirectory)) = 0 Then
                MkDir MAIN_DIR_G
            End If
            MAIN_DIR_G = QualifyPath(MAIN_DIR_G)
                        
            fil = FreeFile
            Open MAIN_DIR_G & "database.txt" For Output As #fil
                Print #fil, "[HEADER]" '...............................................[HEADER]
                Print #fil, "Small Database Utility"
                Print #fil, "Database saved on " & Format$(Date, "DD. MMM YYYY") & " - " & Format$(Time, "HH:MM:SS")
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                Print #fil, MAIN_DIR_G
                Print #fil,
                
                Print #fil, "[RECORDS - 1]" '..........................................[RECORDS]
                For N = 1 To 1
                    Print #fil, GetUniqueID                             ' record id             00
                    For P = 1 To 26
                        Print #fil, "text " & Format$(P, "00")          ' all fields            01 - 26
                    Next P
                    
                    Print #fil, 1                                       ' first keyword is checked
                    For P = 1 To 9
                        Print #fil, 0                                   ' all keyword values    27 - 37"
                    Next P
                    
                    Print #fil, ""                                      ' comments              38
                    Print #fil, "-//-"                                  ' end of record
                Next N
                Print #fil,
                
                Print #fil, "[CONFIGURATION]" '........................................[CONFIGURATION]
                For N = 1 To 44
                    Print #fil, captions(N)
                Next N
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                For N = 47 To 80
                    Print #fil, captions(N)
                Next N
                
                Print #fil,
                Print #fil, "[UPLOAD SETTINGS]" '......................................[UPLOAD SETTINGS]
                Print #fil, "192.168.1.1"
                Print #fil, "my Username"
                Print #fil, "my Password"
                Print #fil, "my_database.zlb"
                Print #fil, "Remote_Folder_Path/"
                Print #fil, "database.zlb"
                Print #fil, "http://my_website.dk"
                Print #fil, "http://my_update.dk"
                Print #fil, "http://my_download.dk"
                Print #fil,
        
                Print #fil, "[ADMINISTRATOR]" '........................................[ADMINISTRATOR]
                Print #fil, StringEncode("Default Administrator") & "::" & StringEncode("sdu")
                Print #fil, StringEncode(NewDB.AdminName) & "::" & StringEncode(NewDB.AdminPsw)
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
        
                Print #fil, "[SELECTION]" '............................................[SELECTION]
                For N = 1 To 37
                    Print #fil, "0;"
                Next N
                Print #fil,
            Close #fil
                        
        Open MAIN_DIR_G & "database.txt" For Binary As #fil
            tmp = Space(LOF(fil))
            Get #fil, , tmp
        Close fil
        Success = String_Compress_Save(tmp, MAIN_DIR_G & "database.zlb")
                                                
        Call CreateSystemFoldersA(MAIN_DIR_G)
        
        ' check upload information
        Success = Get_Upload_Information(MAIN_DIR_G, True)
            
        ' read new empty database
        ReturnValue = Compressed_Database_Read(MAIN_DIR_G)
                        
        Select Case ReturnValue
            Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded, unidentified error"
            Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
            Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed, database loaded from local folder  " & MAIN_DIR_G
            Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
            Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
            Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
        End Select
                        
        Call FieldLabels_CopyFromArrayToMain
        
        If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "03"
        Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
        Main.lblApplicationTitle(0).Caption = NewDB.DatabaseTitle
        Main.lblApplicationTitle(1).Caption = NewDB.SubTitle
        Main.Caption = NewDB.DatabaseTitle & " - " & NewDB.SubTitle
        
        Call Record_ShowSingleA(1)
        
        CurrRecord_G = 1
        NumRecords_G = 1
        
        Me.Hide
        MsgBox "The new database successfully created     ", vbInformation + vbOK, " NEW DATABASE"
        
'ENGLISH DEFAULT
        Case 3
                        
            ' compose and implement new database
            MAIN_DIR_G = "C:\" & NewDB.Filename
            
            ' create main folder for new database
            If Len(Dir(MAIN_DIR_G, vbDirectory)) = 0 Then
                MkDir MAIN_DIR_G
            End If
            MAIN_DIR_G = QualifyPath(MAIN_DIR_G)
            
            fil = FreeFile
            Open MAIN_DIR_G & "database.txt" For Output As #fil
                Print #fil, "[HEADER]" '...............................................[HEADER]
                Print #fil, "Small Database Utility"
                Print #fil, "Database saved on " & Format$(Date, "DD. MMM YYYY") & " - " & Format$(Time, "HH:MM:SS")
                Print #fil, NewDB.DatabaseTitle
                Print #fil, NewDB.SubTitle
                Print #fil, MAIN_DIR_G
                Print #fil,
                
                Print #fil, "[RECORDS - 1]" '..........................................[RECORDS]
                For N = 1 To 1
                    Print #fil, GetUniqueID                             ' record id             00
                    For P = 1 To 26
                        Print #fil, "text " & Format$(P, "00")          ' all fields            01 - 26
                    Next P
                    
                    Print #fil, 1                                       ' first keyword is checked
                    For P = 1 To 9
                        Print #fil, 0                                   ' all keyword values    27 - 37"
                    Next P
                    
                    Print #fil, ""                                      ' comments              38
                    Print #fil, "-//-"                                  ' end of record
                Next N
                Print #fil,
                
                Print #fil, "[CONFIGURATION]" '........................................[CONFIGURATION]
                
                Print #fil, "First name(s)"         ' 26 items
                Print #fil, "Last name(s)"
                Print #fil, "Title"
                Print #fil, "Address"
                Print #fil, "House number"
                Print #fil, "Floor"
                Print #fil, "City"
                Print #fil, "Zip code"
                Print #fil, "Country"
                Print #fil, "Fastnet (private)"
                Print #fil, "Fastnet (work)"
                Print #fil, "Mobil (private)"
                Print #fil, "Mobile (work)"
                Print #fil, "Email (private)"
                Print #fil, "Email (work)"
                Print #fil, "Website (private)"
                Print #fil, "Website (work)"
                Print #fil, "Date of birth"
                Print #fil, "Birthsday"
                Print #fil, "Wedding date"
                Print #fil, "Anniversary"
                Print #fil, "Date of death"
                Print #fil, "Homepage"
                Print #fil, "Graphic work"
                Print #fil, "Photography"
                Print #fil, "More..."
                
                Print #fil, "Comments"              ' 1 item
                
                Print #fil, "Anniversaries"         ' 10 items
                Print #fil, "Customers"
                Print #fil, "WebShop"
                Print #fil, "Family"
                Print #fil, "Friends"
                Print #fil, "Acquaintances"
                Print #fil, "Colleagues"
                Print #fil, "Business"
                Print #fil, "Institution"
                Print #fil, "Company"
                
                Print #fil, "Personal"              ' 7 items
                Print #fil, "Address"
                Print #fil, "Phone"
                Print #fil, "Electronic"
                Print #fil, "Anniversary"
                Print #fil, "Customers"
                Print #fil, "Keywords"
                
                Print #fil, NewDB.DatabaseTitle     ' 12 items
                Print #fil, NewDB.SubTitle
                Print #fil, "04"
                Print #fil, "Arial"
                Print #fil, "20"
                Print #fil, "16777215"
                Print #fil, "0"
                Print #fil, "8421440"
                Print #fil, "0"
                Print #fil, "0"
                Print #fil, "141F18"                    ' default password: "sdu"
                Print #fil, "1C28EA6E6D13280A5E1E592D181B184C78CD1669434E414669A87374091C189D2CEB31223E061345051A4B79BD72745D26B53D3D4973B4FFFF"
                
                ' blanks up to 80
                For N = 1 To 80 - (26 + 1 + 10 + 7 + 12)
                    Print #fil,
                Next N

                Print #fil,
                Print #fil, "[UPLOAD SETTINGS]" '......................................[UPLOAD SETTINGS]
                Print #fil, "my RemoteServer"
                Print #fil, "my Username"
                Print #fil, "my Password"
                Print #fil, "*.zlb"
                Print #fil, "my Remote Folder Path"
                Print #fil, "database.zlb"
                Print #fil, "my WebSite URL"
                Print #fil, "my Update URL"
                Print #fil, "my Download URL"
                Print #fil,
                
                Print #fil, "[ADMINISTRATOR]" '........................................[ADMINISTRATOR]
                Print #fil, StringEncode("Default Administrator") & "::" & StringEncode("sdu")
                Print #fil, StringEncode(NewDB.AdminName) & "::" & StringEncode(NewDB.AdminPsw)
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil, "::"
                Print #fil,
                
                Print #fil, "[SELECTION]" '............................................[SELECTION]
                For N = 1 To 37
                    Print #fil, "0;"
                Next N
            Close #fil
                            
            Open MAIN_DIR_G & "database.txt" For Binary As #fil
                tmp = Space(LOF(fil))
                Get #fil, , tmp
            Close fil
            Success = String_Compress_Save(tmp, MAIN_DIR_G & "database.zlb")
                                                    
            Call CreateSystemFoldersA(MAIN_DIR_G)
            
            ' check upload information
            Success = Get_Upload_Information(MAIN_DIR_G, True)
            
            ' read new empty database
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G)
                            
            Select Case ReturnValue
                Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded, unidentified error"
                Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
                Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed, database loaded from local folder  " & MAIN_DIR_G
                Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
                Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
                Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
            End Select
                            
            Call FieldLabels_CopyFromArrayToMain
            
            If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "02"
            Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
            Main.lblApplicationTitle(0).Caption = NewDB.DatabaseTitle
            Main.lblApplicationTitle(1).Caption = NewDB.SubTitle
            Main.Caption = NewDB.DatabaseTitle & " - " & NewDB.SubTitle
            
            Call Record_ShowSingleA(1)
            
            CurrRecord_G = 1
            NumRecords_G = 1
            
            Me.Hide
            
            MsgBox "The new database successfully created     ", vbInformation + vbOK, " NEW DATABASE"
            
    End Select
    
errhandler:
    Exit Function
End Function


Private Sub cmdAccept_Click(Index As Integer)

On Error GoTo errhandler

Dim msg                     As String
Dim Response                As Long
Dim Databases               As DATABASES_INFO
                
    If Index = 0 Then                               ' Action
                                         
        Databases = Installed_Databases(MAIN_DIR_G)
        
        ' give user a last chance to save current database before creating a new
        If Len(Databases.LastUsedDB) > 0 Then
            msg = "You are about to close the currently loaded database!     " & vbCrLf & vbCrLf & _
                  "Are you sure that all changes are saved?" & vbCrLf & vbCrLf & _
                  "YES" & Chr(9) & "To proceed without saving the current database." & vbCrLf & _
                  "NO" & Chr(9) & "To save the current database before creating the new one.     " & vbCrLf & _
                  "CANCEL" & Chr(9) & "To skip creating the new database."
            Response = MsgBox(msg, vbCritical + vbYesNoCancel, " WARNING")
            
            If Response = vbNo Then
                Call Compressed_Database_Write(MAIN_DIR_G)
            ElseIf Response = vbCancel Then
                GoTo errhandler
            End If
            
        End If
        
        ' test user input before creating the new database
        If Len(txtNewDatabase.Text) = 0 Or _
           Len(txtUserName.Text) = 0 Or _
           Len(txtPassword.Text) = 0 Or _
           Len(txtDatebaseTitle.Text) = 0 Or _
           Len(txtSubTitle.Text) = 0 Then
       
            msg = "One or more of the text fields are empty!     "
            Response = MsgBox(msg, vbInformation + vbOKOnly, " MISSING INFORMATION")
            Exit Sub
        End If
                                                         
        With NewDB ' load parameters for new database
            .Filename = UCase$(Trim$(txtPrefix.Text & txtNewDatabase.Text))
            .AdminName = Trim$(txtUserName.Text)
            .AdminPsw = Trim$(txtPassword.Text)
            .DatabaseTitle = Trim$(txtDatebaseTitle.Text)
            .SubTitle = Trim$(txtSubTitle.Text)
        End With
        
        If optDatabaseTemplate(0).Value Then
            SaveNewDatabase (0)
        ElseIf optDatabaseTemplate(1).Value Then
            SaveNewDatabase (1)
        ElseIf optDatabaseTemplate(2).Value Then
            SaveNewDatabase (2)
        Else
            SaveNewDatabase (3)
        End If
        
    End If
    NewDatabaseSuccess_G = True
    Unload Me
    Exit Sub
    
errhandler:
    NewDatabaseSuccess_G = False
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()

    txtNewDatabase.SetFocus

End Sub

Private Sub Form_Load()

On Error GoTo errhandler
    
Dim msg                     As String
Dim Databases               As DATABASES_INFO
    
    form_StayOnTop NewDatabase, True, "C"
    
    Databases = Installed_Databases(MAIN_DIR_G)
    
    optDatabaseTemplate.Item(0).BackColor = &HC0FFFF
    optDatabaseTemplate.Item(1).BackColor = &HC8D0D4
    optDatabaseTemplate.Item(2).BackColor = &HC8D0D4
    optDatabaseTemplate.Item(3).BackColor = &HC8D0D4
            
    If Len(Databases.LastUsedDB) = 0 Then
        optDatabaseTemplate(0).Visible = True
        optDatabaseTemplate(1).Visible = True
        optDatabaseTemplate(2).Visible = False  ' use currently loaded db as template
        optDatabaseTemplate(3).Visible = True
    Else
        optDatabaseTemplate(0).Visible = True
        optDatabaseTemplate(1).Visible = True
        optDatabaseTemplate(2).Visible = True   ' use currently loaded db as template
        optDatabaseTemplate(3).Visible = True
    End If
    
    ' User info
    msg = "How to create a new database:" & vbCrLf & vbCrLf & _
          "First you have to enter the name of the folder on the c-drive where you wish the database files to be saved. " & _
          "The first part of the folder name and the drive is fixed and cannot be changed." & vbCrLf & vbCrLf & _
          "Then enter your name and an administrator password required for editing records and creating new ones." & vbCrLf & vbCrLf & _
          "Finally you should give your new database a descriptive title and an a short subtitle." & vbCrLf & vbCrLf & _
          "Please note that the new database is automatically loaded into the program replacing the current database. " & _
          "So remember to save / upload the current datadase before you create the new one."
           
    ' format text field for user instructions
    txtCreateNewDatabase.Text = Space$(5) & vbCrLf & msg
    txtCreateNewDatabase.SelStart = 0
    txtCreateNewDatabase.SelLength = Len(txtCreateNewDatabase.Text)
    txtCreateNewDatabase.SelIndent = 100
    txtCreateNewDatabase.SelRightIndent = 100
    txtCreateNewDatabase.SelLength = 0
    
    txtCreateNewDatabase.SelStart = 0
    txtCreateNewDatabase.SelLength = 7
    txtCreateNewDatabase.SelFontSize = 2
    txtCreateNewDatabase.SelLength = 0
    
    txtNewDatabase.Text = vbNullString
        
errhandler:
    Exit Sub
End Sub


    
Private Sub optDatabaseTemplate_Click(Index As Integer)
    
On Error GoTo errhandler

    Select Case Index
        Case 0
            txtSubTitle.Text = "New Blank Database"
            optDatabaseTemplate.Item(0).BackColor = &HC0FFFF
            optDatabaseTemplate.Item(1).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(2).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(3).BackColor = &HC8D0D4
        Case 1
            txtSubTitle.Text = "Danish Default Database"
            optDatabaseTemplate.Item(0).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(1).BackColor = &HC0FFFF
            optDatabaseTemplate.Item(2).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(3).BackColor = &HC8D0D4
        Case 2
            txtSubTitle.Text = "Copy of Loaded Database"
            optDatabaseTemplate.Item(0).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(1).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(2).BackColor = &HC0FFFF
            optDatabaseTemplate.Item(3).BackColor = &HC8D0D4
        Case 3
            txtSubTitle.Text = "English Default Database"
            optDatabaseTemplate.Item(0).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(1).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(2).BackColor = &HC8D0D4
            optDatabaseTemplate.Item(3).BackColor = &HC0FFFF
    End Select
    
errhandler:
End Sub

Private Sub txtNewDatabase_KeyDown(KeyCode As Integer, Shift As Integer)

    'Select Case KeyCode
    '    Case 30 To 39, 65 To 90, 8
    '    Case Else: KeyCode = 0
    'End Select
    
End Sub

Private Sub txtNewDatabase_KeyPress(KeyAscii As Integer)
   
On Error GoTo errhandler

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
        
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 8
        Case Else: KeyAscii = 0
    End Select
    
errhandler:
    Exit Sub
End Sub

Private Sub txtNewDatabase_KeyUp(KeyCode As Integer, Shift As Integer)

    'Select Case KeyCode
    '    Case 30 To 39, 65 To 90, 8
    '    Case Else: KeyCode = 0
    'End Select
    
End Sub

