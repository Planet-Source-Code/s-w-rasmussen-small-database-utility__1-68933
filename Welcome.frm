VERSION 5.00
Begin VB.Form Welcome 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5385
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
   Icon            =   "Welcome.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblNewUpdate 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Caption         =   " A New Update Is Available "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2600
      TabIndex        =   3
      Top             =   1905
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label lblDatabaseSourceInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2300
      TabIndex        =   2
      Top             =   3740
      Width           =   3000
   End
   Begin VB.Label lblWelcomeTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   2
      Left            =   90
      TabIndex        =   1
      Top             =   3300
      Width           =   5220
   End
   Begin VB.Label lblWelcomeTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5220
   End
   Begin VB.Image Image1 
      Height          =   4050
      Left            =   0
      Picture         =   "Welcome.frx":08CA
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetWelcomeFormParameters()
    
On Error Resume Next
    
    ' if an image is not selected then use the first, 01
    If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
    
    ' set image on welcome form
    Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg")
    
    ' set Welcome parameters
    lblWelcomeTitle(1).Font.Name = TitleFontName_G
    lblWelcomeTitle(2).Font.Name = TitleFontName_G
    lblWelcomeTitle(1).FontSize = TitleFontSize_G
    lblWelcomeTitle(2).FontSize = TitleFontSize_G * 0.667
    lblWelcomeTitle(1).ForeColor = TitleColor_G
    lblWelcomeTitle(2).ForeColor = TitleColor_G
    lblWelcomeTitle(1).BackStyle = TitleOpaque_G
    lblWelcomeTitle(2).BackStyle = TitleOpaque_G
    lblWelcomeTitle(1).BackColor = TitleBackcolor_G
    lblWelcomeTitle(2).BackColor = TitleBackcolor_G
    lblWelcomeTitle(1).Font.Bold = TitleBold_G
    lblWelcomeTitle(2).Font.Bold = TitleBold_G
    lblWelcomeTitle(1).Font.Italic = TitleItalic_G
    lblWelcomeTitle(2).Font.Italic = TitleItalic_G
    lblWelcomeTitle(1).Caption = captions(45)
    lblWelcomeTitle(2).Caption = captions(46)
    
End Sub

Private Sub SetMainFormParameters()
    
On Error Resume Next
    
    ' if an image is not selected then use the first, 01
    If Len(ImgIndex_G) <> 2 Or Not IsNumeric(ImgIndex_G) Then ImgIndex_G = "01"
    
    ' set Main parameters
    Main.Image1.Picture = LoadPicture(IMGS_DIR_G & "img" & ImgIndex_G & ".jpg"): DoEvents
    Main.lblApplicationTitle(0).Caption = captions(45)
    Main.lblApplicationTitle(1).Caption = captions(46)
    Main.lblField(28).ToolTipText = " Check to enable section: " & captions(42) & Space$(1)
    Main.lblField(29).ToolTipText = " Check to enable section: " & captions(43) & Space$(1)
    
End Sub



Private Sub Form_Load()

On Error Resume Next

Dim msg                     As String
Dim Response                As Long
Dim ReturnValue             As Long
Dim LastUsedDB              As String
Dim Resident                As Boolean
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    form_StayOnTop Me, True, "C"
    
    Call WebServerConnectionStatus_Refresh(GetSetting("SDU_UK", "User", "DataPath", "C:\"), 15) ' internet/webserver connection check
    
    Load Main
    Load Record
    Call Database_EraseA    ' clear/reset/initialise database
    
    Me.Visible = False
    Randomize (Time)
    
    LastUsedDB = GetSetting("SDU_UK", "User", "DataPath", "C:\")    ' registry: get last used database folder name
    CurrRecord_G = GetSetting("SDU_UK", "User", "Bookmark", 1)      ' registry: get last used bookmark value
    
'NO DATABASES ON DISK - NEW DATABASE

    If databases.TotalNumber = 0 Then
        msg = "There are no valid databases on your harddisk.     " & vbCrLf & vbCrLf & _
              "Do you want to create a new empty database?     "
        Response = MsgBox(msg, vbQuestion + vbYesNo, " DATABASE NOT FOUND")
        
        If Response = vbYes Then
            NewDatabase.Show 1                                                                          ' set MAIN_DIR_G (new database)
            If NewDatabaseSuccess_G Then
                lblDatabaseSourceInfo.Caption = "loading local copy of database..."
                ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)
                GoTo LoadDataBaseResult
            Else
                GoTo errhandler
            End If
        Else
            GoTo errhandler
        End If
        
'LAST USED DATABASE NOT FOUND

    ' last used db is missing but the first db on the db-list exists
    ElseIf Len(databases.LastUsedDB) = 0 Or Not FileExist(databases.LastUsedDB & "database.zlb") And Len(databases.FirstDB) > 0 Then
        msg = "Cannot find the database you used the last time!     " & vbCrLf & vbCrLf & _
              "Do you wish to open:" & vbCrLf & vbCrLf & _
              Chr(9) & Trim$(databases.FirstDB) & Space$(5) & _
              vbCrLf & vbCrLf & "instead?"
        Response = MsgBox(msg, vbInformation + vbYesNo, " DATABASE NOT FOUND")
        
        If Response = vbYes Then
            MAIN_DIR_G = databases.FirstDB                                                              ' set MAIN_DIR_G (first database on list)
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)
            GoTo LoadDataBaseResult
        Else
            GoTo errhandler
        End If
     
    Else
        If Len(databases.LastUsedDB) > 0 And FileExist(databases.LastUsedDB & "database.zlb") Then      ' set MAIN_DIR_G (last used)
            MAIN_DIR_G = databases.LastUsedDB
        ElseIf Len(databases.FirstDB) > 0 And FileExist(databases.FirstDB & "database.zlb") Then
            MAIN_DIR_G = databases.FirstDB
        Else
            GoTo errhandler
        End If
        
    End If
        
    ' create subfolders in current database main folder if not already present
    Call CreateSystemFoldersA(MAIN_DIR_G)
        
'RUN  RESIDENT
    If GetRunSettingsFromRegistry(0) Then
        If WebServerConnectionOK_G Then                                                                 ' SERVER CONNECTION OK (download from website)
            Call Compare_Local_And_Remote_DB(False, True, False)
            lblDatabaseSourceInfo.Caption = "downloading database from website..."
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, True)
        Else                                                                                            ' NO SERVER CONNECTION (get local copy)
            lblDatabaseSourceInfo.Caption = "downloading database from website..."
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)
        End If
        Resident = True
    Else
'RUN ONCE
        If WebServerConnectionOK_G Then                                                                 ' SERVER CONNECTION OK (ask user: website or local copy?)
            msg = "Do you wish to download the database:     " & vbCrLf & vbCrLf & _
                  Space$(5) & uploaddb.RemoteFileName & Space$(5) & vbCrLf & vbCrLf & _
                  "from the website ?     "
            Response = MsgBox(msg, vbQuestion + vbYesNo, " DATABASE SOURCE")
            
            If Response = vbNo Then
                lblDatabaseSourceInfo.Caption = "loading local copy of database..."
                ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)
            ElseIf Response = vbYes Then
                Call Compare_Local_And_Remote_DB(False, True, False)
                lblDatabaseSourceInfo.Caption = "downloading database from website..."
                ReturnValue = Compressed_Database_Read(MAIN_DIR_G, True)
            End If
        Else                                                                                            ' NO SERVER CONNECTION (get local copy)
            lblDatabaseSourceInfo.Caption = "loading local copy of database..."
            ReturnValue = Compressed_Database_Read(MAIN_DIR_G, False)
        End If
        Call SetWelcomeFormParameters: DoEvents
        Me.Visible = True
        Resident = False
    End If

LoadDataBaseResult:

'INFORM USER
    Select Case ReturnValue
        Case 0:  Main.StatusBar1.Panels.Item(3).Text = " Database was not loaded - unidentified error"
        Case 1:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully downloaded from  " & uploaddb.WebsiteURL
        Case 2:  Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed - database loaded from local folder  " & MAIN_DIR_G
        Case 3:  Main.StatusBar1.Panels.Item(3).Text = " Database successfully loaded from local folder  " & MAIN_DIR_G
        Case -1: Main.StatusBar1.Panels.Item(3).Text = " Local database file does not exist in  " & MAIN_DIR_G
        Case -2: Main.StatusBar1.Panels.Item(3).Text = " Download from  " & uploaddb.WebsiteURL & "  failed and a local database file does not exist in  " & MAIN_DIR_G
    End Select
    Main.StatusBar1.Panels.Item(2).Text = MAIN_DIR_G
    
'UPDATE AVAILABLE CHECK
    If WebServerConnectionOK_G Then
        Call GetVersionFromWeb                                                                          ' new version check (resident and run once)
        If NewestVersionOnWebRaw_G > Val(App.Major & App.Minor & Format$(App.Revision, "000")) Then
            lblNewUpdate.Visible = True
            Resident = False
            Me.Visible = True
            DoEvents
        Else
            lblNewUpdate.Visible = False
        End If
        WaitABit 3
    End If
    
    Configuration_Is_Dirty_G = False            ' initialise global flag for changed configuration
    Call Shell_NotifyIcon(NIM_ADD, IconData)    ' display icon in tray
    
'MAIN VISIBLE
    Call SetMainFormParameters: DoEvents
    If Resident Then
        Main.Visible = False
    Else
        Main.Visible = True
    End If
    WaitABit 0.2
    
    Unload Me
    
    Exit Sub
    
errhandler:
    Screen.MousePointer = 0
    Me.Visible = False
    Unload Main
End Sub



