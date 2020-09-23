VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ConvertDatabase 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Convert Database"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConvertDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   900
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   1
      Left            =   3660
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2800
      Width           =   900
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Convert"
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
      Index           =   0
      Left            =   2700
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2800
      Width           =   900
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "blablabla blablabla blablabla blablabla blablabla blablabla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   2500
      Width           =   4500
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "New Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   2200
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Database Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1900
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Selections Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1600
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrators Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1300
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Settings Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1000
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Records Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   400
      Width           =   3000
   End
   Begin VB.Label lblNewSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Header Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   100
      Width           =   3000
   End
End
Attribute VB_Name = "ConvertDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------------------------------------------------
' Function to convert databases created by older versions of SDU
' Inserts section captions: [HEADER]
'                           [RECORDS - 000]
'                           [CONFIGURATION]
'                           [UPLOAD SETTINGS]
'                           [ADMINISTRATOR]
'                           [SELECTION]
' Created 16. February 2009, swr
'------------------------------------------------------------------------------
Private Sub ConvertOldDatabase()
    
On Error Resume Next
    
Dim N                       As Long
Dim P                       As Long
Dim tmp                     As String
Dim fil                     As Long
Dim OldDatabaseText         As String
Dim NewHeader               As String
Dim NewRecords              As String
Dim NewConfiguration        As String
Dim NewUploadSettings       As String
Dim NewAdministrator        As String
Dim NewSelection            As String
Dim OldDatabaseFolder       As String
Dim x()                     As String
Dim y()                     As String
Dim z()                     As String
Static LastDir              As String
    
    For N = 0 To 8
        lblNewSections(N).Caption = vbNullString
    Next N
    
    If Len(LastDir) = 0 Then LastDir = MAIN_DIR_G
    
'==================================================================== SELECT FILE
    cdlg.CancelError = True
    cdlg.DialogTitle = "Open Database to Convert"
    cdlg.InitDir = LastDir
    cdlg.Filename = "*.*"
    cdlg.Filter = "All files (*.*)|*.*"
    cdlg.flags = &H2 Or &H800
    cdlg.ShowOpen
    
    LastDir = cdlg.Filename
    
'==================================================================== LOAD FILE
    fil = FreeFile
    Open cdlg.Filename For Input As #fil
        OldDatabaseText = Input(LOF(fil), fil)
    Close #fil
    
'==================================================================== HEADER SECTION
    y() = Split(OldDatabaseText, "..")
    z() = Split(y(0), vbCrLf)
    
    NewHeader = "[HEADER]" & vbCrLf & _
                "Small Database Utility (converted)" & vbCrLf & _
                z(1) & vbCrLf & _
                z(2) & vbCrLf & _
                z(3) & vbCrLf & _
                z(4) & vbCrLf
    lblNewSections(0).Caption = "Header Section : done!"
    OldDatabaseFolder = Trim$(UCase$(z(4)))
    
    Erase z()
    tmp = vbNullString
    
'==================================================================== RECORD SECTION
    ' split old database into individual records
    x() = Split(y(1), "-//-")
                                                                   
    NewRecords = "[RECORDS - " & UBound(x) - 1 & "]"
    For N = 0 To UBound(x) - 1
        tmp = vbNullString
        z() = Split(x(N), vbCrLf)
        For P = 1 To UBound(z)
            If InStr(z(P), "\") = 0 And InStr(z(P), ".sdu") = 0 Then
                tmp = tmp & vbCrLf & Trim$(z(P))
            End If
        Next P
        NewRecords = NewRecords & Trim$(tmp) & "-//-"
    Next N
    lblNewSections(1).Caption = "Records Section [1 - " & UBound(x) & "] : done!"
    Erase y()
    Erase z()
    tmp = vbNullString
    
'==================================================================== CONFIGURATION SECTION
    y() = Split(x(UBound(x)), vbCrLf)
    P = UBound(y)
    If P > 82 Then P = 82
    NewConfiguration = "[CONFIGURATION]"
    For N = 3 To P
        NewConfiguration = NewConfiguration & vbCrLf & y(N)
    Next N
    lblNewSections(2).Caption = "Configuration Section : done!"
    
'==================================================================== UPLOAD SECTION
    NewUploadSettings = "[UPLOAD SETTINGS]" & vbCrLf & _
        "my Server" & vbCrLf & _
        "2A113F46532A340A010F" & vbCrLf & _
        "2A113A54452B2D041E0E" & vbCrLf & _
        "remote_file_name.zlb" & vbCrLf & _
        "remote Path" & vbCrLf & _
        "database.zlb" & vbCrLf & _
        "my Website/" & vbCrLf
        lblNewSections(3).Caption = "Upload Settings Section : done!"
        
'==================================================================== ADMINISTRATOR SECTION
   NewAdministrator = "[ADMINISTRATOR]" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F" & vbCrLf & _
        "230D0C5443342E4B2D0E051B252539434A2C395D42::140C1F"
        lblNewSections(4).Caption = "Administrators Section : done!"
   
'==================================================================== SELECTION SECTION
    NewSelection = "[SELECTION]"
        For N = 1 To 37
            NewSelection = NewSelection & vbCrLf & "0;"
        Next N
        lblNewSections(5).Caption = "Selections Section : done!"
        
        lblNewSections(6).Caption = "Old Database Folder : " & OldDatabaseFolder
        lblNewSections(7).Caption = "New Path : "
        
'==================================================================== SAVE CONVERTED DATABASE
    cdlg.CancelError = True
    cdlg.DialogTitle = "Save Converted Database"
    cdlg.InitDir = LastDir
    cdlg.Filename = "database.txt"
    cdlg.Filter = "All files (*.*)|*.*"
    cdlg.flags = &H2 Or &H800
    cdlg.ShowSave
    
    LastDir = cdlg.Filename
    
    fil = FreeFile
    Open LastDir For Output As #fil
        Print #fil, NewHeader
        Print #fil, NewRecords
        Print #fil, NewConfiguration
        Print #fil, NewUploadSettings
        Print #fil, NewAdministrator
        Print #fil, NewSelection
    Close #fil
    
    lblNewSections(8).Caption = LCase$(cdlg.Filename)
    
errhandler:
    Exit Sub
End Sub




Private Sub cmdAction_Click(Index As Integer)
    
On Error GoTo errhandler

    Select Case Index
        Case 0
            Call ConvertOldDatabase
            
        Case 1
            Unload Me
            
    End Select
                                                                       
errhandler:
    Exit Sub
End Sub




Private Sub Form_Load()

On Error GoTo errhandler
    
    form_StayOnTop ConvertDatabase, True, "C"
        
    'lblInstructions.Caption = "Conversion function for database files created in previous versions of Small Database Utility." & vbCrLf & vbCrLf & _
                              "The function only works for database files saved in plain ascii format so please check in Notepad before you attempt to perform the conversion." & vbCrLf & vbCrLf & _
                              "As the function creates default Upload and Administrator information, parametres must be manually edited after opening the converted file in SDU. The default password for editing program settings is sdu."

errhandler:
    Exit Sub
End Sub


