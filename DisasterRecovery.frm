VERSION 5.00
Begin VB.Form DisasterRecovery 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Disaster Recovery"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DisasterRecovery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   180
      Picture         =   "DisasterRecovery.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   510
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
      Height          =   300
      Left            =   4690
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   900
   End
   Begin VB.CommandButton cmdRecovery 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Load Disaster Recovery File"
      Height          =   300
      Index           =   1
      Left            =   3200
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2400
   End
   Begin VB.CommandButton cmdRecovery 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Create Disaster Recovery File"
      Height          =   300
      Index           =   0
      Left            =   3200
      MaskColor       =   &H00EEE8E6&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   2400
   End
   Begin VB.Label InfoLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ready..."
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
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label Label 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Download and load Disaster Recovery copy of the current database."
      Height          =   435
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   2900
   End
   Begin VB.Label Label 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Create a disaster recovery copy on the website of the current database. "
      Height          =   435
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2900
   End
End
Attribute VB_Name = "DisasterRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' Download disaster recovery copy of the current database.
' Created 8. August 2011, swr
'------------------------------------------------------------------------------
Private Function RecoveryFile_Retrieve(ByVal SDU_FolderName As String) As Boolean

On Error GoTo errhandler

Dim msg                     As String
Dim fil                     As Long
Dim Response                As Long
Dim RecoveryText            As String
Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String
Dim Success                 As Boolean
    
    InfoLine.Caption = vbNullString
        
    msg = "You are about to detete the currently loaded database     " & vbCrLf & _
          "and load the Disaster Recovery copy downloaded from    " & vbCrLf & _
          "the website instead." & vbCrLf & vbCrLf & _
          "You should consider this as the last resort to be used     " & vbCrLf & _
          "only when the current database is seriously corrupted." & vbCrLf & vbCrLf & _
          Space$(5) & "OK - To retrieve and load the recovery copy.     " & vbCrLf & _
          Space$(5) & "CANCEL - To cancel the operation." & vbCrLf & vbCrLf & _
          "Please note that this operation is irreversible."
    Response = MsgBox(msg, vbExclamation + vbOKCancel, " WARNING")
    
    If Response = vbCancel Then
        InfoLine.Caption = "Operation was cancelled by user..."
        GoTo errhandler
    End If
        
    ' check if upload settings are valid
    If Get_Upload_Information(SDU_FolderName) = False Then
        InfoLine.Caption = "Upload settings not available..."
        GoTo errhandler
    
    ' download recovery file from website
    Else
        ' compose recovery file name based on uploaddb.RemoteFileName
        RemoteFileName = Replace(uploaddb.RemoteFileName, "zlb", "disasterrecovery")
        LocalFilePath = SDU_FolderName & Replace(uploaddb.LocalFileName, "zlb", "txt")
        ret = WebsiteTransfer(vbGet, LocalFilePath, RemoteFileName)
        
        ' exit with False if a recovery file is not found in the web folder
        If Not ret Then
            InfoLine.Caption = "Disaster Recovery copy of database does not exist..."
            GoTo errhandler
        End If
        
        ' open (database.txt), compress and save local database file (database.zlb)
        If FileExist(LocalFilePath) Then
            fil = FreeFile
            Open LocalFilePath For Binary As #fil
                RecoveryText = Space(LOF(fil))
                Get #fil, , RecoveryText
            Close fil
            Call String_Compress_Save(RecoveryText, uploaddb.LocalFileName)
        Else
            InfoLine.Caption = "Local database file missing..."
            GoTo errhandler
        End If
            
        ' parse and load recovery text
        Success = Parse_Database_Text(RecoveryText)
        
        If Success Then
            InfoLine.Caption = "Disaster recovery copy of the current database successfully loaded..."
            RecoveryFile_Retrieve = True
        Else
            GoTo errhandler
        End If
        
    End If
    
    Call LogFile_WriteA(2)
    
    Exit Function

errhandler:
    InfoLine.Caption = "Loading the disaster recovery copy of the current database failed..."
    RecoveryFile_Retrieve = False
    Exit Function
End Function
'------------------------------------------------------------------------------
' If the Remote and Local compressed files are identical then upload a text
' version of the the local database (datase.txt) to the website with the
' extension disasterrecovery.
' Return True if successful else False
' Created 8. August 2011, swr
'------------------------------------------------------------------------------
Private Function RecoveryFile_Create(ByVal SDU_FolderName As String) As Boolean

On Error GoTo errhandler

Dim ret                     As Long
Dim LocalFilePath           As String
Dim RemoteFileName          As String

    InfoLine.Caption = vbNullString
    
    ' check if upload settings are valid
    If Get_Upload_Information(SDU_FolderName) = False Then
        InfoLine.Caption = "Upload settings not available..."
        GoTo errhandler
        
    Else
        If Compare_Local_And_Remote_DB(False, False, False) = False Then
            InfoLine.Caption = "The Local and Remote copy of the database are not identical..."
            GoTo errhandler
        End If
        
        LocalFilePath = SDU_FolderName & Replace(uploaddb.LocalFileName, "zlb", "txt")
        RemoteFileName = Replace(uploaddb.RemoteFileName, "zlb", "disasterrecovery")
        ret = WebsiteTransfer(vbPut, LocalFilePath, RemoteFileName)
                
        InfoLine.Caption = "Disaster Recovery copy of the database successfully uploaded..."
        RecoveryFile_Create = True
        
    End If
    
    Call LogFile_WriteA(5)
    Exit Function

errhandler:
    InfoLine.Caption = "Uploading a Disaster Recovery copy of the database failed..."
    RecoveryFile_Create = False
    Exit Function
End Function
Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub
Private Sub cmdRecovery_Click(Index As Integer)
    
On Error GoTo errhandler

Dim Success                 As Boolean

    Select Case Index
        Case 0
            Success = RecoveryFile_Create(MAIN_DIR_G)
            DoEvents
            If Success Then
                Main.StatusBar1.Panels.Item(3).Text = " The Disaster Recovery File successfully created."
            Else
                Main.StatusBar1.Panels.Item(3).Text = " Creating the Disaster Recovery File failed!"
            End If
        Case 1
            Success = RecoveryFile_Retrieve(MAIN_DIR_G)
            DoEvents
            If Success Then
                Main.StatusBar1.Panels.Item(3).Text = " The Disaster Recovery File successfully loaded."
            Else
                Main.StatusBar1.Panels.Item(3).Text = " Loading the Creating a Disaster Recovery File failed!"
            End If
    End Select
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Load()

    Call form_StayOnTop(DisasterRecovery, True, "C")
    
End Sub


