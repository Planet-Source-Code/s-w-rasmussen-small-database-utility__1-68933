VERSION 5.00
Begin VB.Form Advanced 
   BackColor       =   &H00EEE8E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Advanced Functions"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
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
   ForeColor       =   &H000000C0&
   Icon            =   "Advanced.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Merge Databases"
      Height          =   300
      Index           =   1
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2860
      Width           =   1775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   180
      Picture         =   "Advanced.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3540
      Width           =   510
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Manage Users"
      Height          =   300
      Index           =   0
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   180
      Width           =   1775
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Find Duplicates"
      Height          =   300
      Index           =   3
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   1775
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Find Incomplete"
      Height          =   300
      Index           =   6
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1260
      Width           =   1775
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Upload Settings"
      Height          =   300
      Index           =   8
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1775
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Disaster Recovery"
      Height          =   300
      Index           =   2
      Left            =   3820
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2340
      Width           =   1775
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Close"
      Height          =   300
      Index           =   7
      Left            =   4690
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3740
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE8E6&
      Caption         =   "Merge the current local and the remote database. Differences are transmitted from Local to Remote."
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   180
      TabIndex        =   12
      ToolTipText     =   " WARNING ! - This option is only intended for experts !!"
      Top             =   2860
      Width           =   3600
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE8E6&
      Caption         =   "Create /  Load disaster recovery copy of database"
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   2
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   " WARNING ! - This option is only intended for experts !!"
      Top             =   2340
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Enter / Edit settings for up- and download of database to / from website"
      Height          =   435
      Index           =   8
      Left            =   180
      TabIndex        =   4
      Top             =   1800
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Find records with empty fields in current database"
      Height          =   435
      Index           =   6
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Find duplicate records in database.  Looks for Name and Address"
      Height          =   435
      Index           =   3
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEE8E6&
      Caption         =   "Enter / Edit administrator name(s) and password(s)"
      Height          =   435
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   3600
   End
End
Attribute VB_Name = "Advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Find_Duplicates()

On Error GoTo errhandler

Dim N                       As Long
Dim P                       As Long
Dim count                   As Long
Dim tmp                     As String
Dim a()                     As String
Dim itmX                    As Object
    
    ' clear lvwValidate
    Validate.lvwValidate.ColumnHeaders.Clear
    Validate.lvwValidate.ListItems.Clear
    
    ' add column headers: main.lblfield().caption
    Validate.lvwValidate.ColumnHeaders.add , , "Number", 900, lvwColumnLeft
    Validate.lvwValidate.ColumnHeaders.add , , "Record 1", 900, lvwColumnRight
    Validate.lvwValidate.ColumnHeaders.add , , "Record 2", 900, lvwColumnRight
    Validate.lvwValidate.ColumnHeaders.add , , "Name / Address", 2500, lvwColumnLeft
    
    ' set View property to Report.
    Validate.lvwValidate.View = lvwReport
        
    ReDim a(1 To UBound(nr))
        
    ' First, Last
    For N = 1 To UBound(nr)
        tmp = nr(N).txtField(1) & Space$(1) & _
              nr(N).txtField(2)
        tmp = Replace(tmp, ".", Space$(1))                          ' remove "."
        tmp = Replace(tmp, ",", Space$(1))                          ' remove ","
        tmp = Replace(tmp, "&", Space$(1))                          ' remove "&"
        Do While InStr(tmp, Space$(2))
            tmp = Replace(tmp, Space$(2), Space$(1))                ' remove double spaces
        Loop
        a(N) = Trim$(tmp)
    Next N
       
    ' find identical records
    count = 0
    For N = 1 To UBound(nr)
        For P = N + 1 To UBound(nr) - 1
            If Len(a(N)) > 0 And a(N) = a(P) Then
                count = count + 1
                Set itmX = Validate.lvwValidate.ListItems.add()
                itmX.Text = Format$(count, "0000")
                itmX.SubItems(1) = N
                itmX.SubItems(2) = P
                itmX.SubItems(3) = a(P)
            End If
        Next P
    Next N
        
    ReDim a(1 To UBound(nr))
        
    ' compare: First, Last, Street, Number, Floor
    For N = 1 To UBound(nr)
        tmp = nr(N).txtField(1) & Space$(1) & _
              nr(N).txtField(2) & Space$(1) & _
              nr(N).txtField(4) & Space$(1) & _
              nr(N).txtField(5) & Space$(1) & _
              nr(N).txtField(6)
        tmp = Replace(tmp, ".", Space$(1))                          ' remove "."
        tmp = Replace(tmp, ",", Space$(1))                          ' remove ","
        tmp = Replace(tmp, "&", Space$(1))                          ' remove "&"
        Do While InStr(tmp, Space$(2))
            tmp = Replace(tmp, Space$(2), Space$(1))                ' remove double spaces
        Loop
        a(N) = Trim$(tmp)
    Next N
    
    ' find identical records
    For N = 1 To UBound(nr)
        For P = N + 1 To UBound(nr) - 1
            If Len(a(N)) > 0 And a(N) = a(P) Then
                count = count + 1
                Set itmX = Validate.lvwValidate.ListItems.add()
                itmX.Text = Format$(count, "0000")
                itmX.SubItems(1) = N
                itmX.SubItems(2) = P
                itmX.SubItems(3) = a(P)
            End If
        Next P
    Next N
        
    Validate.Show 1
        
errhandler:
    Exit Sub
End Sub










Public Sub cmdAction_Click(Index As Integer)

On Error GoTo errhandler
    
Dim msg                     As String
Dim Response                As Long

    Select Case Index
    
'MANAGE USERS
        Case 0
            Me.Visible = False
            Main.StatusBar1.Panels.Item(3).Text = " Edit database administrator information..."
            MasterUsers.Show 1
            Me.Visible = True
            
'MERGE DATABASES
        Case 1
            DBCompare = UpdateRemoteRecords(MAIN_DIR_G & "Database.zlb", uploaddb.RemoteFileName)
            If DBCompare.RecordsIdentical Then
                msg = "The two databases are identical     "
            Else
                msg = "The local and the remote copy of the database were different.     " & vbCrLf & vbCrLf & _
                      "A total of " & DBCompare.NumUpdatedRecords & " records in the local copy were included " & vbCrLf & _
                      "in the website copy of the database."
            End If
            Response = MsgBox(msg, vbInformation + vbOKOnly, " DATABASE UPDATE")
                
'CREATE RECOVERY COPY
        Case 2
            Me.Visible = False
            Main.StatusBar1.Panels.Item(3).Text = " Create Disaster Recovery copy of database..."
            DisasterRecovery.Show 1
            Me.Visible = True
            
'FIND DUPLICATE RECORDS
        Case 3
            Me.Visible = False
            Main.StatusBar1.Panels.Item(3).Text = " Search for duplicates in current Record set..."
            Call Find_Duplicates
            Advanced.Visible = False
            AppConfig.Visible = False
            Me.Visible = True
            
'INCOMPLETE RECORDS
        Case 6
            Main.Visible = False
            Me.Visible = False
            Main.StatusBar1.Panels.Item(3).Text = " Find incomplete records in database..."
            FromIncomplete_G = True
            Incomplete.Show
            
'UPLOAD SETTINGS
        Case 8
            Me.Visible = False
            Main.StatusBar1.Panels.Item(3).Text = " Enter settings for up- and download..."
            UploadInfo.Show 1
            Main.FileItem(12).Visible = True
            Main.FileItem(13).Visible = True
            Me.Visible = True
            
'UNLOAD FORM
        Case 7
            FromIncomplete_G = False
            Unload Me
            
    End Select
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Load()

    Call form_StayOnTop(Advanced, True, "C")
            
    If WebServerConnectionOK_G Then
        Me.Height = 4560
        Me.Picture1.Top = 3540
        Me.cmdAction(7).Top = 3740
        cmdAction(1).Visible = True
        Label1(1).Visible = True
        cmdAction(2).Visible = True
        Label1(2).Visible = True
        cmdAction(8).Visible = True
        Label1(8).Visible = True
    Else
        Me.Height = 4560 - 1620
        Me.Picture1.Top = 3540 - 1620
        Me.cmdAction(7).Top = 3740 - 1620
        cmdAction(1).Visible = False
        Label1(1).Visible = False
        cmdAction(2).Visible = False
        Label1(2).Visible = False
        cmdAction(8).Visible = False
        Label1(8).Visible = False
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo errhandler

Dim N                       As Long

    ' scan forms collection and close all loaded forms - except Main - which is in the process of closing anyway
    For N = Forms.count - 1 To 0 Step -1
        If Forms(N).Name <> "Main" Then
            Unload Forms(N)
            WaitABit 0.2
        End If
    Next N
    
errhandler:
    Exit Sub
End Sub


Public Sub Label1_Click(Index As Integer)

End Sub


