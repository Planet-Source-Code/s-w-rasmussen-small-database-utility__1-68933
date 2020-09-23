VERSION 5.00
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "vlmnuplus.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form RecordList 
   Caption         =   " Record List"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   13665
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   9240
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1320
      Width           =   555
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   2790
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   5001
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VLMnuPlus.VLMenuPlus VLMenuPlus1 
      Left            =   960
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   40775.3915856481
      AutoShowHelp    =   0   'False
      UseCustomColors =   -1  'True
      HighlightedTextColor=   2322544
      MenuHighlight   =   12972786
      Language        =   0
      ShowTooltip     =   0   'False
   End
   Begin VB.Menu m_File 
      Caption         =   "File"
      Begin VB.Menu FileItem 
         Caption         =   "Export Record List in Excel Format..."
         Index           =   0
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu FileItem 
         Caption         =   "Add Records to Current Record Set..."
         Index           =   2
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu FileItem 
         Caption         =   "Close Window"
         Index           =   4
      End
   End
   Begin VB.Menu m_close 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "RecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ExcelApp As Excel.Application
Private ExcelWB As Excel.Workbook

Private Sub AppendRecordsToProject()

On Error GoTo errhandler
    
Dim N                       As Long
Dim FolderPath2             As String
    
    FolderPath2 = GetDirectoryPath(cmp2(39, new2(1)))
    
    For N = 1 To lvwMain.ListItems.count
        If FileExist(cmp2(39, new2(N))) Then
            'FileCopy cmp2(39, new2(N)), RECORDS_DIR_G & GetFileTitle(cmp2(39, new2(N)))
        End If
    Next N
    
    Main.cmdAction_Click 4
            
errhandler:
    Exit Sub
End Sub


Private Sub FileItem_Click(Index As Integer)

On Error GoTo errhandler
    
    Select Case Index
    
'EXPORT TO EXCEL
        Case 0
            Call CreateExcelFile
            Main.StatusBar1.Panels(3).Text = " The Record list is now saved as an Excel file in: " & EXCEL_DIR_G
            RecordList.StatusBar1.Panels(3).Text = " The Record list is now saved as an Excel file in: " & EXCEL_DIR_G
                        
'ADD RECORDS TO PROJECT
        Case 2
            Call AppendRecordsToProject
            
'EXIT
        Case 4
            Unload Me
            
    End Select

errhandler:
    Exit Sub
    
End Sub

'------------------------------------------------------------------------------
' Export sequence list as an Excel spreadsheet.
' info: http://msdn2.microsoft.com/en-us/library/aa272268(office.11).aspx
'       http://msdn2.microsoft.com/en-us/library/aa207547(office.11).aspx
' modified 06.12.2005, swr
'------------------------------------------------------------------------------
Public Function CreateExcelFile() As Boolean

On Error GoTo errhandler

If RecordList.lvwMain.ListItems.count < 1 Then GoTo errhandler

Dim N                       As Long
Dim msg                     As String
Dim column                  As Long
Dim row                     As Long
Dim RowCounter              As Long
Dim ColumnCounter           As Long
Dim Response                As Long
Dim ColumnWidth()           As Long
    
    StatusBar1.Panels.Item(3).Text = " Waiting for filename for Excel file..."
    Main.StatusBar1.Panels.Item(3).Text = " Waiting for filename for Excel file..."
    
    ' get save path
    'If Dir(dir_ExcelSave, vbDirectory) = vbNullString Then dir_ExcelSave = dirA_Excel
    Main.cdlg.CancelError = True
    Main.cdlg.DialogTitle = "Export Record List As Excel File"
    Main.cdlg.InitDir = EXCEL_DIR_G
    Main.cdlg.Filename = "*.xls"
    Main.cdlg.Filter = "All files (*.*)|*.*|Excel file (*.xls)|*.xls"
    Main.cdlg.FilterIndex = 2
    Main.cdlg.DefaultExt = "xls"
    Main.cdlg.flags = &H2 Or &H800
    Main.cdlg.ShowSave
         
    ' store excel save path
    EXCEL_DIR_G = GetDirectoryPath(Main.cdlg.Filename)
    
    msg = "Please note that creating the *.xls file takes a while.     " & vbCrLf & vbCrLf & _
         "Observe the cursor: When it changes from an hourglass     " & vbCrLf & _
         "to the normal pointer the process is completed."
    Response = MsgBox(msg, vbExclamation + vbOKOnly, " WARNING")
    
    Screen.MousePointer = 11
    DoEvents
    
    ' create the ExcelApp object
    Set ExcelApp = New Excel.Application
    
    ' disable messages issued by Excel
    ExcelApp.DisplayAlerts = False
    
    StatusBar1.Panels.Item(3).Text = " Creating Excel file, please be patient..."
    Main.StatusBar1.Panels.Item(3).Text = " Creating Excel file, please be patient..."
    
    ' create the ExcelWB object and set its properties
    Set ExcelWB = ExcelApp.Workbooks.add
    
    ' remove excess worksheets (Excel default range 0 - 2 when launched)
    For N = 1 To ExcelApp.ActiveWorkbook.Sheets.count - 1
        ExcelApp.ActiveWorkbook.Sheets(N).Delete
        DoEvents
    Next N
    
    ' set name of active woorksheet (index = 0)
    ExcelApp.ActiveWorkbook.ActiveSheet.Name = "Records"                ' set Name of active sheet to "Record"
    ExcelApp.ActiveWorkbook.ActiveSheet.Cells.NumberFormat = "@"        ' set format of all cells to "Text"
    
    With ExcelWB
        .Title = "Small Database Utility"
        .Subject = "RecordList"
        .SaveAs Main.cdlg.Filename
    End With
                                
    StatusBar1.Panels.Item(3).Text = " Setting minimum column width..."
    Main.StatusBar1.Panels.Item(3).Text = " Setting minimum column width..."
    
    ' set minimum width to 10 for all columns
    ReDim ColumnWidth(1 To RecordList.lvwMain.ColumnHeaders.count) As Long
    For N = LBound(ColumnWidth) To UBound(ColumnWidth)
        ColumnWidth(N) = 10
        DoEvents
    Next N
    
    StatusBar1.Panels.Item(1).Text = " Adding column headers,  please wait..."
    Main.StatusBar1.Panels.Item(1).Text = " Adding column headers,  please wait..."
    
    ' load column titles
    For column = 1 To RecordList.lvwMain.ColumnHeaders.count
        ColumnCounter = ColumnCounter + 1
        ExcelWB.Sheets("Records").Cells(1, ColumnCounter).Value = RecordList.lvwMain.ColumnHeaders(column)
        DoEvents
    Next column
        
    StatusBar1.Panels.Item(3).Text = " Adding Record data,  please wait..."
    Main.StatusBar1.Panels.Item(3).Text = " Adding Record data,  please wait..."
                       
    RowCounter = 1
        
' ADD DATA TO SPREADSHEET
    For row = 1 To RecordList.lvwMain.ListItems.count
        
        ColumnCounter = 1                                                   ' reset column counter for each row
        RowCounter = RowCounter + 1                                         ' increment row counter
        
        ' process description column - first column of sequence list
        ExcelWB.Sheets("Records").Cells(RowCounter, ColumnCounter).Value = CStr(RecordList.lvwMain.ListItems.Item(row))
        
        ' get max length of item for cell formatting
        If Len(RecordList.lvwMain.ListItems.Item(row)) > ColumnWidth(1) Then
            ColumnWidth(1) = Len(RecordList.lvwMain.ListItems.Item(row))
        End If
                    
        ' process subitems - column 2 to last column. Note that the number of subitems is equal to ColumnHeaders.count - 1
        For column = 1 To RecordList.lvwMain.ColumnHeaders.count - 1
        
            ColumnCounter = ColumnCounter + 1
            
            ExcelWB.Sheets("Records").Cells(RowCounter, ColumnCounter).Value = CStr(RecordList.lvwMain.ListItems.Item(row).SubItems(column))
            
            ' get max length of item for cell formatting
            If Len(RecordList.lvwMain.ListItems.Item(row).SubItems(column)) > ColumnWidth(column + 1) Then
                ColumnWidth(column + 1) = Len(RecordList.lvwMain.ListItems.Item(row).SubItems(column))
            End If
            DoEvents
            
        Next column
        StatusBar1.Panels.Item(3).Text = " Adding data from Record: " & Format$(row, "0000")
        Main.StatusBar1.Panels.Item(3).Text = " Adding data from Record: " & Format$(row, "0000")
        DoEvents
    Next row
    
    StatusBar1.Panels.Item(3).Text = " Formatting Excel file,  please wait..."
    Main.StatusBar1.Panels.Item(3).Text = " Formatting Excel file,  please wait..."
    
    ' format spreadsheet
    Dim ExSheet As Excel.Worksheet
       
    Set ExSheet = ExcelApp.Application.ActiveWorkbook.Worksheets(1)
        
    ' set column width and title font for selected columns
    ColumnCounter = 0
    For N = LBound(ColumnWidth) To UBound(ColumnWidth)
        ColumnCounter = ColumnCounter + 1
        ExSheet.Columns(Chr(64 + ColumnCounter) & ":" & Chr(64 + ColumnCounter)).ColumnWidth = ColumnWidth(N) * 1.2
        ExSheet.Cells(1, Chr(64 + ColumnCounter)).Font.Size = 10
        ExSheet.Cells(1, ColumnCounter).Font.Bold = True
        'ExSheet.Cells.Select
        DoEvents
    Next N
            
    ' set font
    ExSheet.Select
    With ExcelApp.Selection.Font
        .Name = "Arial"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
                        
    StatusBar1.Panels.Item(1).Text = " Saving Excel file..."
    Main.StatusBar1.Panels.Item(1).Text = " Saving Excel file..."
    
    ' save spreadsheet
    If TypeName(ExcelApp) <> "Nothing" Then
        ExcelApp.ActiveWorkbook.ActiveSheet.Name = "Records"
        ExcelApp.DefaultFilePath = EXCEL_DIR_G & "RecordList.xls"
        ExcelApp.Application.ActiveWorkbook.SaveAs Main.cdlg.Filename
    End If
    
    StatusBar1.Panels.Item(1).Text = " Done!  Excel file saved in folder: " & EXCEL_DIR_G
    Main.StatusBar1.Panels.Item(1).Text = " Done!  Excel file saved in folder: " & EXCEL_DIR_G
    
    ' close excel application if loaded
    If TypeName(ExcelApp) <> "Nothing" Then
        ExcelApp.Quit
    End If
    
    Set ExcelApp = Nothing
    Set ExSheet = Nothing
    Set ExcelWB = Nothing
    
    Screen.MousePointer = 0
    
    Exit Function
        
errhandler:
    ' close excel application if loaded
    If TypeName(ExcelApp) <> "Nothing" Then
        ExcelApp.Quit
    End If
    
    Set ExcelApp = Nothing
    Set ExSheet = Nothing
    Set ExcelWB = Nothing
    
    Screen.MousePointer = 0
    
    StatusBar1.Panels.Item(1).Text = " Something went wrong!  Export of the Excel file was not completed."
    Main.StatusBar1.Panels.Item(1).Text = " Something went wrong!  Export of the Excel file was not completed."
    
    Exit Function
    
End Function


Private Sub Form_Load()

On Error GoTo errhandler
    
    ' manages forms behaviour
    If MasterUser_G Then
        Record.Visible = False
        Main.Visible = True
    Else
        Record_ShowSingleR (CurrRecord_G)
        Record.Visible = True
        Main.Visible = False
    End If
            
    ' handles different calling functions
    If FromSearch_G Then                ' no action
        ' nothing
    ElseIf FromIncomplete_G Then        ' hide Incomplete
        Incomplete.Visible = False
    Else                                ' hide ComposeList
        ComposeList.Visible = False
    End If
    
    form_StayOnTop RecordList, True, "TR"
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Resize()
    
On Error GoTo errhandler

    If Me.Height < 2000 Then Exit Sub
    If Me.Width < 4000 Then Exit Sub
    
    lvwMain.Top = 90
    lvwMain.Left = 90
    lvwMain.Height = Me.Height - 1200
    lvwMain.Width = Me.Width - 300
    
    If Record.Visible Then
        Record.Top = RecordList.Height + 60
        Record.Left = Me.Width - Record.Width - 30
    End If
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errhandler
        
    If FromSearch_G Then                ' back to Main, no Record
        FromSearch_G = False
        Main.Visible = True
        Record.Visible = False
        
    ElseIf FromIncomplete_G Then        ' back to Incomplete, Main and Record hidden
        FromIncomplete_G = False
        Incomplete.Visible = True
        Main.Visible = False
        Record.Visible = False
        
    Else                                ' back to ComposeList, Main and Record hidden
        ComposeList.Visible = True
        Main.Visible = False
        Record.Visible = False
        
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
On Error GoTo errhandler
        
    ' close personal comments form before selecting a new record
    If InfoBox.Visible Then Unload InfoBox
    
    ' If the ListView is already sorted by the clicked column, just reverse the order. Otherwise, sort the clicked column ascending.
    If lvwMain.Sorted = True And ColumnHeader.SubItemIndex = lvwMain.SortKey Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.Sorted = True
        lvwMain.SortKey = ColumnHeader.SubItemIndex
        lvwMain.SortOrder = lvwAscending
    End If
        
errhandler:
    Exit Sub
End Sub


Private Sub lvwMain_DblClick()

On Error GoTo errhandler
 
Static prevSelNum           As Long
    
    ' close personal comments form before selecting a new record
    If InfoBox.Visible Then Unload InfoBox
    
    If prevSelNum < 1 Or prevSelNum > lvwMain.ListItems.count Then
        prevSelNum = 1
    End If
    
    lvwMain.SmallIcons = Main.ImageList1
    lvwMain.ListItems(prevSelNum).SmallIcon = 3 ' closed = 3, open = 4
    lvwMain.Refresh
    
    ' saves the current record before retrieving a new one
    Call Record_StoreSingleA(CurrRecord_G)
    
    CurrRecord_G = lvwMain.ListItems(lvwMain.SelectedItem.Index).Text
    Call Record_ShowSingleA(CurrRecord_G)
    Call Record_GetNotes(CurrRecord_G, False, False)
    
    lvwMain.ListItems(lvwMain.SelectedItem.Index).SmallIcon = 4
    lvwMain.Refresh
    
    prevSelNum = lvwMain.SelectedItem.Index

    StatusBar1.Panels.Item(1).Text = " Linie = " & lvwMain.SelectedItem.Index
    StatusBar1.Panels.Item(2).Text = " Record = " & lvwMain.ListItems(lvwMain.SelectedItem.Index).Text
    
    Call Record_GetNotes(CurrRecord_G, False, True)

    If Record.Visible = True Then
        Record_ShowSingleR (CurrRecord_G)
    End If
    
errhandler:
    Exit Sub
End Sub


Private Sub m_close_Click()
  
  Unload Me
  
End Sub

Private Sub VLMenuPlus1_SetMenuItemAttributes(ByVal aMenuItem As VLMnuPlus.CMenuItem)

On Error Resume Next

Dim sCaption                As String
            
    VLMenuPlus1.SetImageList Main.ImageList1
    VLMenuPlus1.HighlightStyle = 1
    VLMenuPlus1.HighlightAppearance = 1
    VLMenuPlus1.BitmapBackground = &HE5E1DC
    
    sCaption = VLMenuPlus1.GetCleanCaption(aMenuItem.Caption)
                 
    Select Case sCaption
    
        Case "EXPORT RECORD LIST IN EXCEL FORMAT"
             Set aMenuItem.Picture = Main.ImageList1.ListImages.Item("EXCEL").Picture
             
        Case "ADD RECORDS TO CURRENT RECORD SET"
             Set aMenuItem.Picture = Main.ImageList1.ListImages.Item("REDORDS").Picture
                                       
        Case "CLOSE WINDOW"
             Set aMenuItem.Picture = Main.ImageList1.ListImages.Item("CLOSE").Picture
            
    End Select
        
    If aMenuItem.IsTopLevel = True Then
        Set aMenuItem.Picture = Nothing
    End If
    
End Sub


