VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "DataFX"
   ClientHeight    =   3195
   ClientLeft      =   3060
   ClientTop       =   4080
   ClientWidth     =   2625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2625
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   1800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
      Picture         =   "frmMain.frx":08CA
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   1323
      BandCount       =   2
      _CBWidth        =   2625
      _CBHeight       =   750
      _Version        =   "6.7.8862"
      Child1          =   "tbrMain"
      MinHeight1      =   330
      Width1          =   1605
      NewRow1         =   0   'False
      Child2          =   "picTableName"
      MinWidth2       =   1995
      MinHeight2      =   330
      Width2          =   4005
      NewRow2         =   -1  'True
      Begin VB.PictureBox picTableName 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   2370
         TabIndex        =   3
         Top             =   390
         Width           =   2370
         Begin VB.ComboBox cmbTableName 
            Height          =   315
            Left            =   0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   390
         End
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open a Database"
               Object.Tag             =   "Open"
               ImageKey        =   "Open"
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy Field Name"
               Object.Tag             =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               Object.Tag             =   "Find"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindNext"
               Object.ToolTipText     =   "Find Next"
               Object.Tag             =   "FindNext"
               ImageKey        =   "FindNext"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DataSheet"
               Object.ToolTipText     =   "DataSheet View"
               Object.Tag             =   "DataSheet"
               ImageKey        =   "DataSheet"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Front"
               Object.ToolTipText     =   "Stay On Top"
               Object.Tag             =   "Front"
               ImageKey        =   "Front"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B7A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":424E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43AA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A7E
            Key             =   "Front"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B92
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CB2
            Key             =   "FindNext"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E12
            Key             =   "DataSheet"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4128
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileName 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Developed by Chris George
' 03/19/02
'
' Feel free to use any code here in any of your projects
' If you like this program, all I ask is that you please go and leave a comment here:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=32842&lngWId=1

Public DBPath As String
Public WindowPos As Integer
Public LastSearch As String
Public WholeWord As Integer

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    Me.Form_Resize
End Sub

Private Sub cmbTableName_Click()
    'get fields from the table
    Dim MyDB As Database
    Dim MyRecSet As Recordset
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo Err_Handle
    
    'open the database
    Set MyDB = OpenDatabase(DBPath)
    Set MyRecSet = MyDB.OpenRecordset(Me.cmbTableName.Text)
    
    'clear the list
    Me.lvwFields.ListItems.Clear
    
    'clear the column headers
    Me.lvwFields.ColumnHeaders.Clear
    
    Me.lvwFields.ColumnHeaders.Add , , "Field Name"
    Me.lvwFields.ColumnHeaders.Add , , "Description"
    
    'get all of the field properties
    For i = 0 To MyRecSet.Fields(0).Properties.Count - 1
        If MyRecSet.Fields(0).Properties(i).Name <> "Name" And MyRecSet.Fields(0).Properties(i).Name <> "Value" And MyRecSet.Fields(0).Properties(i).Name <> "Description" Then
            Me.lvwFields.ColumnHeaders.Add , , MyRecSet.Fields(0).Properties(i).Name
        End If
    Next
    
    'list all of the fields
    For j = 0 To MyRecSet.Fields.Count
        On Error Resume Next
        Me.lvwFields.ListItems.Add , "K" & j, MyRecSet.Fields(j).Name
        Me.lvwFields.ListItems("K" & j).SubItems(1) = MyRecSet.Fields(j).Properties("Description").Value
        For i = 2 To Me.lvwFields.ColumnHeaders.Count - 1
            Me.lvwFields.ListItems("K" & j).SubItems(i) = MyRecSet.Fields(Me.lvwFields.ListItems("K" & j).Text).Properties(Me.lvwFields.ColumnHeaders(i + 1)).Value
        Next
    Next
    
    Me.lvwFields.BackColor = vbWindowBackground
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing
        
    Exit Sub
    
Err_Handle:
    MsgBox "Error opening database: " & Error, vbExclamation + vbOKOnly
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing
End Sub

Public Function FileExists(strPath As String) As Boolean
    'checks to see if a file exists
    FileExists = Not (Dir(strPath) = "")
End Function

Private Sub Form_Load()
    On Error Resume Next
    'resize the combo box
    Me.cmbTableName.Width = Me.picTableName.Width
    'load the recent file list
    Me.LoadRecentFiles
    'show the form
    Me.Visible = True
    'show the open dialog
    Me.OpenDB
End Sub

Public Sub Form_Resize()
    'resize the controls when the form is resized
    On Error Resume Next
    Me.lvwFields.Top = Me.cbrMain.Top + Me.cbrMain.Height + 25
    Me.lvwFields.Width = Me.Width - 100
    Me.lvwFields.Height = Me.Height - Me.lvwFields.Top - sbrMain.Height - 400
End Sub

Public Sub LoadRecentFiles()
    Dim FileNum As Integer
    Dim i As Integer
    Dim Index As Integer
    Dim strInput As String
    On Error GoTo Err_Handle
    
    'get a free file number
    FileNum = FreeFile
    Index = 0
    'This sub will go get the most recent files
    If FileExists(App.Path & "\rfiles") Then
        Open App.Path & "\rfiles" For Input As #FileNum
            Do
                'loop through the file and read in each line
                Line Input #FileNum, strInput
                'get the next menu to load
                If Index > 0 Then Load Me.mnuFileName(Index)
                'set the caption to the filename
                Me.mnuFileName(Index).Caption = Trim(strInput)
                Index = Index + 1
            Loop Until EOF(FileNum)
        Close #FileNum
    End If
    
Err_Handle:
    Close
End Sub

Private Sub lvwFields_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Me.lvwFields.Sorted = True
    If Me.lvwFields.SortKey = ColumnHeader.Index - 1 Then
        If Me.lvwFields.SortOrder = lvwAscending Then
            Me.lvwFields.SortOrder = lvwDescending
        Else
            Me.lvwFields.SortOrder = lvwAscending
        End If
    End If
    Me.lvwFields.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwFields_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.sbrMain.Panels(1).Text = ""
End Sub

Private Sub lvwFields_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectCopy
    Data.SetData Me.lvwFields.SelectedItem.Text
End Sub

Public Sub OpenDB(Optional FileName As String)
    'This sub opens a database selected by the user
    Dim MyDB As Database
    Dim MyRecSet As Recordset
    Dim i As Integer
    Dim Count As Integer
    Dim NewFile As Boolean
    
    On Error GoTo Cancel
    If FileName = "" Then
        'set the filter for the common dialog window for access database
        cdMain.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        cdMain.Filter = "Microsoft Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
        cdMain.FilterIndex = 0
        'get the filename to open
        cdMain.ShowOpen
        FileName = cdMain.FileName
        NewFile = True
    End If
    
    On Error GoTo Err_Handle
    
    'make sure the file exists
    If FileExists(FileName) = False Then
        MsgBox FileName & " does not exist.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    'set the global dbpath
    DBPath = FileName
    
    'attempt to open the database
    Set MyDB = OpenDatabase(FileName)
    
    'clear the combo box
    Me.cmbTableName.Clear
    
    'load the combo box with table names
    For i = 0 To MyDB.TableDefs.Count - 1
        If UCase(Mid(MyDB.TableDefs(i).Name, 1, 4)) <> "MSYS" And UCase(Mid(MyDB.TableDefs(i).Name, 1, 1)) <> "~" Then
            Me.cmbTableName.AddItem MyDB.TableDefs(i).Name
        End If
    Next
    
    'load the combo box with query names
    For i = 0 To MyDB.QueryDefs.Count - 1
        If UCase(Mid(MyDB.QueryDefs(i).Name, 1, 4)) <> "MSYS" And UCase(Mid(MyDB.QueryDefs(i).Name, 1, 1)) <> "~" Then
            Me.cmbTableName.AddItem MyDB.QueryDefs(i).Name
        End If
    Next
    
    'select the first item in the list
    Me.cmbTableName.ListIndex = 0
    
    'set the caption
    Me.Caption = "DataFX [" & Dir(FileName) & "]"
    
    If NewFile = True Then
        'load a new menu item for the file
        Count = Me.mnuFileName.Count
        If Count < 6 Then
            'load a new menu item if the first menu is already filled with a file name
            If Me.mnuFileName(0).Caption <> "-" Then
                Load mnuFileName(Count)
            Else
                Count = Count - 1
            End If
            
            For i = Count To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
            Next
            mnuFileName(0).Caption = FileName
        Else
            For i = 5 To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
            Next
            mnuFileName(0).Caption = FileName
        End If
    End If
    'save the recent file list
    Me.SaveRecentFiles
    
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing
        
    Exit Sub
    
Err_Handle:
    MsgBox "Error opening database: " & Error, vbExclamation + vbOKOnly
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing

Cancel:

End Sub


Private Sub mnuFileName_Click(Index As Integer)
    If InStr(1, Me.mnuFileName(Index).Caption, ":") > 0 Then
        Me.OpenDB Me.mnuFileName(Index).Caption
    Else
        Me.OpenDB App.Path & "\" & Me.mnuFileName(Index).Caption
    End If
End Sub

Private Sub picTableName_Resize()
    On Error Resume Next
    'resize the combo box when the container is resized
    Me.cmbTableName.Width = Me.picTableName.Width
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lRetVal As Long
    Dim SearchFor As String
    Dim SearchIndex As Integer
    
    Select Case Button.Tag
        Case "Open"
            Me.OpenDB
        Case "Copy"
            If Me.lvwFields.ListItems.Count > 0 Then
                Clipboard.Clear
                Clipboard.SetText Me.lvwFields.SelectedItem.Text
            End If
        Case "Find"
            If Me.lvwFields.ListItems.Count < 0 Then Exit Sub
            'put the last item searched for in the search box
            'and then select it so the user can quickly delete it
            frmSearch.txtSearch.Text = LastSearch
            frmSearch.txtSearch.SelStart = 0
            frmSearch.txtSearch.SelLength = Len(LastSearch)
            frmSearch.chkWholeWord = WholeWord
            'show the dialog window to get what the user is searching for
            frmSearch.Show vbModal, Me
            With frmSearch
                'if the user clicked cancel then exit the sub
                If .pCancel = True Then Exit Sub
                LastSearch = .txtSearch
                WholeWord = .chkWholeWord.Value
                SearchFor = .txtSearch.Text
                'if the user doesn't have anything selected then make it 1
                'otherwise start from the next item on the list with the search
                'so the user can find the next item
                If Me.lvwFields.SelectedItem Is Nothing Then
                    SearchIndex = 1
                Else
                    SearchIndex = Me.lvwFields.SelectedItem.Index + 1
                    If SearchIndex > Me.lvwFields.ListItems.Count Then SearchIndex = 1
                End If
                'if the user checked find whole word only then only search for the whole word, else do a partial search
                If .chkWholeWord.Value = vbChecked Then
                    Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwWhole)
                Else
                    Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwPartial)
                End If
                'if nothing was found, tell the user, otherwise display it
                If Me.lvwFields.SelectedItem Is Nothing Then
                    MsgBox "String was not found.", vbInformation + vbOKOnly
                Else
                    Me.lvwFields.SelectedItem.EnsureVisible
                End If
            End With
        Case "FindNext"
            If Me.lvwFields.ListItems.Count < 0 Then Exit Sub
            'if the user doesn't have anything selected then make it 1
            'otherwise start from the next item on the list with the search
            'so the user can find the next item
            If Me.lvwFields.SelectedItem Is Nothing Then
                SearchIndex = 1
            Else
                SearchIndex = Me.lvwFields.SelectedItem.Index + 1
                If SearchIndex > Me.lvwFields.ListItems.Count Then SearchIndex = 1
            End If
            SearchFor = LastSearch
            'if the user checked find whole word only then only search for the whole word, else do a partial search
            If WholeWord = vbChecked Then
                Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwWhole)
            Else
                Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwPartial)
            End If
            'if nothing was found, tell the user, otherwise display it
            If Me.lvwFields.SelectedItem Is Nothing Then
                MsgBox "String was not found.", vbInformation + vbOKOnly
            Else
                Me.lvwFields.SelectedItem.EnsureVisible
            End If
        Case "Front"
            If Me.WindowPos = HWND_TOPMOST Then
                lRetVal = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
                Me.WindowPos = HWND_NOTOPMOST
                Button.Image = "Front"
                Button.ToolTipText = "Stay on Top"
            Else
                lRetVal = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
                Me.WindowPos = HWND_TOPMOST
                Button.Image = "Back"
                Button.ToolTipText = "Don't Stay on Top"
            End If
            Me.cbrMain.Refresh
        Case "DataSheet"
            'create a new datasheet
            Dim NewSheet As New frmGrid
            NewSheet.MyDB.DatabaseName = Me.DBPath
            NewSheet.MyDB.RecordSource = Me.cmbTableName.Text
            NewSheet.MyDB.Refresh
            NewSheet.TableName = Me.cmbTableName.Text
            NewSheet.sbrMain.Panels(1).Text = NewSheet.MyDB.Recordset.RecordCount & " Records"
            NewSheet.Caption = "DataSheet - " & Me.cmbTableName.Text
            NewSheet.Visible = True
    End Select
End Sub

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Me.PopupMenu Me.mnuFile, , Button.Left + Me.tbrMain.Left + Me.cbrMain.Left, Button.Top + Button.Height + Me.tbrMain.Top + Me.cbrMain.Top
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set the status bar caption to the tooltip text of the button
    
    For i = 1 To Me.tbrMain.Buttons.Count
        With Me.tbrMain.Buttons(i)
            If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height Then
                Me.sbrMain.Panels(1).Text = .ToolTipText
                Exit Sub
            End If
        End With
    Next
    Me.sbrMain.Panels(1).Text = ""
End Sub

Public Sub SaveRecentFiles()
    Dim FileNum As Integer
    Dim i As Integer
    Dim Index As Integer
    Dim strInput As String
    On Error GoTo Err_Handle
    
    'get a free file number
    FileNum = FreeFile
    
    'This sub will save the most recent files

    Open App.Path & "\rfiles" For Output As #FileNum
        For i = 0 To Me.mnuFileName.Count - 1
            Print #FileNum, Me.mnuFileName(i).Caption
        Next
    Close #FileNum
    
    Exit Sub
    
Err_Handle:
    Close
End Sub
