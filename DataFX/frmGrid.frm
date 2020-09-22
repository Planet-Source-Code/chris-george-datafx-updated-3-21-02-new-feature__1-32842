VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmGrid 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Data"
   ClientHeight    =   5775
   ClientLeft      =   10650
   ClientTop       =   5685
   ClientWidth     =   6000
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7514
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   1920
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":000C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0124
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":023C
            Key             =   "First"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0354
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":046C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0584
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":069C
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":07B4
            Key             =   "Ascending"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":08CC
            Key             =   "Descending"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":09E4
            Key             =   "Front"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0B04
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0C24
            Key             =   "DeleteFilter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0D38
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   6000
      _CBHeight       =   390
      _Version        =   "6.7.8862"
      Child1          =   "tbrMain"
      MinHeight1      =   330
      Width1          =   3840
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "First"
               Object.ToolTipText     =   "Go to the First Record"
               Object.Tag             =   "First"
               ImageKey        =   "First"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Previous"
               Object.ToolTipText     =   "Go to the Previous Record"
               Object.Tag             =   "Previous"
               ImageKey        =   "Previous"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Next"
               Object.ToolTipText     =   "Go to the Next Record"
               Object.Tag             =   "Next"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Last"
               Object.ToolTipText     =   "Go to the Last Record"
               Object.Tag             =   "Last"
               ImageKey        =   "Last"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Add a New Record"
               Object.Tag             =   "Add"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete the Selected Record"
               Object.Tag             =   "Delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Filter"
               Object.ToolTipText     =   "Filter by SQL"
               Object.Tag             =   "Filter"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DeleteFilter"
               Object.ToolTipText     =   "Remove the Filter"
               Object.Tag             =   "DeleteFilter"
               ImageKey        =   "DeleteFilter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Refresh"
               Object.Tag             =   "Refresh"
               ImageKey        =   "Refresh"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Front"
               Object.ToolTipText     =   "Stay on Top"
               Object.Tag             =   "Front"
               ImageKey        =   "Front"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Data MyDB 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2460
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "frmGrid.frx":1090
      Height          =   4695
      Left            =   0
      OleObjectBlob   =   "frmGrid.frx":10A3
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TableName As String
Public WindowPos As Integer

Private Sub DBGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Cancel
    'show the field description in the status bar
    Me.sbrMain.Panels(2).Text = MyDB.Recordset.Fields(Me.DBGrid.Col).Properties("Description").Value
    Exit Sub
    
Cancel:
    Me.sbrMain.Panels(2).Text = ""
    
End Sub

Private Sub DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Cancel
    'show the current record and the record count in the status bar
    Me.sbrMain.Panels(1).Text = "Record " & Me.DBGrid.Row + 1 & " of " & Me.MyDB.Recordset.RecordCount
    'show the field description in the status bar
    Me.sbrMain.Panels(2).Text = MyDB.Recordset.Fields(Me.DBGrid.Col).Properties("Description").Value
    Exit Sub
    
Cancel:
    Me.sbrMain.Panels(2).Text = ""
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'resize the grid to fit on the form
    With Me.DBGrid
        .Width = Me.Width - 150
        .Height = Me.Height - .Top - Me.sbrMain.Height - 400
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo Err_Handle
    
    Select Case Button.Tag
        Case "First"
            Me.MyDB.Recordset.MoveFirst
        Case "Next"
            Me.MyDB.Recordset.MoveNext
        Case "Previous"
            Me.MyDB.Recordset.MovePrevious
        Case "Last"
            Me.MyDB.Recordset.MoveLast
        Case "Add"
            Me.MyDB.Recordset.AddNew
        Case "Delete"
            Me.MyDB.Recordset.Delete
            Me.MyDB.Refresh
        Case "Filter"
            'set the text of the sql textbox to the current recordsource
            frmSQL.txtSQL = Me.MyDB.RecordSource
            'select the text so the user can quickly delete it
            frmSQL.txtSQL.SelStart = 0
            frmSQL.txtSQL.SelLength = Len(frmSQL.txtSQL)
            'show the filter dialog window
            frmSQL.Show vbModal, Me
            'if the user pressed cancel then unload the form and exit
            If frmSQL.pCancel = True Then
                Unload frmSQL
                Exit Sub
            End If
            'reset the recordsource
            Me.MyDB.RecordSource = frmSQL.txtSQL
            Me.MyDB.Refresh
            'unload the dialog window
            Unload frmSQL
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
        Case "DeleteFilter"
            'reset the recordsource back to the original table
            Me.MyDB.RecordSource = Me.TableName
            Me.MyDB.Refresh
        Case "Refresh"
            Me.MyDB.Refresh
    End Select
    
    Exit Sub
    
Err_Handle:
    On Error Resume Next
    MsgBox "Error: " & Error, vbExclamation + vbOKOnly
    Me.MyDB.RecordSource = Me.TableName
    Me.MyDB.Refresh
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set the status bar caption to the tooltip text of the button
    
    For i = 1 To Me.tbrMain.Buttons.Count
        With Me.tbrMain.Buttons(i)
            If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height Then
                Me.sbrMain.Panels(2).Text = .ToolTipText
                Exit Sub
            End If
        End With
    Next
    Me.sbrMain.Panels(2).Text = ""
End Sub
