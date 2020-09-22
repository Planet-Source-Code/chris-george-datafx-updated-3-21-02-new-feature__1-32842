VERSION 5.00
Begin VB.Form frmSQL 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter by SQL"
   ClientHeight    =   1125
   ClientLeft      =   5340
   ClientTop       =   6750
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSQL 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply Filter"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Enter the SQL Statement to Apply:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2940
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pCancel As Boolean

Private Sub cmdCancel_Click()
    pCancel = True
    Me.Visible = False
End Sub

Private Sub cmdApply_Click()
    Me.Visible = False
End Sub


