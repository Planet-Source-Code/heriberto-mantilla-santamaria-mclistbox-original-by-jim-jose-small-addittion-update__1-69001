VERSION 5.00
Begin VB.Form frmBrowser 
   Caption         =   "Select file... McBrowser!!"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   7095
   End
   Begin ListBox.McListBox McListBox1 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      IconFolderOpen  =   "frmBrowser.frx":1D8A
      IconFolderClose =   "frmBrowser.frx":279C
      IconRemovable   =   "frmBrowser.frx":31AE
      IconCDROM       =   "frmBrowser.frx":3A88
      IconDriver      =   "frmBrowser.frx":4362
      CustomIcons     =   -1  'True
      Picture         =   "frmBrowser.frx":4C3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   131072
      FullRowSelect   =   -1  'True
      IconFocus       =   0   'False
      RowHeight       =   24
      GridLines       =   -1  'True
      ShowIcon        =   -1  'True
      Mode            =   2
      Path            =   "D:\"
      ShowSystemFiles =   -1  'True
      ShowHiddenFiles =   -1  'True
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdDone_Click()
    frmListBox.Show
    Unload Me
End Sub

Private Sub McListBox1_SelChange()
    txtPath = McListBox1.Text
End Sub
