VERSION 5.00
Begin VB.Form frmImageList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "McImageList - Test Form !!"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   6900
      TabIndex        =   5
      Top             =   0
      Width           =   6900
      Begin VB.Image Image1 
         Height          =   675
         Left            =   120
         Picture         =   "frmMain.frx":1D8A
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3500
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Single File'd ImageList"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   2325
      End
   End
   Begin ListBox.McImageList McImageList1 
      Left            =   360
      Top             =   4320
      _ExtentX        =   10821
      _ExtentY        =   873
      Images0         =   "frmMain.frx":3595
      Images1         =   "frmMain.frx":392F
      Images2         =   "frmMain.frx":3CC9
      Images3         =   "frmMain.frx":4063
      Images4         =   "frmMain.frx":43FD
      Images5         =   "frmMain.frx":4797
      Images6         =   "frmMain.frx":4B31
      Images7         =   "frmMain.frx":4ECB
      Images8         =   "frmMain.frx":5645
      Images9         =   "frmMain.frx":59DF
      Images10        =   "frmMain.frx":5D79
      Images11        =   "frmMain.frx":6113
      Images12        =   "frmMain.frx":7E1D
      Images13        =   "frmMain.frx":81B7
      Images14        =   "frmMain.frx":8751
      Images15        =   "frmMain.frx":8CEB
      Images16        =   "frmMain.frx":9285
      ImageCount      =   17
   End
   Begin VB.Label Label10 
      Caption         =   "Set the 'Currentimage' and enter the index of new position to the property 'MoveImageTo'"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Label Label9 
      Caption         =   "Move Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "Set the 'CurrentImage' and browse for a new image from the property 'Image'"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label Label8 
      Caption         =   "Set the property 'CurrentImage' and chage the property value of 'RemoveImage' to '[Yes!]'"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label7 
      Caption         =   "Select the property 'AddNewImage' from property window and browse for the image u need."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label6 
      Caption         =   "Remove Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Modify Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Add Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   930
   End
End
Attribute VB_Name = "frmImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
