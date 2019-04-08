VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mate Matcher: Mating Optimization Software"
   ClientHeight    =   7545
   ClientLeft      =   10680
   ClientTop       =   4395
   ClientWidth     =   7140
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      Caption         =   "jcoombs@cns.umass.edu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Label lblCompany 
      Caption         =   "Jason A. Coombs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label lblCopyright 
      Caption         =   "April 5, 2019"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "University of Massachusetts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "U.S. Forest Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Image imgLogo 
      Height          =   6015
      Left            =   240
      Picture         =   "frmAbout.frx":048A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Me.Hide
End Sub

