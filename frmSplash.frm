VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin MSComctlLib.ProgressBar proBar 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   4320
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Caption         =   "Loading Progress: 0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   240
         Picture         =   "frmSplash.frx":5C12
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2017"
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
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblCompany 
         Caption         =   "Jason Coombs"
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
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.1"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "SpawnOpt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   765
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_GotFocus()
    Call getFList4
    frmSplash.lblProgress = "Loading Progress: 10%"
    frmSplash.proBar.Value = 10
    DoEvents
    Call getList1
    frmSplash.lblProgress = "Loading Progress: 10%"
    frmSplash.proBar.Value = 10
    DoEvents
    Call getList2
    frmSplash.lblProgress = "Loading Progress: 20%"
    frmSplash.proBar.Value = 20
    DoEvents
    Call getList3
    frmSplash.lblProgress = "Loading Progress: 30%"
    frmSplash.proBar.Value = 30
    DoEvents
    Call getList4
    frmSplash.lblProgress = "Loading Progress: 40%"
    frmSplash.proBar.Value = 40
    DoEvents
    Call getList5
    frmSplash.lblProgress = "Loading Progress: 50%"
    frmSplash.proBar.Value = 50
    DoEvents
    Call getList6
    frmSplash.lblProgress = "Loading Progress: 60%"
    frmSplash.proBar.Value = 60
    DoEvents
    Call getList7
    frmSplash.lblProgress = "Loading Progress: 70%"
    frmSplash.proBar.Value = 70
    DoEvents
    Call getList8
    frmSplash.lblProgress = "Loading Progress: 80%"
    frmSplash.proBar.Value = 80
    DoEvents
    'Call getList9
    frmSplash.lblProgress = "Loading Progress: 90%"
    frmSplash.proBar.Value = 90
    DoEvents
    'Call getList10
    frmSplash.lblProgress = "Loading Progress: 100%"
    frmSplash.proBar.Value = 100
    DoEvents

    Unload frmSplash
    Load frmMain
End Sub
