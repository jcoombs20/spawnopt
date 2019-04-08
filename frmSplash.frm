VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8865
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7305
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
   ScaleHeight     =   8865
   ScaleWidth      =   7305
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
      Height          =   8595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin MSComctlLib.ProgressBar proBar 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   8040
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgLogo 
         Height          =   6015
         Left            =   240
         Picture         =   "frmSplash.frx":048A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6615
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
         TabIndex        =   7
         Top             =   7200
         Width           =   1575
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
         TabIndex        =   6
         Top             =   6960
         Width           =   2415
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
         TabIndex        =   5
         Top             =   7680
         Width           =   3495
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
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
         Left            =   5160
         TabIndex        =   2
         Top             =   6960
         Width           =   1215
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
         TabIndex        =   1
         Top             =   6600
         Width           =   2295
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         TabIndex        =   3
         Top             =   6600
         Width           =   1275
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
    Call getList9
    frmSplash.lblProgress = "Loading Progress: 90%"
    frmSplash.proBar.Value = 90
    DoEvents
    Call getList10
    frmSplash.lblProgress = "Loading Progress: 100%"
    frmSplash.proBar.Value = 100
    DoEvents


    Unload frmSplash
    Load frmMain
End Sub
