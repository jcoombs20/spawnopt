VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmMain 
   Caption         =   "Spawning Optimization"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   16410
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   16410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAvgPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   253
      TabStop         =   0   'False
      ToolTipText     =   "Average proportion of shared alleles between all spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   9
      Left            =   15960
      TabIndex        =   48
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   8
      Left            =   15960
      TabIndex        =   44
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   7
      Left            =   15960
      TabIndex        =   40
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   6
      Left            =   15960
      TabIndex        =   36
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   5
      Left            =   15960
      TabIndex        =   32
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   4
      Left            =   15960
      TabIndex        =   28
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   3
      Left            =   15960
      TabIndex        =   24
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   2
      Left            =   15960
      TabIndex        =   20
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   1
      Left            =   15960
      TabIndex        =   16
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   9
      Left            =   15240
      TabIndex        =   47
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   8
      Left            =   15240
      TabIndex        =   43
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   7
      Left            =   15240
      TabIndex        =   39
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   6
      Left            =   15240
      TabIndex        =   35
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   5
      Left            =   15240
      TabIndex        =   31
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   4
      Left            =   15240
      TabIndex        =   27
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   3
      Left            =   15240
      TabIndex        =   23
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   2
      Left            =   15240
      TabIndex        =   19
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   1
      Left            =   15240
      TabIndex        =   15
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   0
      Left            =   15960
      TabIndex        =   12
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   0
      Left            =   15240
      TabIndex        =   11
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CheckBox chkPrefix 
      Caption         =   "Include Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   248
      Top             =   360
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdRelQuarts 
      Caption         =   "Refresh Relatedness Quartiles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   $"Main Form.frx":058A
      Top             =   8520
      Width           =   3735
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   10920
      TabIndex        =   45
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   10920
      TabIndex        =   41
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   10920
      TabIndex        =   37
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   10920
      TabIndex        =   33
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   10920
      TabIndex        =   29
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   10920
      TabIndex        =   25
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   10920
      TabIndex        =   21
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   10920
      TabIndex        =   17
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   10920
      TabIndex        =   13
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   10920
      TabIndex        =   9
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPitCur2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      TabIndex        =   107
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number of most recently scanned individual"
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPShare 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSpawnUp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Update Spawning Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   49
      ToolTipText     =   "Update 'tblBroodMating' with highlighted spawning pairs"
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   9
      Left            =   14160
      TabIndex        =   46
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   8
      Left            =   14160
      TabIndex        =   42
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   7
      Left            =   14160
      TabIndex        =   38
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   6
      Left            =   14160
      TabIndex        =   34
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   5
      Left            =   14160
      TabIndex        =   30
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   4
      Left            =   14160
      TabIndex        =   26
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   3
      Left            =   14160
      TabIndex        =   22
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   2
      Left            =   14160
      TabIndex        =   18
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   1
      Left            =   14160
      TabIndex        =   14
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   0
      Left            =   14160
      TabIndex        =   10
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMale 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning male"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Tag number of spawning female"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdInputPIT 
      Caption         =   "Input Spawner PIT Tags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Open form to input spawning individuals data"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mating Design"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 to 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   59
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 to 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   58
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 to 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   57
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 to 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   56
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 to 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   55
         ToolTipText     =   "Optimizes matings  in a 1 female to 1 male relationship"
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Main Form.frx":0667
      Left            =   3360
      List            =   "Main Form.frx":068C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Number of males to optimize for spawning"
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox cmbFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Main Form.frx":06B3
      Left            =   3360
      List            =   "Main Form.frx":06D5
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Number of females to optimize for spawning"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtPitCur 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      LinkItem        =   "Field(1)"
      LinkTopic       =   "winwedge|COM1"
      TabIndex        =   52
      ToolTipText     =   "PIT tag of currently scanned individual"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   50
      ToolTipText     =   "Exit out of the optimization software"
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6240
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6240
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6240
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6240
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6240
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6240
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6240
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   6240
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6240
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   6240
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9000
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9000
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9000
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9000
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9000
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9000
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9000
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9000
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   9000
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtLocMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   9000
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtCommentInt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12000
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   0
      Left            =   12360
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   140
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   141
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   3
      Left            =   12360
      TabIndex        =   142
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   4
      Left            =   12360
      TabIndex        =   143
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   5
      Left            =   12360
      TabIndex        =   144
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   6
      Left            =   12360
      TabIndex        =   145
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   7
      Left            =   12360
      TabIndex        =   146
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   8
      Left            =   12360
      TabIndex        =   147
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   9
      Left            =   12360
      TabIndex        =   148
      TabStop         =   0   'False
      ToolTipText     =   "Opens form to add a comment for that mating pair"
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5520
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   5520
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5520
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPicFem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5880
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5880
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   5880
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5880
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   5880
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   5880
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   5880
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   5880
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   5880
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightFem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   5880
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7560
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7560
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7560
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7560
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7560
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7560
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   7560
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   7560
      TabIndex        =   176
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   7560
      TabIndex        =   177
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtPicMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   7560
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9000
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9000
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9000
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9000
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9000
      TabIndex        =   183
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9000
      TabIndex        =   184
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9000
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9000
      TabIndex        =   186
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   9000
      TabIndex        =   187
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtWeightMale 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   9000
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6600
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6600
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6600
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6600
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6600
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6600
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6600
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   6600
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6600
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   6600
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   8280
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   8280
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   8280
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   8280
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   8280
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   8280
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   8280
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   8280
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   8280
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtDrainageM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   8280
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   9720
      TabIndex        =   210
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9720
      TabIndex        =   211
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   9720
      TabIndex        =   212
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   9720
      TabIndex        =   213
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   9720
      TabIndex        =   214
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9720
      TabIndex        =   215
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9720
      TabIndex        =   216
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtQuartile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9720
      TabIndex        =   218
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   0
      Left            =   11040
      TabIndex        =   217
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   1
      Left            =   11040
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   2
      Left            =   11040
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   3
      Left            =   11040
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   4
      Left            =   11040
      TabIndex        =   222
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   5
      Left            =   11040
      TabIndex        =   223
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   6
      Left            =   11040
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   7
      Left            =   11040
      TabIndex        =   225
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   8
      Left            =   11040
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   9
      Left            =   11040
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6840
      TabIndex        =   228
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6840
      TabIndex        =   229
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6840
      TabIndex        =   230
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6840
      TabIndex        =   231
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6840
      TabIndex        =   232
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6840
      TabIndex        =   233
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6840
      TabIndex        =   234
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   6840
      TabIndex        =   235
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6840
      TabIndex        =   236
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtfYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   6840
      TabIndex        =   237
      Text            =   "Text1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   8640
      TabIndex        =   238
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   8640
      TabIndex        =   239
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   8640
      TabIndex        =   240
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   8640
      TabIndex        =   241
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   8640
      TabIndex        =   242
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   8640
      TabIndex        =   243
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   8640
      TabIndex        =   244
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   8640
      TabIndex        =   245
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   8640
      TabIndex        =   246
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtmYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   8640
      TabIndex        =   247
      Text            =   "Text1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblAvgPShare 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Average"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   254
      ToolTipText     =   "Average proportion of shared alleles between all spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Family"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   252
      ToolTipText     =   "Tag number of spawning female"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15720
      TabIndex        =   251
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   1290
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Released"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   250
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fem."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15000
      TabIndex        =   249
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   1290
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4440
      X2              =   16440
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Flag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   209
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   9
      Left            =   12960
      Picture         =   "Main Form.frx":06F8
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   8
      Left            =   12960
      Picture         =   "Main Form.frx":13BA
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   7
      Left            =   12960
      Picture         =   "Main Form.frx":207C
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   6
      Left            =   12960
      Picture         =   "Main Form.frx":2D3E
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   5
      Left            =   12960
      Picture         =   "Main Form.frx":3A00
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   4
      Left            =   12960
      Picture         =   "Main Form.frx":46C2
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   3
      Left            =   12960
      Picture         =   "Main Form.frx":5384
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   2
      Left            =   12960
      Picture         =   "Main Form.frx":6046
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   1
      Left            =   12960
      Picture         =   "Main Form.frx":6D08
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   0
      Left            =   12960
      Picture         =   "Main Form.frx":79CA
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2625
      Left            =   120
      Picture         =   "Main Form.frx":868C
      Stretch         =   -1  'True
      ToolTipText     =   "Picture of an Atlantic salmon. credit ""http://www.davidmillerart.co.uk/game_fish_paintings.htm"""
      Top             =   6360
      Width           =   4260
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Alleles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   96
      ToolTipText     =   "The proportion of shared alleles between the mating pair"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prop. Shared"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   95
      ToolTipText     =   "Proportion of shared alleles between spawning pairs (0 = no shared alleles, 1 = all alleles shared)"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Spawned"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   94
      ToolTipText     =   "Check box to highlight pairs that have already been spawned"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   93
      ToolTipText     =   "Tag number of spawning male"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   92
      ToolTipText     =   "Tag number of spawning female"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4560
      TabIndex        =   71
      ToolTipText     =   "Family ID of mating pair"
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4560
      TabIndex        =   70
      ToolTipText     =   "Family ID of mating pair"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   69
      ToolTipText     =   "Family ID of mating pair"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4560
      TabIndex        =   68
      ToolTipText     =   "Family ID of mating pair"
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4560
      TabIndex        =   67
      ToolTipText     =   "Family ID of mating pair"
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   66
      ToolTipText     =   "Family ID of mating pair"
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   65
      ToolTipText     =   "Family ID of mating pair"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   64
      ToolTipText     =   "Family ID of mating pair"
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   63
      ToolTipText     =   "Family ID of mating pair"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblFem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   62
      ToolTipText     =   "Family ID of mating pair"
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current PIT Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   61
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4440
      X2              =   16440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Spawning Input Criteria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   60
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Spawning Males"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   54
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Spawning Females"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   53
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   5655
      Left            =   120
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   7320
      TabIndex        =   117
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   7320
      TabIndex        =   116
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   7320
      TabIndex        =   115
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   7320
      TabIndex        =   114
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7320
      TabIndex        =   113
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7320
      TabIndex        =   112
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7320
      TabIndex        =   111
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7320
      TabIndex        =   110
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7320
      TabIndex        =   109
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7320
      TabIndex        =   108
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu databaseSettings 
         Caption         =   "&Database Settings"
         HelpContextID   =   1
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appWin, i As Long, strFile As String, Chan, MyData, PauseTime As Single, Start As Single
Dim Finish As Single, wrkJet As Workspace, dbsNew As Database, rstTemp As Recordset
Dim Msg, Style, Title, Response, j As Long, k As Long, spawnComment(10) As String
Dim rstTemp2 As Recordset, m As Long, tempAll1 As String, tempAll2 As String
Dim temp2All1 As String, temp2All2 As String, pShared As Long, allNum As Long
Dim tblTemp As TableDef, rstTblTemp As Recordset, strTemp As String, strDrainage As String
Dim batchTally As Long, taLLY(8) As String, rstGenetics As Recordset
Dim femScored As Long, maleScored As Long, locScored As Long, x As Long, totAlleles As Long
'Dim empID(50) As Long, empName(50) As String
Dim q As Long, tmpStr As Variant, tmpChunk As Variant
Dim tmpDrain() As Variant, tmpCnt As Long, n As Long, p As Long
Dim strFactorial As String, tmpArray(11) As Long

Public Function removeComment()
    spawnComment(CInt(frmMain.txtCommentInt.Text)) = ""
    frmComment.txtComment.Text = spawnComment(CInt(frmMain.txtCommentInt.Text))
End Function

Public Function loadComment()
    frmComment.txtComment.Text = spawnComment(CInt(frmMain.txtCommentInt.Text))
End Function

Public Function addComment()
    spawnComment(CInt(frmMain.txtCommentInt.Text)) = frmComment.txtComment.Text
End Function

Private Function COMCheckPort(Port As Long) As Boolean

  'Returns false if port cannot be opened or does not exist
  'Returns true if port is available and can be opened
   
  'Handle all errors
   On Error GoTo OpenCom_Error
   
  'Set port number for test
   MSComm1.CommPort = Port
   
   If MSComm1.PortOpen = True Then
      COMCheckPort = False
      Exit Function
   Else
     'Test the port by opening and closing it
      MSComm1.PortOpen = True
      MSComm1.PortOpen = False
      COMCheckPort = True
      Exit Function
   End If
   
OpenCom_Error:
   COMCheckPort = False
   
End Function

Private Sub chkSpawned_Click(Index As Integer)
    If frmMain.chkSpawned(Index).Value = 1 Then
        If frmMain.txtFem(Index).BackColor = &HFFFF& Then 'yellow
            frmMain.txtFem(Index).BackColor = &H80FF&
        Else
            frmMain.txtFem(Index).BackColor = &HFF&
        End If
        
        If frmMain.txtMale(Index).BackColor = &HFFFF& Then 'yellow
            frmMain.txtMale(Index).BackColor = &H80FF& 'orange
        Else
            frmMain.txtMale(Index).BackColor = &HFF& 'red
        End If
        
        frmMain.txtPShare(Index).BackColor = &HFF&
    Else
        If frmMain.txtFem(Index).BackColor = &H80FF& Then 'orange
            frmMain.txtFem(Index).BackColor = &HFFFF& 'yellow
        Else
            frmMain.txtFem(Index).BackColor = &H80000005 'white
        End If

        If frmMain.txtMale(Index).BackColor = &H80FF& Then 'orange
            frmMain.txtMale(Index).BackColor = &HFFFF& 'yellow
        Else
            frmMain.txtMale(Index).BackColor = &H80000005 'white
        End If

        frmMain.txtPShare(Index).BackColor = &H80000005
    End If
End Sub

Private Sub cmdComment_Click(Index As Integer)
    frmMain.txtCommentInt.Text = Index
    frmComment.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload frmMain
End Sub

Private Sub cmdInputPIT_Click()
    '1 to 1
    If frmMain.optMatDes(0).Value = True And CInt(frmMain.cmbFem.Text) > CInt(frmMain.cmbMale.Text) Then
        Msg = "In a 1 to 1 mating design the number of males must be equal to or greater than the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If

    '1 to 2
    If frmMain.optMatDes(1).Value = True And CInt(frmMain.cmbFem.Text) > (CInt(frmMain.cmbMale.Text) / 2) Then
        Msg = "In a 1 to 2 mating design the number of males must be equal to or greater than twice the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If

    '2 to 2
    If frmMain.optMatDes(2).Value = True And ((CInt(frmMain.cmbFem.Text) > CInt(frmMain.cmbMale.Text)) Or (CInt(frmMain.cmbFem.Text) < 2) Or (CInt(frmMain.cmbFem.Text) Mod 2 = 1)) Then
        Msg = "In a 2 to 2 mating design there must an even number of females greater than or equal to 2, and the number of males must be equal to or greater than the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    '1 to 3
    If frmMain.optMatDes(3).Value = True And CInt(frmMain.cmbFem.Text) > (CInt(frmMain.cmbMale.Text) / 3) Then
        Msg = "In a 1 to 3 mating design the number of males must be equal to or greater than three times the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    '3 to 3
    If frmMain.optMatDes(4).Value = True And ((CInt(frmMain.cmbFem.Text) > CInt(frmMain.cmbMale.Text)) Or (CInt(frmMain.cmbFem.Text) < 3) Or (CInt(frmMain.cmbFem.Text) Mod 3 <> 0)) Then
        Msg = "In a 3 to 3 mating design there must a multiple of three females greater than or equal to 3, and the number of males must be equal to or greater than the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    'Can't reoptimize 2 to 2 or 3 to 3 mating designs
    q = 0
    For i = 0 To frmMain.chkSpawned.Count - 1
        If frmMain.chkSpawned(i).Value = 1 And frmMain.chkSpawned(i).Visible = True Then
            q = q + 1
        End If
    Next i
    
    If (frmMain.optMatDes(2).Value = True Or frmMain.optMatDes(4).Value = True) And q > 0 Then
        Msg = "Reoptimization is not possible for a 2 to 2 or 3 to 3 mating design when some matings have already occurred." & Chr(13) & Chr(13) & "Please check all completed or possible matings, change the mating design to 1 to 1, and replace the necessary individuals."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Reoptimation Not Possible"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    frmInput.Show 1
End Sub

Private Sub cmdRelQuarts_Click()
    Msg = "Do you wish to run the drainage specific spawning relatedness distributions?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Title = "Run Relatedness Distributions"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = 7 Then
        Exit Sub
    End If
    
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    With dbsNew
        Set rstTemp = dbsNew.OpenRecordset("SELECT Genetics.* From Genetics WHERE (((Genetics.Drainage) Is Not Null And (Genetics.Drainage)<>'Unknown') AND ((Genetics.Year)>=" & Year(Date) - 6 & " And (Genetics.Year)<=" & Year(Date) & ")) ORDER BY Genetics.Drainage, Genetics.Year", dbOpenDynaset)
        Set rstTemp2 = dbsNew.OpenRecordset("SELECT Genetics.* From Genetics WHERE (((Genetics.Drainage) Is Not Null And (Genetics.Drainage)<>'Unknown') AND ((Genetics.Year)>=" & Year(Date) - 6 & " And (Genetics.Year)<=" & Year(Date) & ")) ORDER BY Genetics.Drainage, Genetics.Year", dbOpenDynaset)

        'Set rstTemp = dbsNew.OpenRecordset("SELECT tblBrood.Drainage, Genetics.Year, Genetics.PIT, Genetics.[197 1], Genetics.[197 2], Genetics.[202 1], Genetics.[202 2], Genetics.[289 1], Genetics.[289 2], Genetics.[14 1], Genetics.[14 2], Genetics.[171 1], Genetics.[171 2], Genetics.[85 1], Genetics.[85 2], Genetics.[L82 1], Genetics.[L82 2], Genetics.[311 1], Genetics.[311 2], Genetics.[438 1], Genetics.[438 2], Genetics.[25 1], Genetics.[25 2], Genetics.[L85 1], Genetics.[L85 2] FROM Genetics LEFT JOIN tblBrood ON Genetics.PIT = tblBrood.Mark Where (((tblBrood.Drainage) > 0 And (tblBrood.Drainage) < 8) And ((Genetics.PIT) Is Not Null And (Genetics.PIT) <> '?' And (Genetics.PIT) <> 'notlegible')) ORDER BY tblBrood.Drainage, Genetics.Year", dbOpenDynaset)
        'Set rstTemp2 = dbsNew.OpenRecordset("SELECT tblBrood.Drainage, Genetics.Year, Genetics.PIT, Genetics.[197 1], Genetics.[197 2], Genetics.[202 1], Genetics.[202 2], Genetics.[289 1], Genetics.[289 2], Genetics.[14 1], Genetics.[14 2], Genetics.[171 1], Genetics.[171 2], Genetics.[85 1], Genetics.[85 2], Genetics.[L82 1], Genetics.[L82 2], Genetics.[311 1], Genetics.[311 2], Genetics.[438 1], Genetics.[438 2], Genetics.[25 1], Genetics.[25 2], Genetics.[L85 1], Genetics.[L85 2] FROM Genetics LEFT JOIN tblBrood ON Genetics.PIT = tblBrood.Mark Where (((tblBrood.Drainage) > 0 And (tblBrood.Drainage) < 8) And ((Genetics.PIT) Is Not Null And (Genetics.PIT) <> '?' And (Genetics.PIT) <> 'notlegible')) ORDER BY tblBrood.Drainage, Genetics.Year", dbOpenDynaset)
        
        On Error Resume Next
        dbsNew.TableDefs.Delete "Temp"
        On Error GoTo 0

        
        Set tblTemp = dbsNew.CreateTableDef("Temp")
        With tblTemp
            tblTemp.Fields.Append .CreateField("Drainage", dbText)
            tblTemp.Fields.Append .CreateField("PropShared", dbSingle)
            tblTemp.Fields.Append .CreateField("lnPropShared", dbSingle)
        End With
        
        dbsNew.TableDefs.Append tblTemp
        
        Set rstTblTemp = dbsNew.OpenRecordset("SELECT Temp.* FROM Temp", dbOpenDynaset)
        
        With rstTblTemp
            With rstTemp
                k = 1
                rstTemp.MoveFirst
                Do Until rstTemp.EOF
                    j = 0
                    For i = 7 To rstTemp.Fields.Count - 2
                        If IsNull(rstTemp.Fields(i).Value) = False Then
                            j = j + 1
                        End If
                    Next i
                
                    If j > 0 Then
                        With rstTemp2
                            rstTemp2.MoveFirst
                            If k < rstTemp2.RecordCount Then
                                rstTemp2.Move k
                            Else
                                Exit Do
                            End If
                            Do Until rstTemp2.EOF
                                If rstTemp!Drainage = rstTemp2!Drainage Then
                                    m = 0
                                    For i = 7 To rstTemp.Fields.Count - 2
                                        If IsNull(rstTemp2.Fields(i).Value) = False Then
                                            m = m + 1
                                        End If
                                    Next i
                            
                                    If m > 0 Then
                                        'calculate relatedness
                                        allNum = 0: pShared = 0
                                        For i = 7 To rstTemp.Fields.Count - 2
                                            If IsNull(rstTemp.Fields(i).Value) = False Then
                                                tempAll1 = rstTemp.Fields(i).Value
                                                i = i + 1
                                                tempAll2 = rstTemp.Fields(i).Value
                                                i = i - 1
                                                If IsNull(rstTemp2.Fields(i).Value) = False Then
                                                    temp2All1 = rstTemp2.Fields(i).Value
                                                    i = i + 1
                                                    temp2All2 = rstTemp2.Fields(i).Value
                                                
                                                    allNum = allNum + 2
                                                
                                                    If tempAll1 = tempAll2 Then
                                                        If tempAll1 = temp2All1 And tempAll1 = temp2All2 Then
                                                            pShared = pShared + 2
                                                        ElseIf tempAll1 = temp2All1 Or tempAll1 = temp2All2 Then
                                                            pShared = pShared + 1
                                                        End If
                                                    ElseIf temp2All1 = temp2All2 Then
                                                        If tempAll1 = temp2All1 And tempAll2 = temp2All1 Then
                                                            pShared = pShared + 2
                                                        ElseIf tempAll1 = temp2All1 Or tempAll2 = temp2All1 Then
                                                            pShared = pShared + 1
                                                        End If
                                                    Else
                                                        If tempAll1 = temp2All1 Or tempAll1 = temp2All2 Then
                                                            pShared = pShared + 1
                                                        End If
                        
                                                        If tempAll2 = temp2All1 Or tempAll2 = temp2All2 Then
                                                            pShared = pShared + 1
                                                        End If
                                                    End If
                                                Else
                                                    i = i + 1
                                                End If
                                            Else
                                                i = i + 1
                                            End If
                                        Next i
                                    
                                        rstTblTemp.AddNew
                                            rstTblTemp!Drainage = rstTemp!Drainage
                                            If allNum > 0 Then
                                                rstTblTemp!PropShared = Format(pShared / allNum, "0.000")
                                            End If
                                        rstTblTemp.Update
                                    End If
                                    rstTemp2.MoveNext
                                Else
                                    Exit Do
                                End If
                            Loop
                        End With
                    End If
                    k = k + 1
                    rstTemp.MoveNext
                Loop
            End With
        End With
        
        'My computer
        Open "C:\Maine Genetics Program\Spawning Form\Drainage Quartiles.txt" For Output As #1
        
        'Denise's computer
        'Open "C:\Databases\MaineBroodstock\SpawnOpt\Spawning Form\Drainage Quartiles.txt" For Output As #1

        
            Set rstTemp = dbsNew.OpenRecordset("SELECT Genetics.Drainage From Genetics GROUP BY Genetics.Drainage HAVING (((Genetics.Drainage) Is Not Null And (Genetics.Drainage)<>'Unknown'))", dbOpenDynaset)
            ReDim tmpDrain(rstTemp.RecordCount, 2)
            rstTemp.MoveFirst
            For i = 1 To rstTemp.RecordCount
                tmpDrain(i, 1) = rstTemp!Drainage
                Select Case UCase(rstTemp!Drainage)
                    Case "PENOBSCOT"
                        tmpDrain(i, 2) = 1
                    Case "NARRAGUAGUS"
                        tmpDrain(i, 2) = 3
                    Case "PLEASANT"
                        tmpDrain(i, 2) = 4
                    Case "MACHIAS"
                        tmpDrain(i, 2) = 5
                    Case "EAST MACHIAS"
                        tmpDrain(i, 2) = 6
                    Case "DENNYS"
                        tmpDrain(i, 2) = 7
                    Case "SHEEPSCOT"
                        tmpDrain(i, 2) = 11
                End Select
                rstTemp.MoveNext
            Next i
            
            tmpCnt = rstTemp.RecordCount
            
            For i = 1 To tmpCnt
                Set rstTemp = dbsNew.OpenRecordset("SELECT Temp.* From Temp Where (((Temp.Drainage) = '" & tmpDrain(i, 1) & "')) ORDER BY Temp.PropShared", dbOpenDynaset)
                    
                With rstTemp
                    If rstTemp.RecordCount > 0 Then
                        rstTemp.MoveFirst
                        rstTemp.Move (Int(rstTemp.RecordCount * 0.75))
                        Print #1, tmpDrain(i, 2) & Chr(9) & rstTemp!PropShared
                    Else
                        Print #1, tmpDrain(i, 2) & Chr(9) & "0.75"
                    End If
                End With
            Next i
            
            Set rstTemp = Nothing
            Set rstTemp2 = Nothing
            Set rstTblTemp = Nothing
            
            dbsNew.TableDefs.Delete "Temp"
        Close #1
        
    End With
    
    dbsNew.Close
    wrkJet.Close
    
    Msg = "The new quartile values have been calculated for all drainages."
    Style = vbOKOnly + vbInformation + vbDefaultButton1
    Title = "Quartiles Calculated"
    Response = MsgBox(Msg, Style, Title)
    
End Sub

Private Sub cmdSpawnUp_Click()
    'If frmMain.cmbEmployee.Text = "" Then
    '    Msg = "Please enter your name."
    '    Style = vbOKOnly + vbInformation + vbDefaultButton1
    '    Title = "Unknown Identity"
    '    Response = MsgBox(Msg, Style, Title)
        
    '    frmMain.cmbEmployee.SetFocus
    '    Exit Sub
    'End If
        
    j = 0
    k = 0
    For i = 0 To 9
        If frmMain.txtFem(i).Visible = True Then j = j + 1
        If frmMain.chkSpawned(i).Value = 1 Then k = k + 1
    Next i
    
    If j = 0 Or k = 0 Then
        Msg = "There are no matings selected to update."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "No Matings"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    If k < j Then
        Msg = "There are " & j - k & " matings that aren't selected." & Chr(13) & Chr(13) & "Do you wish to continue without selecting those matings?" & Chr(13) & "Warning: All data will be removed after updating 'tblBroodMating'!"
        Style = vbYesNo + vbCritical + vbDefaultButton1
        Title = "Selected Matings"
        Response = MsgBox(Msg, Style, Title)
        
        If Response = 7 Then
            Exit Sub
        End If
    End If
    
    'My computer
    Open "C:\Maine Genetics Program\Spawning Form\Batch Tally.txt" For Input As #1
    
    'Denise's computer
    'Open "C:\Databases\MaineBroodstock\SpawnOpt\Spawning Form\Batch Tally.txt" For Input As #1
    
    j = 1
    Do Until EOF(1)
        Input #1, strTemp
        If CInt(Left(strTemp, 2)) = CInt(txtDrainageF(0)) Then
            If Mid(strTemp, 3, 5) = Format(Date, "YYYY") Then
                i = 8
                Do Until Mid(strTemp, i, 1) = ""
                    i = i + 1
                Loop
                batchTally = CInt(Mid(strTemp, 8, i - 8)) + 1
                taLLY(j) = Left(strTemp, 7) & batchTally
                GoTo skipTally
            Else
                batchTally = 1
                taLLY(j) = Left(strTemp, 2) & Format(Date, "YYYY") & " " & batchTally
                GoTo skipTally
            End If
        End If
        taLLY(j) = strTemp
skipTally:
        j = j + 1
    Loop
    
    Close #1
    
    'My computer
    Open "C:\Maine Genetics Program\Spawning Form\Batch Tally.txt" For Output As #1
    
    'Denise's computer
    'Open "C:\Databases\MaineBroodstock\SpawnOpt\Spawning Form\Batch Tally.txt" For Output As #1
    
    For i = 1 To 8
        Print #1, taLLY(i)
    Next i
    
    Close #1
        
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    With dbsNew
        Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMating.* FROM tblBroodMating", dbOpenDynaset)
        With rstTemp
            For i = 0 To j
                If frmMain.chkSpawned(i).Value = 1 Then
                    rstTemp.AddNew
                    Select Case CInt(frmMain.txtDrainageF(i).Text)
                        Case 1
                            !Family = "PN" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "PE"
                        Case 11
                            !Family = "SHP" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "SH"
                        Case 3
                            !Family = "NG" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "NA"
                        Case 4
                            !Family = "PL" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "PL"
                        Case 5
                            !Family = "MC" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "MA"
                        Case 6
                            !Family = "EM" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "EM"
                        Case 7
                            !Family = "DE" & Format(Date, "YYYY") & "A" & frmMain.lblFem(i).Caption
                            !HatcheryStandardCode = "DE"
                    End Select
                    
                    !FemaleCaptureYear = CInt(frmMain.txtfYear(i).Text)
                    !MaleCaptureYear = CInt(frmMain.txtmYear(i).Text)
                    !GreenEggBatch = Null
                    !Dam = frmMain.txtFem(i).Text
                    !Sire = frmMain.txtMale(i).Text
                    !DamScoredLoci = frmMain.txtLocFem(i).Text
                    !SireScoredLoci = frmMain.txtLocMale(i).Text
                    !PropSharedAlleles = CSng(frmMain.txtPShare(i).Text)
                    !TakeDate = Date
                    !TakeTime = Time
                    !Dispositonid = 1
                    !Optimized = frmMain.txtOptimized(i).Text
                    !Batch = CInt(batchTally)
                    If spawnComment(i) <> "" Then
                        !Comments = spawnComment(i)
                    End If
                    !ReleasedF = frmMain.chkReleaseF(i).Value
                    !ReleasedM = frmMain.chkReleaseM(i).Value
                    'If frmMain.txtPicFem(i).Text <> "" Then
                    '    !FemalePicNum = frmMain.txtPicFem(i).Text
                    'End If
                    '!FemaleWeight = CSng(frmMain.txtWeightFem(i).Text)
                    'If frmMain.txtPicMale(i).Text <> "" Then
                    '    !MalePicNum = frmMain.txtPicMale(i).Text
                    'End If
                    '!MaleWeight = CSng(frmMain.txtWeightMale(i).Text)
                    
                    'q = 0
                    'Do Until empName(q) = frmMain.cmbEmployee.Text
                    '    q = q + 1
                    'Loop
                    '!TableEditor = empID(q)
                    
                    '!TableEditDate = Date
                    
                    On Error GoTo describeErr
                    rstTemp.Update
                End If
            Next i
        End With
        
        x = 0
        For i = 0 To 9
            If frmInput.txtTagMale(i).Text <> "" Then x = x + 1
        Next i
        
        Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMatingExtras.* FROM tblBroodMatingExtras", dbOpenDynaset)
        With rstTemp
            For i = 0 To x - 1
                rstTemp.AddNew
                    
                    Select Case CInt(frmMain.txtDrainageF(0).Text)
                        Case 1
                            !Drainage = "Penobscot"
                        Case 11
                            !Drainage = "Sheepscot"
                        Case 3
                            !Drainage = "Narraguagas"
                        Case 4
                            !Drainage = "Pleasant"
                        Case 5
                            !Drainage = "Machias"
                        Case 6
                            !Drainage = "East Machias"
                        Case 7
                            !Drainage = "Dennys"
                    End Select
                    
                    If frmInput.txtTagFem(i).Text <> "" Then
                        !Female = frmInput.txtTagFem(i).Text
                    End If
                    !Male = frmInput.txtTagMale(i).Text
                
                    Set rstGenetics = dbsNew.OpenRecordset("SELECT Genetics.* FROM Genetics", dbOpenDynaset)
                    
                    ReDim fAlleles(1, rstGenetics.Fields.Count - 8, 2)
                    ReDim mAlleles(1, rstGenetics.Fields.Count - 8, 2)
                    
                    With rstGenetics
                        rstGenetics.MoveFirst
                        Do Until rstGenetics.EOF
                            If UCase(rstGenetics![pit]) = UCase(frmMain.txtFem(i).Text) Then
                                m = 1
                                locScored = 0
                                For k = 7 To rstGenetics.Fields.Count - 2
                                    fAlleles(1, m, 1) = rstGenetics.Fields(k).Value
                                    If IsNull(rstGenetics.Fields(k).Value) = False Then locScored = locScored + 1
                                    k = k + 1
                                    fAlleles(1, m, 2) = rstGenetics.Fields(k).Value
                                    If IsNull(rstGenetics.Fields(k).Value) = False Then locScored = locScored + 1
                                    m = m + 1
                                Next k
                                    femScored = locScored / 2
                            End If
                            
                            If UCase(rstGenetics![pit]) = UCase(frmInput.txtTagMale(i).Text) Then
                                m = 1
                                locScored = 0
                                For k = 7 To rstGenetics.Fields.Count - 2
                                    mAlleles(1, m, 1) = rstGenetics.Fields(k).Value
                                    If IsNull(rstGenetics.Fields(k).Value) = False Then locScored = locScored + 1
                                    k = k + 1
                                    mAlleles(1, m, 2) = rstGenetics.Fields(k).Value
                                    If IsNull(rstGenetics.Fields(k).Value) = False Then locScored = locScored + 1
                                    m = m + 1
                                Next k
                                maleScored = locScored / 2
                            End If
                            rstGenetics.MoveNext
                        Loop
                        
                        totAlleles = rstGenetics.Fields.Count - 8: pShared = 0
                        For j = 1 To totAlleles / 2
                            If IsNull(fAlleles(1, j, 1)) = True Or IsNull(fAlleles(1, j, 2)) = True Or IsNull(mAlleles(1, j, 1)) = True Or IsNull(mAlleles(1, j, 2)) = True Then
                                totAlleles = totAlleles - 2
                                GoTo nextLoci
                            End If
                    
                            If fAlleles(1, j, 1) = fAlleles(1, j, 2) Then
                                If fAlleles(1, j, 1) = mAlleles(1, j, 1) And fAlleles(1, j, 1) = mAlleles(1, j, 2) Then
                                    pShared = pShared + 2
                                ElseIf fAlleles(1, j, 1) = mAlleles(1, j, 1) Or fAlleles(1, j, 1) = mAlleles(1, j, 2) Then
                                    pShared = pShared + 1
                                End If
                            ElseIf mAlleles(1, j, 1) = mAlleles(1, j, 2) Then
                                If fAlleles(1, j, 1) = mAlleles(1, j, 1) And fAlleles(1, j, 2) = mAlleles(1, j, 1) Then
                                    pShared = pShared + 2
                                ElseIf fAlleles(1, j, 1) = mAlleles(1, j, 1) Or fAlleles(1, j, 2) = mAlleles(1, j, 1) Then
                                    pShared = pShared + 1
                                End If
                            Else
                                If fAlleles(1, j, 1) = mAlleles(1, j, 1) Or fAlleles(1, j, 1) = mAlleles(1, j, 2) Then
                                    pShared = pShared + 1
                                End If
                        
                                If fAlleles(1, j, 2) = mAlleles(1, j, 1) Or fAlleles(1, j, 2) = mAlleles(1, j, 2) Then
                                    pShared = pShared + 1
                                End If
                            End If
nextLoci:
                        Next j
                
                        If femScored = 0 Then
                            rstTemp!PropSharedAlleles = 1
                            rstTemp!Comments = "This female was not scored at any loci"
                        ElseIf maleScored = 0 Then
                            rstTemp!PropSharedAlleles = 1
                            If rstTemp!Comments = "" Then
                                rstTemp!Comments = "This male was not scored at any loci"
                            Else
                                rstTemp!Comments = rstTemp!Comments & "\This male was not scored at any loci"
                            End If
                        Else
                            rstTemp!PropSharedAlleles = Format((pShared / totAlleles), "0.000")
                        End If
                    End With
                    !Batch = CInt(batchTally)
                    !Date = Date
                    !Time = Format(Time, "HH:MM")
                    !Comments = frmMain.txtFlag(i).Text
                rstTemp.Update
            Next i
        End With
    End With
    
    dbsNew.Close
    wrkJet.Close
    
    x = 0
    For i = 0 To 9
        If frmMain.txtFem(i).Visible = True Then x = x + 1
    Next i
    
    For i = 0 To x - 1
        frmMain.lblFem(i).Visible = False
        frmMain.txtFem(i).Text = ""
        frmMain.txtFem(i).Visible = False
        frmMain.txtLocFem(i).Text = ""
        frmMain.txtLocFem(i).Visible = False
        frmMain.lblMale(i).Caption = i + 1
        frmMain.lblMale(i).Visible = False
        frmMain.txtMale(i).Text = ""
        frmMain.txtMale(i).Visible = False
        frmMain.txtLocMale(i).Text = ""
        frmMain.txtLocMale(i).Visible = False
        frmMain.txtPShare(i).Text = ""
        frmMain.txtPShare(i).Visible = False
        frmMain.txtOptimized(i).Text = ""
        frmMain.imgFlag(i).Visible = False
        frmMain.txtFlag(i).Text = ""
        frmMain.cmdComment(i).Visible = False
        frmMain.chkSpawned(i).Value = 0
        frmMain.chkSpawned(i).Visible = False
        frmMain.chkReleaseF(i).Value = 0
        frmMain.chkReleaseF(i).Visible = False
        frmMain.chkReleaseM(i).Value = 0
        frmMain.chkReleaseM(i).Visible = False
    Next i
    
    For i = 0 To 9
        spawnComment(i) = ""
    Next i
    
    Unload frmInput
    
    Msg = "The table 'tblBroodMating' has been updated."
    Style = vbOKOnly + vbInformation + vbDefaultButton1
    Title = "Update Complete"
    Response = MsgBox(Msg, Style, Title)
    
Exit Sub

describeErr:
    If Err.Number <> 0 Then
        Msg = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    Exit Sub
End Sub



Private Sub databaseSettings_Click(Index As Integer)
    frmDataSpec.Show 1
End Sub

Public Sub Form_Load()
    Dim Port As Long

    For Port = 1 To 16
      If COMCheckPort(Port) Then
         frmMain.MSComm1.CommPort = Port
         frmMain.MSComm1.Settings = "9600,N,8,1"
         frmMain.MSComm1.PortOpen = True
         frmMain.MSComm1.InputLen = 0
         frmMain.MSComm1.RThreshold = 1
         frmMain.MSComm1.EOFEnable = True
         Exit For
      End If
    Next

    frmMain.Show 0
    
    frmMain.cmbFem.Text = 5
    frmMain.cmbMale.Text = 5
    frmMain.optMatDes(0).Value = True
      
    'frmMain.txtPitCur.LinkMode = 1
    
    'Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    'Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    'With dbsNew
    '    Set rstTemp = dbsNew.OpenRecordset("SELECT CodeEmployee.* From CodeEmployee WHERE (((CodeEmployee.LastName)='King')) OR (((CodeEmployee.LastName)='Craig')) OR (((CodeEmployee.LastName)='Buckley')) OR (((CodeEmployee.LastName)='Tozier')) OR (((CodeEmployee.LastName)='Thies'))", dbOpenDynaset)
    '    With rstTemp
    '        i = 0
    '        rstTemp.MoveFirst
    '        Do Until rstTemp.EOF
    '            frmMain.cmbEmployee.AddItem !LastName & ", " & !FirstName
    '            empID(i) = rstTemp!EmployeeID
    '            empName(i) = !LastName & ", " & !FirstName
    '            i = i + 1
    '            rstTemp.MoveNext
    '        Loop
    '    End With
    'End With
    
    'dbsNew.Close
    'wrkJet.Close
       
    'my computer
    Open "C:\Maine Genetics Program\Spawning Form\Drainage Quartiles.txt" For Input As #1
    
    'Denise's Computer
    'Open "C:\Databases\MaineBroodstock\SpawnOpt\Spawning Form\Drainage Quartiles.txt" For Input As #1
    
    m = 0
    Do Until EOF(1)
        Input #1, strTemp
        i = 1
        Do Until Mid(strTemp, i, 1) = Chr(9)
            i = i + 1
        Loop
        strDrainage = Mid(strTemp, 1, i - 1)
        
        j = i
        Do Until Mid(strTemp, j, 1) = ""
            j = j + 1
        Loop
        frmMain.txtQuartile(m).Text = Mid(strTemp, i + 1, j - i)
        m = m + 1
    Loop
    Close #1
    
    frmMain.txtQuartile(7) = "0.75"
    
    
    'Call getList1
    'Call getList2
    'Call getList3
    'Call getList4
    'Call getList5
    'Call getList6
    'Call getList7
    'Call getList8
    'Call getList9
    'Call getList10

    'Open "C:\Maine Genetics Program\Factorial Combinations\10Factorial.txt" For Input As #1
    'j = 0
    'Do Until EOF(1)
    '    m = 1:  n = 0: p = 1
    '    'acquiring male order
    '    Input #1, strFactorial
    '    j = j + 1
    '    Do Until Mid(strFactorial, m, 1) = ""
    '        If Mid(strFactorial, m, 1) = " " Or Mid(strFactorial, m, 1) = ";" Then
    '            tmpArray(p) = CLng(Mid(strFactorial, m - n, n))
    '            p = p + 1
    '            n = -1
    '        End If
    '        m = m + 1
    '        n = n + 1
    '    Loop
    '    mList10.Add tmpArray
    'Loop
    'Close #1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Msg = "Do you wish to exit the spawning optimization software?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Title = "Exit Software"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = 7 Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub imgFlag_Click(Index As Integer)
    Msg = frmMain.txtFlag(Index).Text
    Style = vbOKOnly + vbCritical + vbDefaultButton1
    Title = "Mating Concerns"
    Response = MsgBox(Msg, Style, Title)
End Sub

Private Sub lblMatDes_Click(Index As Integer)
    frmMain.optMatDes(Index).Value = True
End Sub

Private Sub optMatDes_Click(Index As Integer)
    Select Case Index
        Case 0 '1 to 1
            frmMain.cmbFem.Clear
            frmMain.cmbFem.AddItem ("1")
            frmMain.cmbFem.AddItem ("2")
            frmMain.cmbFem.AddItem ("3")
            frmMain.cmbFem.AddItem ("4")
            frmMain.cmbFem.AddItem ("5")
            frmMain.cmbFem.AddItem ("6")
            frmMain.cmbFem.AddItem ("7")
            frmMain.cmbFem.AddItem ("8")
            frmMain.cmbFem.AddItem ("9")
            frmMain.cmbFem.AddItem ("10")
            
            frmMain.cmbMale.Clear
            frmMain.cmbMale.AddItem ("1")
            frmMain.cmbMale.AddItem ("2")
            frmMain.cmbMale.AddItem ("3")
            frmMain.cmbMale.AddItem ("4")
            frmMain.cmbMale.AddItem ("5")
            frmMain.cmbMale.AddItem ("6")
            frmMain.cmbMale.AddItem ("7")
            frmMain.cmbMale.AddItem ("8")
            frmMain.cmbMale.AddItem ("9")
            frmMain.cmbMale.AddItem ("10")
            'frmMain.cmbMale.AddItem ("11")
            
            frmMain.cmbFem.Text = 5
            frmMain.cmbMale.Text = 5
            frmInput.lblMD.Caption = "1 to 1 Mating Design"
        Case 1 '1 to 2
            frmMain.cmbFem.Clear
            frmMain.cmbFem.AddItem ("1")
            frmMain.cmbFem.AddItem ("2")
            frmMain.cmbFem.AddItem ("3")
            frmMain.cmbFem.AddItem ("4")
            frmMain.cmbFem.AddItem ("5")
            
            frmMain.cmbMale.Clear
            frmMain.cmbMale.AddItem ("1")
            frmMain.cmbMale.AddItem ("2")
            frmMain.cmbMale.AddItem ("3")
            frmMain.cmbMale.AddItem ("4")
            frmMain.cmbMale.AddItem ("5")
            frmMain.cmbMale.AddItem ("6")
            frmMain.cmbMale.AddItem ("7")
            frmMain.cmbMale.AddItem ("8")
            frmMain.cmbMale.AddItem ("9")
            frmMain.cmbMale.AddItem ("10")
            'frmMain.cmbMale.AddItem ("11")
            
            frmMain.cmbFem.Text = 5
            frmMain.cmbMale.Text = 10
            frmInput.lblMD.Caption = "1 to 2 Mating Design"
        Case 2 '2 to 2
            frmMain.cmbFem.Clear
            frmMain.cmbFem.AddItem ("2")
            frmMain.cmbFem.AddItem ("4")
            
            frmMain.cmbMale.Clear
            frmMain.cmbMale.AddItem ("2")
            frmMain.cmbMale.AddItem ("3")
            frmMain.cmbMale.AddItem ("4")
            frmMain.cmbMale.AddItem ("5")
            frmMain.cmbMale.AddItem ("6")
            frmMain.cmbMale.AddItem ("7")
            frmMain.cmbMale.AddItem ("8")
            frmMain.cmbMale.AddItem ("9")
            frmMain.cmbMale.AddItem ("10")
            'frmMain.cmbMale.AddItem ("11")
            
            frmMain.cmbFem.Text = 4
            frmMain.cmbMale.Text = 4
            frmInput.lblMD.Caption = "2 to 2 Mating Design"
        Case 3 '1 to 3
            frmMain.cmbFem.Clear
            frmMain.cmbFem.AddItem ("1")
            frmMain.cmbFem.AddItem ("2")
            frmMain.cmbFem.AddItem ("3")
            
            frmMain.cmbMale.Clear
            frmMain.cmbMale.AddItem ("3")
            frmMain.cmbMale.AddItem ("4")
            frmMain.cmbMale.AddItem ("5")
            frmMain.cmbMale.AddItem ("6")
            frmMain.cmbMale.AddItem ("7")
            frmMain.cmbMale.AddItem ("8")
            frmMain.cmbMale.AddItem ("9")
            frmMain.cmbMale.AddItem ("10")
            'frmMain.cmbMale.AddItem ("11")
            
            frmMain.cmbFem.Text = 3
            frmMain.cmbMale.Text = 9
            frmInput.lblMD.Caption = "1 to 3 Mating Design"
        Case 4 '3 to 3
            frmMain.cmbFem.Clear
            frmMain.cmbFem.AddItem ("3")
            
            frmMain.cmbMale.Clear
            frmMain.cmbMale.AddItem ("3")
            frmMain.cmbMale.AddItem ("4")
            frmMain.cmbMale.AddItem ("5")
            frmMain.cmbMale.AddItem ("6")
            frmMain.cmbMale.AddItem ("7")
            frmMain.cmbMale.AddItem ("8")
            frmMain.cmbMale.AddItem ("9")
            frmMain.cmbMale.AddItem ("10")
            'frmMain.cmbMale.AddItem ("11")
            
            frmMain.cmbFem.Text = 3
            frmMain.cmbMale.Text = 3
            frmInput.lblMD.Caption = "3 to 3 Mating Design"
    End Select
End Sub

Private Sub MSComm1_OnComm()
    If frmMain.MSComm1.CommEvent = comEvReceive Then
        tmpChunk = frmMain.MSComm1.Input
        tmpStr = tmpStr & tmpChunk
        'Debug.Print (tmpStr)
        For i = 1 To Len(tmpStr) - 1
            If (Mid(tmpStr, i, 1)) = "." And (i + 10) <= Len(tmpStr) Then
                If frmInput.Visible = True Then
                    frmInput.txtTagCur.Text = Mid(tmpStr, i - 3, 14)
                    tmpStr = ""
                    Exit For
                ElseIf frmMain.Visible = True Then
                    frmMain.txtPitCur.Text = Mid(tmpStr, i - 3, 14)
                    tmpStr = ""
                    Exit For
                End If
            End If
        Next i
    End If
End Sub

Public Function highlightTag()
        j = -1
        For i = 0 To frmMain.txtFem.Count - 1
            If frmMain.txtFem(i).Visible = True Then
                j = j + 1
            End If
        Next i
    
        For i = 0 To j
            If frmMain.txtFem(i).Text = frmMain.txtPitCur2.Text Then
                If frmMain.txtFem(i).BackColor = &H80000005 Then 'white
                    frmMain.txtFem(i).BackColor = &HFFFF& 'yellow
                Else
                    frmMain.txtFem(i).BackColor = &H80FF& 'orange
                End If
            Else
                If frmMain.txtFem(i).BackColor = &H80FF& Or frmMain.txtFem(i).BackColor = &HFF& Then 'orange or red
                    frmMain.txtFem(i).BackColor = &HFF&       'red
                Else
                    frmMain.txtFem(i).BackColor = &H80000005 'white
                End If
            End If
        
            If frmMain.txtMale(i).Text = frmMain.txtPitCur2.Text Then
                If frmMain.txtMale(i).BackColor = &H80000005 Then 'white
                    frmMain.txtMale(i).BackColor = &HFFFF& 'yellow
                Else
                    frmMain.txtMale(i).BackColor = &H80FF& 'orange
                End If
            Else
                If frmMain.txtMale(i).BackColor = &H80FF& Or frmMain.txtMale(i).BackColor = &HFF& Then  'orange or red
                    frmMain.txtMale(i).BackColor = &HFF&       'red
                Else
                    frmMain.txtMale(i).BackColor = &H80000005 'white
                End If
            End If
        Next i
End Function

Private Sub txtPitCur_Change()
    If frmMain.txtPitCur.Text <> " No ID Found" And frmMain.txtPitCur.Text <> " LOOKING" And frmMain.txtPitCur.Text <> " Low Battery" And frmMain.txtPitCur.Text <> "AVID/FECAVA/ISO" Then
        If frmMain.chkPrefix.Value = 1 Then
            frmMain.txtPitCur2.Text = frmMain.txtPitCur.Text
        Else
            frmMain.txtPitCur2.Text = Right(frmMain.txtPitCur.Text, 10)
        End If
    Else
        frmMain.txtPitCur2.Text = ""
    End If
End Sub

Private Sub txtPitCur2_Change()
    k = 1
    Do Until Mid(frmMain.txtPitCur2.Text, k, 1) = ""
        k = k + 1
    Loop
    
    frmMain.txtPitCur2.Text = UCase(frmMain.txtPitCur2.Text)
    
    'If frmMain.txtPitCur2.Text <> "" And (k = 11 Or Mid(frmMain.txtPitCur2.Text, 4, 1) = ".") Then
        Call highlightTag
    'End If
    frmMain.txtPitCur2.SelStart = k - 1
    
End Sub
