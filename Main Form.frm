VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmMain 
   Caption         =   "Mate Matcher: Mating Optimization Software"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   16410
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   16410
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRelMetric 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Main Form.frx":048A
      Left            =   360
      List            =   "Main Form.frx":0491
      MouseIcon       =   "Main Form.frx":04B3
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   255
      ToolTipText     =   "The metric that will be used to calculate pairwise relatedness between males and females"
      Top             =   7440
      Width           =   3615
   End
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
      TabIndex        =   249
      TabStop         =   0   'False
      ToolTipText     =   "Mean genetic relatedness of all mating pairs (0 = low, 1 = high)"
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   9
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0605
      MousePointer    =   99  'Custom
      TabIndex        =   48
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   8
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0757
      MousePointer    =   99  'Custom
      TabIndex        =   44
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   7
      Left            =   15960
      MouseIcon       =   "Main Form.frx":08A9
      MousePointer    =   99  'Custom
      TabIndex        =   40
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   6
      Left            =   15960
      MouseIcon       =   "Main Form.frx":09FB
      MousePointer    =   99  'Custom
      TabIndex        =   36
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   5
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0B4D
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   4
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0C9F
      MousePointer    =   99  'Custom
      TabIndex        =   28
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   3
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0DF1
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   2
      Left            =   15960
      MouseIcon       =   "Main Form.frx":0F43
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   1
      Left            =   15960
      MouseIcon       =   "Main Form.frx":1095
      MousePointer    =   99  'Custom
      TabIndex        =   16
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   9
      Left            =   15240
      MouseIcon       =   "Main Form.frx":11E7
      MousePointer    =   99  'Custom
      TabIndex        =   47
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   8
      Left            =   15240
      MouseIcon       =   "Main Form.frx":1339
      MousePointer    =   99  'Custom
      TabIndex        =   43
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   7
      Left            =   15240
      MouseIcon       =   "Main Form.frx":148B
      MousePointer    =   99  'Custom
      TabIndex        =   39
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   6
      Left            =   15240
      MouseIcon       =   "Main Form.frx":15DD
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   5
      Left            =   15240
      MouseIcon       =   "Main Form.frx":172F
      MousePointer    =   99  'Custom
      TabIndex        =   31
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   4
      Left            =   15240
      MouseIcon       =   "Main Form.frx":1881
      MousePointer    =   99  'Custom
      TabIndex        =   27
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   3
      Left            =   15240
      MouseIcon       =   "Main Form.frx":19D3
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   2
      Left            =   15240
      MouseIcon       =   "Main Form.frx":1B25
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   1
      Left            =   15240
      MouseIcon       =   "Main Form.frx":1C77
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseM 
      Height          =   375
      Index           =   0
      Left            =   15960
      MouseIcon       =   "Main Form.frx":1DC9
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Check box to indicate male was released into the wild post-mating"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkReleaseF 
      Height          =   375
      Index           =   0
      Left            =   15240
      MouseIcon       =   "Main Form.frx":1F1B
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Check box to indicate female was released into the wild post-mating"
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
      Caption         =   "Include Prefix (PIT Tags)"
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
      Left            =   11280
      MouseIcon       =   "Main Form.frx":206D
      MousePointer    =   99  'Custom
      TabIndex        =   244
      ToolTipText     =   "Include the prefix for the PIT tag identifier"
      Top             =   360
      Value           =   1  'Checked
      Width           =   2415
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
      MouseIcon       =   "Main Form.frx":21BF
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2311
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2463
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":25B5
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2707
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2859
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":29AB
      MousePointer    =   99  'Custom
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
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2AFD
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "Main Form.frx":2C4F
      MousePointer    =   99  'Custom
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
      Left            =   9120
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of most recently scanned individual"
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
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   101
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   97
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   96
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   94
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
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
      TabIndex        =   93
      TabStop         =   0   'False
      ToolTipText     =   "Genetic relatedness between the mating pair (0 = low, 1 = high)"
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSpawnUp 
      BackColor       =   &H0080FF80&
      Caption         =   "Update Matings Table"
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
      MouseIcon       =   "Main Form.frx":2DA1
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Append the highlighted mated pairs to the matings table"
      Top             =   8360
      Width           =   2775
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   9
      Left            =   14160
      MouseIcon       =   "Main Form.frx":2EF3
      MousePointer    =   99  'Custom
      TabIndex        =   46
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   8
      Left            =   14160
      MouseIcon       =   "Main Form.frx":3045
      MousePointer    =   99  'Custom
      TabIndex        =   42
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   7
      Left            =   14160
      MouseIcon       =   "Main Form.frx":3197
      MousePointer    =   99  'Custom
      TabIndex        =   38
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   6
      Left            =   14160
      MouseIcon       =   "Main Form.frx":32E9
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   5
      Left            =   14160
      MouseIcon       =   "Main Form.frx":343B
      MousePointer    =   99  'Custom
      TabIndex        =   30
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   4
      Left            =   14160
      MouseIcon       =   "Main Form.frx":358D
      MousePointer    =   99  'Custom
      TabIndex        =   26
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   3
      Left            =   14160
      MouseIcon       =   "Main Form.frx":36DF
      MousePointer    =   99  'Custom
      TabIndex        =   22
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   2
      Left            =   14160
      MouseIcon       =   "Main Form.frx":3831
      MousePointer    =   99  'Custom
      TabIndex        =   18
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   1
      Left            =   14160
      MouseIcon       =   "Main Form.frx":3983
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSpawned 
      Height          =   375
      Index           =   0
      Left            =   14160
      MouseIcon       =   "Main Form.frx":3AD5
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Check box to highlight pairs that have already been mated"
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
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   87
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   85
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating male"
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
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
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
      TabIndex        =   69
      TabStop         =   0   'False
      ToolTipText     =   "Unique identifier of mating female"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdInputPIT 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Mating Individuals"
      Enabled         =   0   'False
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
      Left            =   765
      MouseIcon       =   "Main Form.frx":3C27
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Opens form to add individuals for potential mating and view information about that individual"
      Top             =   8160
      Width           =   2775
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
      Left            =   1245
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   360
         MouseIcon       =   "Main Form.frx":3D79
         MousePointer    =   99  'Custom
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "3 females to 3 males mating design"
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   360
         MouseIcon       =   "Main Form.frx":3ECB
         MousePointer    =   99  'Custom
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "1 female to 3 males mating design"
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   360
         MouseIcon       =   "Main Form.frx":401D
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "2 females to 2 males mating design"
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   360
         MouseIcon       =   "Main Form.frx":416F
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "1 female to 2 males mating design"
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton optMatDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   360
         MouseIcon       =   "Main Form.frx":42C1
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "1 female to 1 male mating design"
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
         Left            =   720
         TabIndex        =   56
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
         Left            =   720
         TabIndex        =   55
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
         Left            =   720
         TabIndex        =   54
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
         Left            =   720
         TabIndex        =   53
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
         Left            =   720
         TabIndex        =   52
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
      ItemData        =   "Main Form.frx":4413
      Left            =   2760
      List            =   "Main Form.frx":4438
      MouseIcon       =   "Main Form.frx":445F
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Number of males to add to candidate pool for optimization"
      Top             =   3840
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
      ItemData        =   "Main Form.frx":45B1
      Left            =   2760
      List            =   "Main Form.frx":45D3
      MouseIcon       =   "Main Form.frx":45F6
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Number of males to add to candidate pool for optimization"
      Top             =   3240
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
      Left            =   9120
      LinkItem        =   "Field(1)"
      LinkTopic       =   "winwedge|COM1"
      TabIndex        =   51
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
      Left            =   5280
      MouseIcon       =   "Main Form.frx":4748
      MousePointer    =   99  'Custom
      TabIndex        =   50
      ToolTipText     =   "Exit out of Mate Matcher"
      Top             =   8360
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
      TabIndex        =   114
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
      TabIndex        =   115
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
      TabIndex        =   116
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
      TabIndex        =   117
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
      TabIndex        =   118
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
      TabIndex        =   119
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
      TabIndex        =   120
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
      TabIndex        =   121
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
      TabIndex        =   122
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
      TabIndex        =   123
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
      TabIndex        =   124
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
      TabIndex        =   125
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
      TabIndex        =   126
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
      TabIndex        =   127
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
      TabIndex        =   128
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
      TabIndex        =   129
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
      TabIndex        =   130
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
      TabIndex        =   131
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
      TabIndex        =   132
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
      TabIndex        =   133
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
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   0
      Left            =   12360
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOptimized 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   136
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
      TabIndex        =   137
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
      TabIndex        =   138
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
      TabIndex        =   139
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
      TabIndex        =   140
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
      TabIndex        =   141
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
      TabIndex        =   142
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
      TabIndex        =   143
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
      TabIndex        =   144
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
      TabIndex        =   145
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
      TabIndex        =   146
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
      TabIndex        =   147
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
      TabIndex        =   148
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
      TabIndex        =   149
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
      TabIndex        =   150
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
      TabIndex        =   151
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
      TabIndex        =   152
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
      TabIndex        =   153
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
      TabIndex        =   154
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
      TabIndex        =   155
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
      TabIndex        =   156
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
      TabIndex        =   157
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
      TabIndex        =   158
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
      TabIndex        =   159
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
      TabIndex        =   160
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
      TabIndex        =   161
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
      TabIndex        =   162
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
      TabIndex        =   163
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
      TabIndex        =   164
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
      TabIndex        =   165
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
      TabIndex        =   166
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
      TabIndex        =   167
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
      TabIndex        =   168
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
      TabIndex        =   169
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
      TabIndex        =   170
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
      TabIndex        =   171
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
      TabIndex        =   172
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
      TabIndex        =   173
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
      TabIndex        =   174
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
      TabIndex        =   175
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
      TabIndex        =   176
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
      TabIndex        =   177
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
      TabIndex        =   178
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
      TabIndex        =   179
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
      TabIndex        =   180
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
      TabIndex        =   181
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
      TabIndex        =   182
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
      TabIndex        =   183
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
      TabIndex        =   184
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
      TabIndex        =   185
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
      TabIndex        =   186
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
      TabIndex        =   187
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
      TabIndex        =   188
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
      TabIndex        =   189
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
      TabIndex        =   190
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
      TabIndex        =   191
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
      TabIndex        =   192
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
      TabIndex        =   193
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
      TabIndex        =   194
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
      TabIndex        =   195
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
      TabIndex        =   196
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
      TabIndex        =   197
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
      TabIndex        =   198
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
      TabIndex        =   199
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
      TabIndex        =   200
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
      TabIndex        =   201
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
      TabIndex        =   202
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
      TabIndex        =   203
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
      TabIndex        =   204
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
      TabIndex        =   206
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
      TabIndex        =   207
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
      TabIndex        =   208
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
      TabIndex        =   209
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
      TabIndex        =   210
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
      TabIndex        =   211
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
      TabIndex        =   212
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
      TabIndex        =   214
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
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   1
      Left            =   11040
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   2
      Left            =   11040
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   3
      Left            =   11040
      TabIndex        =   217
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   4
      Left            =   11040
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   5
      Left            =   11040
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   6
      Left            =   11040
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   7
      Left            =   11040
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   8
      Left            =   11040
      TabIndex        =   222
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFlag 
      Height          =   285
      Index           =   9
      Left            =   11040
      TabIndex        =   223
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
      TabIndex        =   224
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
      TabIndex        =   225
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
      TabIndex        =   226
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
      TabIndex        =   227
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
      TabIndex        =   228
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
      TabIndex        =   229
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
      TabIndex        =   230
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
      TabIndex        =   231
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
      TabIndex        =   232
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
      TabIndex        =   233
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
      TabIndex        =   234
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
      TabIndex        =   235
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
      TabIndex        =   236
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
      TabIndex        =   237
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
      TabIndex        =   238
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
      TabIndex        =   239
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
      TabIndex        =   240
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
      TabIndex        =   241
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
      TabIndex        =   242
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
      TabIndex        =   243
      Text            =   "Text1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Relatedness Metric"
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
      Left            =   1080
      TabIndex        =   254
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Females"
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
      Left            =   570
      TabIndex        =   253
      Top             =   3285
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Males"
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
      Left            =   840
      TabIndex        =   252
      Top             =   3885
      Width           =   1815
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
      TabIndex        =   250
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
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   248
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15720
      TabIndex        =   247
      ToolTipText     =   "Check box to indicate male was released post-spawn"
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Released"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   246
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15000
      TabIndex        =   245
      ToolTipText     =   "Check box to indicate female was released post-spawn"
      Top             =   1230
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
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   205
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   9
      Left            =   12960
      MouseIcon       =   "Main Form.frx":489A
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":49EC
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   7140
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   8
      Left            =   12960
      MouseIcon       =   "Main Form.frx":395B0
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":39702
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   6540
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   7
      Left            =   12960
      MouseIcon       =   "Main Form.frx":6E2C6
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":6E418
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   5940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   6
      Left            =   12960
      MouseIcon       =   "Main Form.frx":A2FDC
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":A312E
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   5340
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   5
      Left            =   12960
      MouseIcon       =   "Main Form.frx":D7CF2
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":D7E44
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   4740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   4
      Left            =   12960
      MouseIcon       =   "Main Form.frx":10CA08
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":10CB5A
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   4140
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   3
      Left            =   12960
      MouseIcon       =   "Main Form.frx":14171E
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":141870
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   3540
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   2
      Left            =   12960
      MouseIcon       =   "Main Form.frx":176434
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":176586
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   2940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   1
      Left            =   12960
      MouseIcon       =   "Main Form.frx":1AB14A
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":1AB29C
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   2340
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Index           =   0
      Left            =   12960
      MouseIcon       =   "Main Form.frx":1DFE60
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":1DFFB2
      Stretch         =   -1  'True
      ToolTipText     =   "A flag indicating that something could be wrong with this mating pair; Click to display concerns"
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   120
      MouseIcon       =   "Main Form.frx":214B76
      MousePointer    =   99  'Custom
      Picture         =   "Main Form.frx":214CC8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4140
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Relatedness"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   92
      ToolTipText     =   "The proportion of shared alleles between the mating pair"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mated"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   91
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
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   90
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
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   89
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4320
      TabIndex        =   68
      ToolTipText     =   "Family ID of mating pair"
      Top             =   7260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   8
      Left            =   4320
      TabIndex        =   67
      ToolTipText     =   "Family ID of mating pair"
      Top             =   6660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   7
      Left            =   4320
      TabIndex        =   66
      ToolTipText     =   "Family ID of mating pair"
      Top             =   6060
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   65
      ToolTipText     =   "Family ID of mating pair"
      Top             =   5460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   64
      ToolTipText     =   "Family ID of mating pair"
      Top             =   4860
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   63
      ToolTipText     =   "Family ID of mating pair"
      Top             =   4260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   62
      ToolTipText     =   "Family ID of mating pair"
      Top             =   3660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   61
      ToolTipText     =   "Family ID of mating pair"
      Top             =   3060
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   60
      ToolTipText     =   "Family ID of mating pair"
      Top             =   2460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFem 
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
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   59
      ToolTipText     =   "Family ID of mating pair"
      Top             =   1860
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current ID"
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
      Left            =   6840
      TabIndex        =   58
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
      Caption         =   "Mating Input Criteria"
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
      TabIndex        =   57
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   6300
      Left            =   120
      Top             =   2520
      Width           =   4095
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
      TabIndex        =   113
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
      TabIndex        =   112
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
      TabIndex        =   111
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
      TabIndex        =   110
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
      TabIndex        =   109
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
      TabIndex        =   108
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
      TabIndex        =   107
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
      TabIndex        =   106
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
      TabIndex        =   105
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
      TabIndex        =   104
      ToolTipText     =   "Number of male referring to order that male was entered on input form"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDbaseConnection 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Establishing database connection..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   6120
      TabIndex        =   251
      Top             =   3600
      Width           =   8415
   End
   Begin VB.Menu file 
      Caption         =   "&Options"
      Index           =   0
      Begin VB.Menu databaseSettings 
         Caption         =   "&Database Settings"
         HelpContextID   =   1
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu flagSettings 
         Caption         =   "&Warning Flag Settings"
         Index           =   2
         Shortcut        =   ^F
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Index           =   4
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Index           =   6
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
Dim q As Long, tmpStr As Variant, tmpChunk As Variant
Dim tmpDrain() As Variant, tmpCnt As Long, n As Long, p As Long
Dim strFactorial As String, tmpArray(11) As Long, tmpBatch As Long, batchArray() As Variant
Dim tmpDate As Variant, tmpTime As Variant
Dim yesDbase As String, yesFlags As String, yesPercs As String
Dim accessApp As Access.Application, dbTemp As Database

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

Private Sub about_Click(Index As Integer)
    frmAbout.Show 1
End Sub

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
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If

    '1 to 2
    If frmMain.optMatDes(1).Value = True And CInt(frmMain.cmbFem.Text) > (CInt(frmMain.cmbMale.Text) / 2) Then
        Msg = "In a 1 to 2 mating design the number of males must be equal to or greater than twice the number of females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If

    '2 to 2
    If frmMain.optMatDes(2).Value = True And ((CInt(frmMain.cmbFem.Text) > CInt(frmMain.cmbMale.Text)) Or (CInt(frmMain.cmbFem.Text) < 2) Or (CInt(frmMain.cmbFem.Text) Mod 2 = 1)) Then
        Msg = "In a 2 to 2 mating design there must an even number of females greater than or equal to 2, and the number of males must be equal to or greater than the number of females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    '1 to 3
    If frmMain.optMatDes(3).Value = True And CInt(frmMain.cmbFem.Text) > (CInt(frmMain.cmbMale.Text) / 3) Then
        Msg = "In a 1 to 3 mating design the number of males must be equal to or greater than three times the number of females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    '3 to 3
    If frmMain.optMatDes(4).Value = True And ((CInt(frmMain.cmbFem.Text) > CInt(frmMain.cmbMale.Text)) Or (CInt(frmMain.cmbFem.Text) < 3) Or (CInt(frmMain.cmbFem.Text) Mod 3 <> 0)) Then
        Msg = "In a 3 to 3 mating design there must a multiple of three females greater than or equal to 3, and the number of males must be equal to or greater than the number of females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
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
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Reoptimation Not Possible"
        Response = MsgBox(Msg, Style, Title)
        
        Exit Sub
    End If
    
    frmInput.Show 1
End Sub

Private Sub cmdSpawnUp_Click()
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
        Msg = "There are " & j - k & " matings that aren't selected." & Chr(13) & Chr(13) & "Do you wish to continue without selecting those matings?" & Chr(13) & "WARNING: All data will be removed after updating the matings table"
        Style = vbYesNo + vbInformation + vbDefaultButton1
        Title = "Unselected Matings"
        Response = MsgBox(Msg, Style, Title)
        
        If Response = 7 Then
            Exit Sub
        End If
    End If
    
    On Error GoTo describeErr
    
    'Get maximum batch number
    Set rstTemp = dbsNew.OpenRecordset("SELECT Max([" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbBatch.Text & "]) AS MaxOf" & frmDataSpec.cmbBatch.Text & " FROM [" & frmDataSpec.cmbMatingsTable.Text & "];", dbOpenDynaset)
    rstTemp.MoveLast
    rstTemp.MoveFirst
    batchArray = rstTemp.GetRows(rstTemp.RecordCount)
    If IsNull(batchArray(0, 0)) = False Then
        tmpBatch = batchArray(0, 0) + 1
    Else
        tmpBatch = 1
    End If
    
    tmpDate = Format(Date, "MM/DD/YYYY")
    tmpTime = Format(Time, "HH:NN:SS")
    
    'Add new mating records to matings table
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbMatingsTable.Text & "].* FROM [" & frmDataSpec.cmbMatingsTable.Text & "];", dbOpenDynaset)
    For i = 0 To j - 1
        If frmMain.chkSpawned(i).Value = 1 Then
            rstTemp.AddNew
            rstTemp(frmDataSpec.cmbBatch.Text) = tmpBatch
            rstTemp(frmDataSpec.cmbFamID.Text) = frmMain.lblFem(i).Caption
            rstTemp(frmDataSpec.cmbDamID.Text) = frmMain.txtFem(i).Text
            If frmDataSpec.chkPop.Value = 1 Then
                rstTemp(frmDataSpec.cmbDamPop.Text) = frmMain.txtDrainageF(i).Text
            End If
            If frmDataSpec.chkCohort.Value = 1 Then
                rstTemp(frmDataSpec.cmbDamCohort.Text) = frmMain.txtfYear(i).Text
            End If
            rstTemp(frmDataSpec.cmbDamLoci.Text) = frmMain.txtLocFem(i).Text
            rstTemp(frmDataSpec.cmbDamReleased.Text) = frmMain.chkReleaseF(i).Value
            
            rstTemp(frmDataSpec.cmbSireID.Text) = frmMain.txtMale(i).Text
            If frmDataSpec.chkPop.Value = 1 Then
                rstTemp(frmDataSpec.cmbSirePop.Text) = frmMain.txtDrainageM(i).Text
            End If
            If frmDataSpec.chkCohort.Value = 1 Then
                rstTemp(frmDataSpec.cmbSireCohort.Text) = frmMain.txtmYear(i).Text
            End If
            rstTemp(frmDataSpec.cmbSireLoci.Text) = frmMain.txtLocMale(i).Text
            rstTemp(frmDataSpec.cmbSireReleased.Text) = frmMain.chkReleaseM(i).Value
            
            rstTemp(frmDataSpec.cmbMetric.Text) = frmMain.cmbRelMetric.Text
            If IsNumeric(frmMain.txtPShare(i).Text) = True Then
                rstTemp(frmDataSpec.cmbRelatedness.Text) = frmMain.txtPShare(i).Text
            End If
            rstTemp(frmDataSpec.cmbOptimized.Text) = frmMain.txtOptimized(i).Text
            rstTemp(frmDataSpec.cmbDate.Text) = tmpDate
            rstTemp(frmDataSpec.cmbTime.Text) = tmpTime
            If spawnComment(i) <> "" Then
                rstTemp(frmDataSpec.cmbComments.Text) = spawnComment(i)
            End If
            rstTemp.Update
        End If
    Next i
    
    'Add non-optimized matings to matings_non-opt table
    x = 0
    For i = 0 To 9
        If frmInput.txtTagMale(i).Text <> "" Then x = x + 1
    Next i
        
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.txtMatingsNonopt.Text & "].* FROM [" & frmDataSpec.txtMatingsNonopt.Text & "];", dbOpenDynaset)
    For i = 0 To j - 1
        If frmMain.chkSpawned(i).Value = 1 Then
            rstTemp.AddNew
            rstTemp("Batch") = tmpBatch
            rstTemp("Family") = frmMain.lblFem(i).Caption
            rstTemp("Dam") = frmMain.txtFem(i).Tag
            If frmDataSpec.chkPop.Value = 1 Then
                rstTemp("Dam_Pop") = frmMain.txtDrainageF(i).Tag
            End If
            If frmDataSpec.chkCohort.Value = 1 Then
                rstTemp("Dam_Cohort") = frmMain.txtfYear(i).Tag
            End If
            rstTemp("Dam_Scored_Loci") = frmMain.txtLocFem(i).Tag
            
            rstTemp("Sire") = frmMain.txtMale(i).Tag
            If frmDataSpec.chkPop.Value = 1 Then
                rstTemp("Sire_Pop") = frmMain.txtDrainageM(i).Tag
            End If
            If frmDataSpec.chkCohort.Value = 1 Then
                rstTemp("Sire_Cohort") = frmMain.txtmYear(i).Tag
            End If
            rstTemp("Sire_Scored_Loci") = frmMain.txtLocMale(i).Tag
            
            rstTemp("Metric") = frmMain.cmbRelMetric.Text
            If IsNumeric(frmMain.txtPShare(i).Tag) = True Then
                rstTemp("Relatedness") = frmMain.txtPShare(i).Tag
            End If
            rstTemp("Optimized") = "No"
            rstTemp("Date") = tmpDate
            rstTemp("Time") = tmpTime
            rstTemp.Update
        End If
    Next i

    'Clear all matings
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
    
    frmMain.txtAvgPShare.Text = ""
    
    Unload frmInput
    
    Msg = "The selected matings have been added to the table '" & frmDataSpec.cmbMatingsTable.Text & "'."
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
    Call frmDataSpec.formShow
    frmDataSpec.Show 1
End Sub

Private Sub exit_Click(Index As Integer)
    Unload frmMain
End Sub

Private Sub flagSettings_Click(Index As Integer)
    frmFlags.Show 1
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
    frmMain.cmbRelMetric.ListIndex = 0
       
    Set accessApp = CreateObject("Access.Application")
    
    'Load database settings if present
    If Dir(App.Path & "\dbase_settings.mmf") <> "" Then
        Screen.MousePointer = vbHourglass
        yesDbase = frmDataSpec.importDbaseSettings
    Else
        yesDbase = False
    End If
    
    'Load flags settings if present
    If Dir(App.Path & "\flag_settings.mmf") <> "" Then
        Screen.MousePointer = vbHourglass
        yesFlags = frmFlags.importFlagSettings
    Else
        yesFlags = False
    End If
    
    'Load relatedness thresholds if present
    If Dir(App.Path & "\perc_settings.mmf") <> "" Then
        Screen.MousePointer = vbHourglass
        yesPercs = frmFlags.importPercSettings
    Else
        yesPercs = False
    End If
    
    'Open settings forms if necessary
    Screen.MousePointer = vbDefault
    If yesDbase = False Then
        frmMain.lblDbaseConnection.Caption = "Specify 'Database Settings' under the 'Options' menu"
        If frmDataSpec.cmbGeneticsTable.ListIndex = -1 Then
            Call frmDataSpec.disableGenFrame
        End If
        If frmDataSpec.cmbMatingsTable.ListIndex = -1 Then
            Call frmDataSpec.disableMatFrame
        End If
        frmDataSpec.Show 1
    ElseIf yesFlags = False Then
        frmMain.lblDbaseConnection.Caption = "Specify 'Flag Settings' under the 'Options' menu"
        frmFlags.Show 1
    ElseIf yesPercs = False Then
        If frmFlags.chkUppQuart.Value = 1 Then
            frmMain.lblDbaseConnection.Caption = "Calculate relatedness values by opening 'Flag Settings' under the 'Options' menu"
            frmFlags.Show 1
        Else
            frmMain.cmdInputPIT.Enabled = True
            frmMain.lblDbaseConnection.Visible = False
        End If
    Else
        frmMain.cmdInputPIT.Enabled = True
        frmMain.lblDbaseConnection.Visible = False
    End If
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
    
    End
End Sub

Private Sub Image1_Click()
        frmAbout.Show 1
End Sub

Private Sub imgFlag_Click(Index As Integer)
    Msg = frmMain.txtFlag(Index).Text
    Style = vbOKOnly + vbExclamation + vbDefaultButton1
    Title = "Mating Flags"
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

Public Function setDBase()
    Set accessApp = CreateObject("Access.Application")
    On Error Resume Next
    accessApp.CloseCurrentDatabase
    accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
    On Error GoTo 0
    Set dbsNew = accessApp.CurrentDb
End Function
