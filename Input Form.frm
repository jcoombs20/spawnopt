VERSION 5.00
Begin VB.Form frmInput 
   Caption         =   "PIT Tag Input"
   ClientHeight    =   8640
   ClientLeft      =   915
   ClientTop       =   1245
   ClientWidth     =   13680
   Icon            =   "Input Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearTag 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdSexChange 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Update the sex of the current individual"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   10
      Left            =   9120
      TabIndex        =   102
      Top             =   7200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   9
      Left            =   9120
      TabIndex        =   101
      Top             =   6600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   8
      Left            =   9120
      TabIndex        =   100
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   7
      Left            =   9120
      TabIndex        =   99
      Top             =   5400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   98
      Top             =   4800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   5
      Left            =   9120
      TabIndex        =   97
      Top             =   4200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   96
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   3
      Left            =   9120
      TabIndex        =   95
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   94
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   93
      Top             =   1800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   0
      Left            =   9120
      TabIndex        =   92
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3360
      TabIndex        =   38
      ToolTipText     =   "Closes the current form"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CheckBox chkOptimize 
      Caption         =   "Optimize Matings"
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
      Left            =   8640
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Choice of whether or not to optimize matings"
      Top             =   8160
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.ListBox lstXSpawned 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "List of prior spawning dates for current individual"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   22
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   20
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   16
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   6
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFamID 
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
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   1185
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDrainage 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Drainage of most recently scanned individual"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtTagCur2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Most recently scanned PIT tag"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddCurTag 
      Caption         =   "Add Current Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Adds current PIT tag to the appropriate list of females and males"
      Top             =   6240
      Width           =   3015
   End
   Begin VB.TextBox txtSex 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Sex of most recently scanned individual"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtTagCur 
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
      Left            =   240
      LinkItem        =   "Field(1)"
      LinkTopic       =   "winwedge|COM1"
      TabIndex        =   85
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
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
      Left            =   6000
      TabIndex        =   37
      ToolTipText     =   "Clears all data for spawning females and males"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Optimize"
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
      Left            =   12240
      TabIndex        =   36
      ToolTipText     =   "Optimizes female and male pairings and displays them on the initial form"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   12360
      TabIndex        =   34
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Index           =   10
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   12360
      TabIndex        =   33
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   12360
      TabIndex        =   32
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   12360
      TabIndex        =   31
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   12360
      TabIndex        =   30
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   12360
      TabIndex        =   29
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   12360
      TabIndex        =   28
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   12360
      TabIndex        =   27
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   12360
      TabIndex        =   26
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   12360
      TabIndex        =   25
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearM 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   12360
      TabIndex        =   24
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagMale 
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   7320
      TabIndex        =   23
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   7320
      TabIndex        =   21
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   7320
      TabIndex        =   19
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   7320
      TabIndex        =   17
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   7320
      TabIndex        =   15
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7320
      TabIndex        =   13
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   7320
      TabIndex        =   7
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearF 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   7320
      TabIndex        =   5
      ToolTipText     =   "Clears this individuals PIT tag"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTagFem 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtYear 
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
      Left            =   120
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label lblMD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1 to 1 Mating Design"
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
      Left            =   6480
      TabIndex        =   106
      Top             =   90
      Width           =   4935
   End
   Begin VB.Label lblYearclass 
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Year"
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
      Left            =   240
      TabIndex        =   104
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use"
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
      Left            =   8880
      TabIndex        =   91
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      X1              =   8565
      X2              =   8565
      Y1              =   600
      Y2              =   7680
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
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
      Left            =   3480
      TabIndex        =   90
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Drainage"
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
      Left            =   240
      TabIndex        =   89
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Prior Spawnings"
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
      Left            =   240
      TabIndex        =   87
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1560
      Left            =   120
      Picture         =   "Input Form.frx":030A
      Stretch         =   -1  'True
      ToolTipText     =   "Picture of a DNA double helix"
      Top             =   6990
      Width           =   3075
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Left            =   240
      TabIndex        =   84
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tag"
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
      Left            =   240
      TabIndex        =   83
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Males"
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
      Left            =   10320
      TabIndex        =   82
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Females"
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
      Left            =   5400
      TabIndex        =   81
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
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
      Index           =   10
      Left            =   9600
      TabIndex        =   79
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   77
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   75
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   73
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   71
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   69
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   67
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   65
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   63
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   61
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMale 
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
      Left            =   9600
      TabIndex        =   59
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   57
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   55
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   53
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   51
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   49
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   47
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   45
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   43
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3480
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   7335
      Left            =   3360
      Top             =   480
      Width           =   10215
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg, Style, Title, Response, Default, wrkJet As Workspace, dbsNew As Database
Dim rstTemp As Recordset, famID() As String, fOrder() As String, mOrder() As String
Dim fAlleles(), mAlleles(), maleFinal() As Integer, femFinal() As Integer, pTempMeasure() As Single
Dim pFinalMeasure() As Single, strFactorial As String
Dim pTotal As Single, pFinal As Single, pShared As Long
Dim pMeasure As Single, totAlleles As Long, maleOrder() As Long
Dim priorSpawn As Long, femScored() As Long, maleScored() As Long, locScored As Long
Dim tempTag As String, fPic() As String, fWeight() As Single, mPic() As String
Dim mWeight() As Single, mDrainage(11) As String, fDrainage(10) As String
Dim rstCull As Recordset, flaGG() As Variant, numDrainage As Long, disAbledF As Long
Dim disAbledM As Long, enNumF() As Long, enNumM() As Long, enableCnt As Long
Dim mYear(11) As Long, fYear(10) As Long, z As Long, enTrans As String
Dim tmpMOrder() As String, pwPSA() As Single, numFem As Long
Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long
Dim femMD As Integer, numF As Integer, numM As Integer, mateNum As Integer, matDes As Integer
Dim pPartial As Single, tmpBi As Integer

Private Sub chkOptimize_Click()
    If frmInput.chkOptimize.Value = 0 Then
        Msg = "Confirm that you want to disable mating optimization"
        Style = vbYesNo + vbQuestion + vbDefaultButton1
        Title = "Disable Optimization"
        Response = MsgBox(Msg, Style, Title)
                
        If Response = 7 Then
            frmInput.chkOptimize.Value = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub chkUse_Click(Index As Integer)
    For z = 0 To (CInt(frmMain.cmbMale.Text) - 1)
        If frmInput.chkUse(z).Enabled = False Or frmInput.chkUse(z).Value = 0 Then
            frmInput.lblMale(z).Enabled = False
            frmInput.txtTagMale(z).Enabled = False
            'frmInput.txtPicMale(z).Enabled = False
            'frmInput.txtWeightMale(z).Enabled = False
            frmInput.cmdClearM(z).Enabled = False
        Else
            frmInput.lblMale(z).Enabled = True
            frmInput.txtTagMale(z).Enabled = True
            'frmInput.txtPicMale(z).Enabled = True
            'frmInput.txtWeightMale(z).Enabled = True
            frmInput.cmdClearM(z).Enabled = True
        End If
    Next z
End Sub

Private Sub cmdAddCurTag_Click()
    i = 1
    Do Until Mid(frmInput.txtTagCur2.Text, i, 1) = ""
        i = i + 1
    Loop
    
    If i <> 11 And i <> 15 Then
        Msg = "Tag numbers must be 10 or 14 characters long."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Tag Error"
        Response = MsgBox(Msg, Style, Title)
        
        frmInput.txtTagCur2.SetFocus
        Exit Sub
    End If
    
    If UCase(frmInput.txtSex.Text) = "MALE" Then
        For i = 0 To CInt(frmMain.cmbMale.Text) - 1
            If UCase(frmInput.txtTagCur2.Text) = UCase(frmInput.txtTagMale(i).Text) Then
                Msg = "This tag number is already present in male number " & i + 1 & "." & Chr(13) & "Do you wish to add it again?"
                Style = vbYesNo + vbCritical + vbDefaultButton1
                Title = "Duplicate Tag"
                Response = MsgBox(Msg, Style, Title)
                
                If Response = 7 Then Exit Sub
            End If
        Next i
        
        For i = 0 To 10
            If frmInput.txtTagMale(i).Visible = True Then
                If frmInput.txtTagMale(i).Text = "" Then
                    frmInput.txtTagMale(i).Locked = False
                    frmInput.txtTagMale(i).Text = UCase(frmInput.txtTagCur2.Text)
                    frmInput.txtTagMale(i).Locked = True
                    mDrainage(i) = frmInput.txtDrainage.Text
                    mYear(i) = CInt(frmInput.txtYear.Text)
                    GoTo clearData
                End If
            End If
        Next i
        Msg = "There are no empty spaces to add an additional MALE tag number."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
    ElseIf UCase(frmInput.txtSex.Text) = "FEMALE" Then
        For i = 0 To CInt(frmMain.cmbFem.Text) - 1
            If UCase(frmInput.txtTagCur2.Text) = UCase(frmInput.txtTagFem(i).Text) Then
                Msg = "This tag number is already present in female number " & i + 1 & "." & Chr(13) & "Do you wish to add it again?"
                Style = vbYesNo + vbCritical + vbDefaultButton1
                Title = "Duplicate Tag"
                Response = MsgBox(Msg, Style, Title)
                
                If Response = 7 Then Exit Sub
            End If
        Next i
        
        For i = 0 To 9
            If frmInput.txtTagFem(i).Visible = True Then
                If frmInput.txtTagFem(i).Text = "" Then
                    frmInput.txtTagFem(i).Locked = False
                    frmInput.txtTagFem(i).Text = UCase(frmInput.txtTagCur2.Text)
                    frmInput.txtTagFem(i).Locked = False
                    fDrainage(i) = frmInput.txtDrainage.Text
                    fYear(i) = CInt(frmInput.txtYear.Text)
                    GoTo clearData
                End If
            End If
        Next i
        Msg = "There are no empty spaces to add an additional FEMALE tag number."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
    
    End If

clearData:
    frmInput.txtTagCur.Text = ""
    frmInput.txtTagCur2.Text = ""
    frmInput.txtSex.Text = ""
    frmInput.txtDrainage.Text = ""
    frmInput.txtYear.Text = ""
    frmInput.lstXSpawned.Clear
    
    frmInput.txtTagCur2.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    frmInput.txtTagCur2.Text = ""
    frmInput.txtSex.Text = ""
    frmInput.txtDrainage.Text = ""
    frmInput.lstXSpawned.Clear
    
    frmInput.Hide
End Sub

Private Sub cmdClear_Click()
    Msg = "Do you want to clear all tag information?"
    Style = vbYesNo + vbQuestion + vbDefaultButton1
    Title = "Clear All Tags"
    Response = MsgBox(Msg, Style, Title)
            
    If Response = 7 Then
        Exit Sub
    End If
    
    For i = 0 To CInt(frmMain.cmbFem.Text) - 1
        frmInput.txtFamID(i).Text = ""
        frmInput.txtTagFem(i).Text = ""
        'frmInput.txtPicFem(i).Text = ""
        'frmInput.txtWeightFem(i).Text = ""
    Next i
   
    For j = 0 To CInt(frmMain.cmbMale.Text) - 1
        frmInput.chkUse(j).Value = 1
        frmInput.txtTagMale(j).Text = ""
        frmInput.txtTagMale(j).Enabled = True
        'frmInput.txtPicMale(j).Text = ""
        'frmInput.txtPicMale(j).Enabled = True
        'frmInput.txtWeightMale(j).Text = ""
        'frmInput.txtWeightMale(j).Enabled = True
    Next j
End Sub

Private Sub cmdClearF_Click(Index As Integer)
    If frmInput.txtTagFem(Index).Text <> "" Then
        Msg = "Would you like to provide a reason for clearing this individual?"
        Style = vbYesNo + vbQuestion + vbDefaultButton1
        Title = "Reason for Clearing"
        Response = MsgBox(Msg, Style, Title)
    
        If Response = 6 Then 'yes
            Dim tmpTag As String
            Dim secondForm As New frmAddReason
            tmpTag = frmInput.txtTagFem(Index).Text
            secondForm.tmpTagReason = tmpTag
            secondForm.Show 1
        End If
    End If
    
    'frmInput.txtFamID(Index).Text = ""
    frmInput.txtTagFem(Index).Text = ""
    'frmInput.txtPicFem(Index).Text = ""
    'frmInput.txtWeightFem(Index).Text = ""
    Unload frmAddReason
    
End Sub

Private Sub cmdClearM_Click(Index As Integer)
    frmInput.txtTagMale(Index).Text = ""
    'frmInput.txtPicMale(Index).Text = ""
    'frmInput.txtWeightMale(Index).Text = ""
End Sub

Private Sub cmdClearTag_Click()
    frmInput.txtTagCur2.Text = ""
    frmInput.txtTagCur2.SetFocus
End Sub

Public Function getAlleles(rstTemp)
    With rstTemp
        ReDim fAlleles(femMD - disAbledF, rstTemp.Fields.Count - 8, 2)
        ReDim mAlleles(numM - disAbledM, rstTemp.Fields.Count - 8, 2)
        ReDim femFinal(femMD - disAbledF)
        ReDim maleFinal(femMD - disAbledF)
        ReDim pTempMeasure(femMD - disAbledF)
        ReDim pFinalMeasure(femMD - disAbledF)
        ReDim maleOrder(numM - disAbledM)
        ReDim femScored(femMD - disAbledF)
        ReDim maleScored(numM - disAbledM)
                
        'acquiring male and female allele values
        For j = 1 To femMD - disAbledF
            rstTemp.MoveFirst
            Do Until rstTemp.EOF
                If UCase(![pit]) = UCase(fOrder(j)) Then
                    m = 1
                    locScored = 0
                    For k = 7 To rstTemp.Fields.Count - 2
                        fAlleles(j, m, 1) = rstTemp.Fields(k).Value
                        If IsNull(rstTemp.Fields(k).Value) = False Then locScored = locScored + 1
                        k = k + 1
                        fAlleles(j, m, 2) = rstTemp.Fields(k).Value
                        If IsNull(rstTemp.Fields(k).Value) = False Then locScored = locScored + 1
                        m = m + 1
                    Next k
                    femScored(j) = locScored / 2
                    Exit Do
                End If
                rstTemp.MoveNext
            Loop
        Next j
            
        For j = 1 To numM - disAbledM
            rstTemp.MoveFirst
            Do Until rstTemp.EOF
                If UCase(![pit]) = UCase(mOrder(j)) Then
                    m = 1
                    locScored = 0
                    For k = 7 To rstTemp.Fields.Count - 2
                        mAlleles(j, m, 1) = rstTemp.Fields(k).Value
                        If IsNull(rstTemp.Fields(k).Value) = False Then locScored = locScored + 1
                        k = k + 1
                        mAlleles(j, m, 2) = rstTemp.Fields(k).Value
                        If IsNull(rstTemp.Fields(k).Value) = False Then locScored = locScored + 1
                        m = m + 1
                    Next k
                    maleScored(j) = locScored / 2
                    Exit Do
                End If
                rstTemp.MoveNext
            Loop
        Next j
    End With
End Function

Public Function propSA()
    'calculating proportion of shared alleles
    ReDim pwPSA(femMD - disAbledF, numM - disAbledM)
    
    For f = 1 To femMD - disAbledF
        For m = 1 To numM - disAbledM
            pTotal = 0
            totAlleles = rstTemp.Fields.Count - 8: pShared = 0
            For j = 1 To totAlleles / 2
                If IsNull(fAlleles(f, j, 1)) = True Or IsNull(fAlleles(f, j, 2)) = True Or IsNull(mAlleles(m, j, 1)) = True Or IsNull(mAlleles(m, j, 2)) = True Then
                    totAlleles = totAlleles - 2
                    GoTo nextLoci
                End If
        
                If fAlleles(f, j, 1) = fAlleles(f, j, 2) Then
                    If fAlleles(f, j, 1) = mAlleles(m, j, 1) And fAlleles(f, j, 1) = mAlleles(m, j, 2) Then
                        pShared = pShared + 2
                    ElseIf fAlleles(f, j, 1) = mAlleles(m, j, 1) Or fAlleles(f, j, 1) = mAlleles(m, j, 2) Then
                        pShared = pShared + 1
                    End If
                ElseIf mAlleles(m, j, 1) = mAlleles(m, j, 2) Then
                    If fAlleles(f, j, 1) = mAlleles(m, j, 1) And fAlleles(f, j, 2) = mAlleles(m, j, 1) Then
                        pShared = pShared + 2
                    ElseIf fAlleles(f, j, 1) = mAlleles(m, j, 1) Or fAlleles(f, j, 2) = mAlleles(m, j, 1) Then
                        pShared = pShared + 1
                    End If
                Else
                    If fAlleles(f, j, 1) = mAlleles(m, j, 1) Or fAlleles(f, j, 1) = mAlleles(m, j, 2) Then
                        pShared = pShared + 1
                    End If
            
                    If fAlleles(f, j, 2) = mAlleles(m, j, 1) Or fAlleles(f, j, 2) = mAlleles(m, j, 2) Then
                        pShared = pShared + 1
                    End If
                End If
nextLoci:
            Next j
            
            If totAlleles = 0 Then
                pwPSA(f, m) = 0
            ElseIf pShared = 0 Then
                pwPSA(f, m) = 10
            ElseIf pShared = totAlleles Then
                pwPSA(f, m) = 0
            Else
                pwPSA(f, m) = Format((-Log(pShared / totAlleles)), "0.000")
                'Debug.Print (fOrder(f) & " " & mOrder(m) & " " & Format(pShared / totAlleles, "0.000"))
            End If
        Next m
    Next f
End Function

Public Function getMales(tmpColl)
    'numFem = CInt(frmMain.cmbFem.Text) - disAbledF
    Dim arr As Variant
    For Each arr In tmpColl
        
        'summing proportion of shared alleles across matings
        pTotal = 0
        For n = 1 To femMD - disAbledF
            pTotal = pTotal + pwPSA(n, arr(n - 1))
            pTempMeasure(n) = pwPSA(n, arr(n - 1))
        Next n
                            
        'determining which combination is best
        If pTotal > pFinal Then
            pFinal = pTotal
            For n = 1 To femMD - disAbledF
                maleFinal(n) = arr(n - 1)
                pFinalMeasure(n) = Format(Exp(-(pTempMeasure(n))), "0.000")
            Next n
        End If
    Next
    
    For n = 1 To femMD - disAbledF
        femFinal(n) = n
    Next n
End Function

Public Function getMalesRecip(tmpColl)
    Dim fColl As New Collection
    Dim arrF As Variant
    Dim arr As Variant
    
    If numF > mateNum Then
        If matDes = 2 Then
            For Each arr In fList4
                fColl.Add (arr)
            Next
        End If
    Else
        If mateNum = 2 Then
            fColl.Add Array(1, 2, 3, 4)
        Else
            fColl.Add Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
        End If
    End If
    
    For Each arrF In fColl
        For Each arr In tmpColl
            'summing proportion of shared alleles across matings
            pTotal = 0
            m = 1
            For n = 1 To femMD - disAbledF
                For k = 1 To mateNum
                    j = 0
                    For i = m To m + (mateNum - 1)
                        pTotal = pTotal + pwPSA(arrF(n), arr(i - 1))
                        pTempMeasure(n + j) = pwPSA(arrF(n), arr(i - 1))
                        j = j + 1
                    Next i
                    n = n + mateNum
                Next k
                m = m + mateNum
                n = n - 1
            Next n
                                
            'determining which combination is best
            If pTotal > pFinal Then
                pFinal = pTotal
                For n = 1 To femMD - disAbledF
                    pFinalMeasure(n) = Format(Exp(-(pTempMeasure(n))), "0.000")
                Next n
                
                For n = 1 To femMD - disAbledF
                    femFinal(n) = arrF(n - 1)
                Next n
                
                m = 1
                For n = 1 To femMD - disAbledF
                    For k = 1 To mateNum
                        j = 0
                        For i = m To m + (mateNum - 1)
                            maleFinal(n + j) = arr(i - 1)
                            j = j + 1
                        Next i
                        n = n + mateNum
                    Next k
                    m = m + mateNum
                    n = n - 1
                Next n
            End If
        Next
    Next
    
    ReDim Preserve maleScored(numM * mateNum)
    j = numM
    For i = (numM * mateNum) - disAbledM To 1 Step -mateNum
        For m = 0 To mateNum - 1
            maleScored(i - m) = maleScored(j)
        Next m
        j = j - 1
    Next i
End Function

Public Sub cmdOpt_Click()
    disAbledF = 0: disAbledM = 0
    numF = CInt(frmMain.cmbFem.Text)
    numM = CInt(frmMain.cmbMale.Text)
    
    For i = 0 To frmMain.optMatDes.Count - 1
        If frmMain.optMatDes(i).Value = True Then
            matDes = i
            Exit For
        End If
    Next i
    
    If matDes = 0 Then
        mateNum = 1
    ElseIf matDes <= 2 Then
        mateNum = 2
    Else
        mateNum = 3
    End If
    femMD = numF * mateNum

        
    For i = 0 To femMD - 1
        If frmInput.txtTagFem(i).Visible = True And frmInput.txtTagFem(i).Text = "" Then
            Msg = "There is an empty space in the Female list." + Chr(13) + "Please either scan a tag into the empty space or reselect the appropriate number of females."
            Style = vbOKOnly + vbCritical + vbDefaultButton1
            Title = "Missing Female"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
        
        If frmInput.txtFamID(i).Text = "" Then
            Msg = "There is an empty space in the Family ID list." + Chr(13) + "Please enter in the appropriate family identification number."
            Style = vbOKOnly + vbCritical + vbDefaultButton1
            Title = "Missing Family ID"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
            
        If frmInput.txtFamID(i).Enabled = False Then
            disAbledF = disAbledF + 1
        End If
    Next i
        
    If disAbledF = femMD Then
        Msg = "There are no enabled female tags to optimize." + Chr(13) + "Please cancel this form and either deselect a spawning or reselect the number of females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "No Females to Optimize"
        Response = MsgBox(Msg, Style, Title)
            
        Exit Sub
    End If
                
    For i = 0 To numM - 1
        If frmInput.txtTagMale(i).Text = "" Then
            Msg = "There is an empty space in the Male list." + Chr(13) + "Please either scan a tag into the empty space or reselect the appropriate number of males."
            Style = vbOKOnly + vbCritical + vbDefaultButton1
            Title = "Missing Male"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
            
        If frmInput.txtTagMale(i).Enabled = False Then
            disAbledM = disAbledM + 1
        End If
    Next i
                    
    If femMD - disAbledF > numM - disAbledM Then
        Msg = "There must be at least as many enabled males as enabled females."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Not Enough Enabled Males"
        Response = MsgBox(Msg, Style, Title)
            
        Exit Sub
    End If
    
    ReDim famID(femMD - disAbledF) As String
    ReDim fOrder(femMD - disAbledF) As String
    ReDim enNumF(femMD - disAbledF) As Long
    
    ReDim flaGG(femMD - disAbledF) As Variant
    
    'read in female tags and fam ids
    j = 1
    For i = 1 To femMD
        For m = 1 To mateNum
            If frmInput.txtFamID((i - 1) + (m - 1)).Enabled = True Then
                enNumF(j) = (i - 1) + (m - 1)
                famID(j) = frmInput.txtFamID((i - 1) + (m - 1)).Text
                fOrder(j) = frmInput.txtTagFem(i - 1).Text
                fDrainage(j - 1) = fDrainage(i - 1)
                j = j + 1
            End If
        Next m
        i = i + (mateNum - 1)
    Next i
    
    ReDim mOrder(numM - disAbledM) As String
    ReDim enNumM(numM - disAbledM) As Long
    
    'read in male tags
    j = 1
    For i = 1 To numM
        If frmInput.txtTagMale(i - 1).Enabled = True Then
            enNumM(j) = i - 1
            mOrder(j) = frmInput.txtTagMale(i - 1).Text
            j = j + 1
        End If
    Next i
    
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    Set rstTemp = dbsNew.OpenRecordset("SELECT Genetics.* FROM Genetics", dbOpenDynaset)
    
    If frmInput.chkOptimize.Value = 1 Then
        '***Read in mating individuals genotypes
        Call getAlleles(rstTemp)
        
        '***Get proportion of shared alleles (add other relatedness options here)
        Call propSA
        
        
        '***Find optimal matings
        pFinal = 0
        
        Select Case numM - disAbledM
            Case 1
                Call getMales(mList1)
            Case 2
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList2)
                Else
                    Call getMales(mList2)
                End If
            Case 3
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList3)
                Else
                    Call getMales(mList3)
                End If
            Case 4
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList4)
                Else
                    Call getMales(mList4)
                End If
            Case 5
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList5)
                Else
                    Call getMales(mList5)
                End If
            Case 6
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList6)
                Else
                    Call getMales(mList6)
                End If
            Case 7
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList7)
                Else
                    Call getMales(mList7)
                End If
            Case 8
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList8)
                Else
                    Call getMales(mList8)
                End If
            Case 9
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList9)
                Else
                    Call getMales(mList9)
                End If
            Case 10
                If matDes = 2 Or matDes = 4 Then
                    Call getMalesRecip(mList10)
                Else
                    Call getMales(mList10)
                End If
            Case 11
                'Call getMales(mList11)
        End Select
                
        
        '******Update main form
        For m = 0 To femMD - 1
            frmMain.lblFem(m).Visible = True
            frmMain.txtFem(m).Visible = True
            frmMain.lblMale(m).Visible = True
            frmMain.txtMale(m).Visible = True
            frmMain.txtPShare(m).Visible = True
            frmMain.cmdComment(m).Visible = True
            frmMain.chkSpawned(m).Visible = True
            frmMain.chkReleaseF(m).Visible = True
            frmMain.chkReleaseM(m).Visible = True
        Next m
        
        Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMating.* FROM tblBroodMating", dbOpenDynaset)
        
        m = 1
        'pFinal = 0
        For p = 1 To femMD
            If frmInput.txtFamID(p - 1).Enabled = True Then
                frmMain.lblFem(enNumF(m)).Caption = famID(femFinal(m))
                frmMain.txtFem(enNumF(m)).Locked = False
                frmMain.txtFem(enNumF(m)).Text = fOrder(femFinal(m))
                frmMain.txtFem(enNumF(m)).Locked = True
                frmMain.txtLocFem(enNumF(m)).Text = femScored(femFinal(m))
                
                Select Case UCase(fDrainage(enNumF(femFinal(m))))
                    Case "PENOBSCOT"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 1
                    Case "SHEEPSCOT"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 11
                    Case "NARRAGUAGUS"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 3
                    Case "PLEASANT"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 4
                    Case "MACHIAS"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 5
                    Case "EAST MACHIAS"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 6
                    Case "DENNYS"
                        frmMain.txtDrainageF(enNumF(femFinal(m))).Text = 7
                End Select
                
                
                frmMain.txtfYear(enNumF(m)).Text = fYear(enNumF(femFinal(m)))
                frmMain.lblMale(enNumF(m)).Caption = enNumM(maleFinal(m)) + 1
                frmMain.txtMale(enNumF(m)).Locked = False
                frmMain.txtMale(enNumF(m)).Text = mOrder(maleFinal(m))
                frmMain.txtMale(enNumF(m)).Locked = True
                frmMain.txtLocMale(enNumF(m)).Text = maleScored(maleFinal(m))
                        
                Select Case UCase(mDrainage(maleFinal(m) - 1))
                    Case "PENOBSCOT"
                        frmMain.txtDrainageM(enNumF(m)).Text = 1
                    Case "SHEEPSCOT"
                        frmMain.txtDrainageM(enNumF(m)).Text = 11
                    Case "NARRAGUAGUS"
                        frmMain.txtDrainageM(enNumF(m)).Text = 3
                    Case "PLEASANT"
                        frmMain.txtDrainageM(enNumF(m)).Text = 4
                    Case "MACHIAS"
                        frmMain.txtDrainageM(enNumF(m)).Text = 5
                    Case "EAST MACHIAS"
                        frmMain.txtDrainageM(enNumF(m)).Text = 6
                    Case "DENNYS"
                        frmMain.txtDrainageM(enNumF(m)).Text = 7
                End Select
                
                frmMain.txtmYear(enNumF(m)).Text = mYear(maleFinal(m) - 1)
                frmMain.txtPShare(enNumF(m)).Locked = False
                frmMain.txtPShare(enNumF(m)).Text = Format(pFinalMeasure(m), "0.000")
                'pFinal = pFinal + pFinalMeasure(m)
                frmMain.txtPShare(enNumF(m)).Locked = True
                
                If disAbledF > 0 Or disAbledM > 0 Then
                    frmMain.txtOptimized(enNumF(m)) = "REOPTIMIZED"
                Else
                    frmMain.txtOptimized(enNumF(m)) = "YES"
                End If
                
                'Determine mating flags
                If UCase(frmMain.txtDrainageF(enNumF(femFinal(m))).Text) <> UCase(frmMain.txtDrainageM(enNumF(m)).Text) Then
                    flaGG(m - 1) = "The male and female for this mating come from two different drainages." & Chr(13) & Chr(13)
                End If
                
                Select Case CInt(frmMain.txtDrainageF(enNumF(femFinal(m))).Text)
                    Case 7
                        numDrainage = 0
                    Case 6
                        numDrainage = 1
                    Case 5
                        numDrainage = 2
                    Case 3
                        numDrainage = 3
                    Case 1
                        numDrainage = 4
                    Case 4
                        numDrainage = 5
                    Case 11
                        numDrainage = 6
                    Case Else
                        numDrainage = 7
                End Select
                
                If femScored(femFinal(m)) > 0 And maleScored(maleFinal(m)) > 0 Then
                    If pFinalMeasure(m) > CSng(frmMain.txtQuartile(numDrainage).Text) Then
                        flaGG(m - 1) = flaGG(m - 1) & "The proportion of shared alleles for this mating is in the upper quartile for the " & fDrainage(enNumF(m)) & " drainage (>" & frmMain.txtQuartile(numDrainage).Text & ")." & Chr(13) & Chr(13)
                    End If
                End If
                
                Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMating.* FROM tblBroodMating", dbOpenDynaset)
                With rstTemp
                    rstTemp.MoveFirst
                    Do Until rstTemp.EOF
                        If UCase(rstTemp!Dam) = UCase(frmMain.txtFem(enNumF(m)).Text) Then
                            If UCase(rstTemp!Sire) = UCase(frmMain.txtMale(enNumF(m)).Text) Then
                                flaGG(m - 1) = flaGG(m - 1) & "This female and male pairing has already been spawned on " & rstTemp!TakeDate & Chr(13) & Chr(13)
                            End If
                        End If
                        rstTemp.MoveNext
                    Loop
                End With
                        
                If femScored(femFinal(m)) = 0 Then
                    flaGG(m - 1) = flaGG(m - 1) & "This female has not been scored at any loci" & Chr(13) & Chr(13)
                End If
                
                If maleScored(maleFinal(m)) = 0 Then
                    flaGG(m - 1) = flaGG(m - 1) & "This male has not been scored at any loci" & Chr(13) & Chr(13)
                End If
                
                If flaGG(m - 1) <> Empty Then
                    frmMain.imgFlag(enNumF(m)).Visible = True
                    frmMain.txtFlag(enNumF(m)).Text = flaGG(m - 1)
                Else
                    frmMain.imgFlag(enNumF(m)).Visible = False
                End If
                
                m = m + 1
            End If
        Next p
        
        'frmMain.lblAvgPShare.Visible = True
        'frmMain.txtAvgPShare.Text = Format((pFinal / femMD), "0.000")
        'frmMain.txtAvgPShare.Visible = True
        
        For m = femMD To 9
            frmMain.lblFem(m).Visible = False
            frmMain.txtFem(m).Visible = False
            frmMain.txtFem(m).Locked = False
            frmMain.txtFem(m).Text = ""
            frmMain.txtFem(m).Locked = True
            frmMain.lblMale(m).Visible = False
            frmMain.txtMale(m).Visible = False
            frmMain.txtMale(m).Locked = False
            frmMain.txtMale(m).Text = ""
            frmMain.txtMale(m).Locked = True
            frmMain.txtPShare(m).Visible = False
            frmMain.cmdComment(m).Visible = False
            frmMain.imgFlag(m).Visible = False
            frmMain.chkSpawned(m).Visible = False
            frmMain.chkReleaseF(m).Visible = False
            frmMain.chkReleaseM(m).Visible = False
        Next m
    
        dbsNew.Close
        wrkJet.Close
        
    Else '***No optimization, use straight order
        '***Read in mating individuals genotypes
        Call getAlleles(rstTemp)
                
        '***Get proportion of shared alleles (add other relatedness options here)
        Call propSA
        
        '***Adjust mOrder for 2 to 2 and 3 to 3 mating designs
        If matDes = 2 Or matDes = 4 Then
            m = 1
            For n = 1 To femMD - disAbledF
                For k = 1 To mateNum
                    j = 0
                    For i = m To m + (mateNum - 1)
                        maleFinal(n + j) = i
                        j = j + 1
                    Next i
                    n = n + mateNum
                Next k
                m = m + mateNum
                n = n - 1
            Next n
            
            ReDim Preserve mOrder(femMD)
            ReDim Preserve maleScored(femMD)
            For n = femMD To 1 Step -1
                mOrder(n) = mOrder(maleFinal(n))
                maleScored(n) = maleScored(maleFinal(n))
            Next n
        Else
            For n = 1 To femMD - disAbledF
                maleFinal(n) = n
            Next n
        End If
        
        
        z = 1
        'pFinal = 0
        For i = 0 To femMD - 1
            frmMain.lblFem(i).Visible = True
            frmMain.txtFem(i).Visible = True
            frmMain.lblMale(i).Visible = True
            frmMain.txtMale(i).Visible = True
            frmMain.txtPShare(i).Visible = True
            frmMain.cmdComment(i).Visible = True
            frmMain.chkSpawned(i).Visible = True
            frmMain.chkReleaseF(i).Visible = True
            frmMain.chkReleaseM(i).Visible = False
            
            If frmMain.txtFem(i).BackColor <> &HFF& Then
                frmMain.lblFem(i).Caption = famID(z)
                frmMain.txtFem(i).Text = fOrder(z)
                frmMain.txtLocFem(i).Text = femScored(z)
                frmMain.txtfYear(i).Text = fYear(enNumF(z))
            
                Select Case UCase(fDrainage(enNumF(z)))
                    Case "PENOBSCOT"
                        frmMain.txtDrainageF(i).Text = 1
                    Case "SHEEPSCOT"
                        frmMain.txtDrainageF(i).Text = 11
                    Case "NARRAGUAGUS"
                        frmMain.txtDrainageF(i).Text = 3
                    Case "PLEASANT"
                        frmMain.txtDrainageF(i).Text = 4
                    Case "MACHIAS"
                        frmMain.txtDrainageF(i).Text = 5
                    Case "EAST MACHIAS"
                        frmMain.txtDrainageF(i).Text = 6
                    Case "DENNYS"
                        frmMain.txtDrainageF(i).Text = 7
                End Select
            
                frmMain.lblMale(i).Caption = frmInput.lblMale(maleFinal(z) - 1).Caption
                frmMain.txtMale(i).Text = mOrder(z)
                frmMain.txtLocMale(i).Text = maleScored(z)
                frmMain.txtmYear(i).Text = mYear(maleFinal(z) - 1)
            
                Select Case UCase(mDrainage(maleFinal(z) - 1))
                    Case "PENOBSCOT"
                        frmMain.txtDrainageM(i).Text = 1
                    Case "SHEEPSCOT"
                        frmMain.txtDrainageM(i).Text = 11
                    Case "NARRAGUAGUS"
                        frmMain.txtDrainageM(i).Text = 3
                    Case "PLEASANT"
                        frmMain.txtDrainageM(i).Text = 4
                    Case "MACHIAS"
                        frmMain.txtDrainageM(i).Text = 5
                    Case "EAST MACHIAS"
                        frmMain.txtDrainageM(i).Text = 6
                    Case "DENNYS"
                        frmMain.txtDrainageM(i).Text = 7
                End Select
                
                frmMain.txtPShare(i).Text = Format(Exp(-(pwPSA(i + 1, maleFinal(z)))), "0.000")
                'pFinal = pFinal + Format(Exp(-(pwPSA(i + 1, maleFinal(z)))), "0.000")
                frmMain.txtOptimized(i).Text = "NO"
                z = z + 1
            End If
        Next i
        
        'frmMain.lblAvgPShare.Visible = True
        'frmMain.txtAvgPShare.Text = Format((pFinal / femMD), "0.000")
        'frmMain.txtAvgPShare.Visible = True
        
        For i = femMD To 9
            frmMain.lblFem(i).Visible = False
            frmMain.txtFem(i).Visible = False
            frmMain.txtFem(i).Locked = False
            frmMain.txtFem(i).Text = ""
            frmMain.txtFem(i).Locked = True
            frmMain.lblMale(i).Visible = False
            frmMain.txtMale(i).Visible = False
            frmMain.txtMale(i).Locked = False
            frmMain.txtMale(i).Text = ""
            frmMain.txtMale(i).Locked = True
            frmMain.txtPShare(i).Visible = False
            frmMain.cmdComment(i).Visible = False
            frmMain.imgFlag(i).Visible = False
            frmMain.chkSpawned(i).Visible = False
            frmMain.chkReleaseF(i).Visible = False
            frmMain.chkReleaseM(i).Visible = False
        Next i
        
        With dbsNew
            Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMating.* FROM tblBroodMating", dbOpenDynaset)
        
            For m = 1 To numF
                If frmMain.txtFem(m - 1).BackColor <> &HFF& Then
                    If CInt(frmMain.txtDrainageF(m - 1).Text) <> CInt(frmMain.txtDrainageM(m - 1).Text) Then
                        flaGG(m - 1) = "The male and female for this mating come from two different drainages." & Chr(13) & Chr(13)
                    End If
                    
                    With rstTemp
                        rstTemp.MoveFirst
                        Do Until rstTemp.EOF
                            If UCase(rstTemp!Dam) = UCase(frmMain.txtFem(m - 1).Text) Then
                                If UCase(rstTemp!Sire) = UCase(frmMain.txtMale(m - 1).Text) Then
                                    flaGG(m - 1) = flaGG(m - 1) & "This female and male pairing has already been spawned on " & rstTemp!TakeDate & Chr(13) & Chr(13)
                                End If
                            End If
                            rstTemp.MoveNext
                        Loop
                    End With
                
                    If flaGG(m - 1) <> Empty Then
                        frmMain.imgFlag(m - 1).Visible = True
                        frmMain.txtFlag(m - 1).Text = flaGG(m - 1)
                    Else
                        frmMain.imgFlag(m - 1).Visible = False
                    End If
                End If
            Next m
        End With
        
        dbsNew.Close
        wrkJet.Close
    End If
    
    pFinal = 0
    For i = 0 To femMD - 1
        pFinal = pFinal + CSng(frmMain.txtPShare(i).Text)
    Next i
    
    frmMain.lblAvgPShare.Visible = True
    frmMain.txtAvgPShare.Text = Format((pFinal / femMD), "0.000")
    frmMain.txtAvgPShare.Visible = True
    
    Call frmMain.highlightTag
    frmInput.Hide
End Sub

Private Sub cmdSexChange_Click()
    i = 1
    Do Until Mid(frmInput.txtTagCur2.Text, i, 1) = ""
        i = i + 1
    Loop
    
    If frmInput.txtTagCur2.Text <> "" Then
        Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
        Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
        With dbsNew
            Set rstTemp = dbsNew.OpenRecordset("SELECT tblBrood.* FROM tblBrood", dbOpenDynaset)
            With rstTemp
                rstTemp.MoveFirst
                Do Until rstTemp.EOF
                    If UCase(!Mark) = UCase(frmInput.txtTagCur2.Text) Then
reenterSex:
                        Msg = "Please enter in a sex for this tag." & Chr(13) & Chr(13) & "'M' for male, 'F' for female"
                        Title = "Gender Update"
                        Default = "M"
                        Response = InputBox(Msg, Title, Default)
                            
                        If Response = "M" Or Response = "m" Or Response = "F" Or Response = "f" Then
                            rstTemp.Edit
                            !Gender = UCase(Response)
                            rstTemp.Update
                            If !Gender = "M" Then
                                frmInput.txtSex.Text = "Male"
                            Else
                                frmInput.txtSex.Text = "Female"
                            End If
                        ElseIf Response = "" Then
                            Exit Sub
                        Else
                            GoTo reenterSex
                        End If
                        GoTo tagFound
                    End If
                    rstTemp.MoveNext
                Loop
                GoTo tagNotFound
tagFound:
            End With
        End With
        
        dbsNew.Close
        wrkJet.Close
    End If
    
    Exit Sub
    
tagNotFound:
    Msg = "The tag " & UCase(frmInput.txtTagCur2.Text) & " is not located in the 'tblbrood' table."
    Style = vbOKOnly + vbCritical + vbDefaultButton1
    Title = "Tag Not Found"
    Response = MsgBox(Msg, Style, Title)
    
End Sub

Private Sub Form_Activate()
    If frmMain.optMatDes(0).Value = True Then
        mateNum = 1
    ElseIf frmMain.optMatDes(1).Value = True Or frmMain.optMatDes(2).Value = True Then
        mateNum = 2
    Else
        mateNum = 3
    End If
    
    numF = CInt(frmMain.cmbFem.Text) * mateNum
    
    For i = 0 To numF - 1
        If i Mod mateNum = 0 Then
            frmInput.txtFamID(i).Visible = True
            frmInput.txtTagFem(i).Visible = True
            frmInput.cmdClearF(i).Visible = True
            'frmInput.txtPicFem(i).Visible = True
            'frmInput.txtWeightFem(i).Visible = True
        Else
            frmInput.txtFamID(i).Visible = True 'only show family id
            frmInput.txtTagFem(i).Visible = False
            frmInput.txtTagFem(i).Locked = False
            frmInput.txtTagFem(i).Text = ""
            frmInput.txtTagFem(i).Locked = True
            frmInput.cmdClearF(i).Visible = False
        End If
        
        If frmMain.txtFem(i).BackColor = &HFF& Or frmMain.txtFem(i).BackColor = &H80FF& Then
            tmpBi = 0
            For j = 0 To numF - 1
                If frmMain.txtFem(j).Text = frmMain.txtFem(i).Text Then
                    If frmMain.txtFem(j).BackColor = &HFF& Or frmMain.txtFem(j).BackColor = &H80FF& Then
                        tmpBi = tmpBi + 1
                    End If
                End If
            Next j
                
            frmInput.lblFem(i).Enabled = False
            frmInput.txtFamID(i).Enabled = False
            If mateNum = 1 Or (mateNum = 2 And tmpBi = 2) Or (mateNum = 3 And tmpBi = 3) Then
                frmInput.txtTagFem(i).Enabled = False
                frmInput.cmdClearF(i).Enabled = False
            Else
                frmInput.txtTagFem(i).Enabled = True
                frmInput.cmdClearF(i).Enabled = True
            End If
            'frmInput.txtPicFem(i).Enabled = False
            'frmInput.txtWeightFem(i).Enabled = False
        Else
            frmInput.lblFem(i).Enabled = True
            frmInput.txtFamID(i).Enabled = True
            frmInput.txtTagFem(i).Enabled = True
            frmInput.cmdClearF(i).Enabled = True
            'frmInput.txtPicFem(i).Enabled = True
            'frmInput.txtWeightFem(i).Enabled = True
        End If
    Next i
    
    For i = numF To 9
        frmInput.lblFem(i).Visible = False
        frmInput.txtFamID(i).Visible = False
        frmInput.txtTagFem(i).Visible = False
        frmInput.cmdClearF(i).Visible = False
        'frmInput.txtPicFem(i).Visible = False
        'frmInput.txtWeightFem(i).Visible = False
    Next i
    
    For i = 0 To CInt(frmMain.cmbMale.Text) - 1
        If frmInput.chkUse(i).Value = 1 Then
            frmInput.chkUse(i).Visible = True
            frmInput.lblMale(i).Visible = True
            frmInput.txtTagMale(i).Visible = True
            frmInput.cmdClearM(i).Visible = True
            'frmInput.txtPicMale(i).Visible = True
            'frmInput.txtWeightMale(i).Visible = True
            
            frmInput.chkUse(i).Enabled = True
            frmInput.lblMale(i).Enabled = True
            frmInput.txtTagMale(i).Enabled = True
            frmInput.cmdClearM(i).Enabled = True
            'frmInput.txtPicMale(i).Enabled = True
            'frmInput.txtWeightMale(i).Enabled = True
        End If
    Next i
        
    For i = 0 To CInt(frmMain.txtMale.Count) - 1
        If frmMain.txtMale(i).BackColor = &HFF& Then
            For j = 0 To CInt(frmMain.cmbMale.Text) - 1
                If UCase(frmInput.txtTagMale(j).Text) = UCase(frmMain.txtMale(i).Text) Then
                    'frmInput.chkUse(j).Value = 0
                    frmInput.chkUse(j).Enabled = False
                    frmInput.lblMale(j).Enabled = False
                    frmInput.txtTagMale(j).Enabled = False
                    frmInput.cmdClearM(j).Enabled = False
                    'frmInput.txtPicMale(j).Enabled = False
                    'frmInput.txtWeightMale(j).Enabled = False
                End If
            Next j
        Else
            For j = 0 To CInt(frmMain.cmbMale.Text) - 1
                If UCase(frmInput.txtTagMale(j).Text) = UCase(frmMain.txtMale(i).Text) Then
                    frmInput.chkUse(j).Value = 1
                    frmInput.chkUse(j).Enabled = True
                    frmInput.lblMale(j).Enabled = True
                    frmInput.txtTagMale(j).Enabled = True
                    frmInput.cmdClearM(j).Enabled = True
                    'frmInput.txtPicMale(j).Enabled = True
                    'frmInput.txtWeightMale(j).Enabled = True
                    
                    GoTo skipEnable
                End If
            Next j
            
            For j = 0 To CInt(frmMain.cmbMale.Text) - 1
                enTrans = "YES"
                For z = 0 To CInt(frmMain.cmbMale.Text) - 1
                    If UCase(frmInput.txtTagMale(j).Text) = UCase(frmMain.txtMale(z).Text) Then
                        enTrans = "NO"
                    End If
                Next z
                
                If frmInput.chkUse(j).Enabled = False And enTrans = "YES" Then
                    frmInput.chkUse(j).Value = 1
                    frmInput.chkUse(j).Enabled = True
                    frmInput.lblMale(j).Enabled = True
                    frmInput.txtTagMale(j).Enabled = True
                    frmInput.cmdClearM(j).Enabled = True
                    'frmInput.txtPicMale(j).Enabled = True
                    'frmInput.txtWeightMale(j).Enabled = True
                End If
            Next j
skipEnable:
        End If
    Next i
                        
    For i = CInt(frmMain.cmbMale.Text) To 10
        frmInput.chkUse(i).Visible = False
        frmInput.lblMale(i).Visible = False
        frmInput.txtTagMale(i).Visible = False
        frmInput.cmdClearM(i).Visible = False
        'frmInput.txtPicMale(i).Visible = False
        'frmInput.txtWeightMale(i).Visible = False
        
        frmInput.chkUse(i).Enabled = True
        frmInput.lblMale(i).Enabled = True
        frmInput.txtTagMale(i).Enabled = True
        frmInput.cmdClearM(i).Enabled = True
        'frmInput.txtPicMale(i).Enabled = True
        'frmInput.txtWeightMale(i).Enabled = True
    Next i
    
    'frmInput.txtTagCur.LinkMode = 1
End Sub

Private Sub txtFamID_LostFocus(Index As Integer)
    For i = 0 To UCase(frmMain.cmbFem.Text) - 1
        If i <> Index Then
            If UCase(frmInput.txtFamID(i).Text) = UCase(frmInput.txtFamID(Index).Text) And frmInput.txtFamID(Index).Text <> "" Then
                Msg = "The family ID " & frmInput.txtFamID(Index).Text & " is already present in female " & i + 1 & "'s family ID." & Chr(13) & "Please enter a different family ID."
                Style = vbOKOnly + vbCritical + vbDefaultButton1
                Title = "Duplicate Family ID"
                Response = MsgBox(Msg, Style, Title)
                
                frmInput.txtFamID(Index).Text = ""
                frmInput.txtFamID(Index).SetFocus
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub txtTagCur_Change()
    If frmMain.txtPitCur.Text <> " No ID Found" And frmMain.txtPitCur.Text <> " LOOKING" And frmMain.txtPitCur.Text <> " Low Battery" And frmMain.txtPitCur.Text <> "AVID/FECAVA/ISO" Then
        If frmMain.chkPrefix.Value = 1 Then
            frmInput.txtTagCur2.Text = frmInput.txtTagCur.Text
        Else
            frmInput.txtTagCur2.Text = Right(frmInput.txtTagCur.Text, 10)
        End If
    Else
        frmInput.txtTagCur2.Text = ""
        
        frmInput.txtSex.Locked = False
        frmInput.txtSex.Text = ""
        frmInput.txtSex.Locked = True
    End If
End Sub

Private Sub txtTagCur2_Change()
    k = 1
    Do Until Mid(frmInput.txtTagCur2.Text, k, 1) = ""
        k = k + 1
    Loop
    
    frmInput.txtTagCur2.Text = UCase(frmInput.txtTagCur2.Text)
    frmInput.txtTagCur2.SelStart = k - 1
    
    If frmInput.txtTagCur2.Text <> "" And ((k = 11 And Mid(txtTagCur2.Text, 4, 1) <> ".") Or Left(Right(txtTagCur2.Text, 11), 1) = ".") And frmInput.txtTagCur2.Text <> tempTag Then
        Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
        Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
        
        With dbsNew
            Set rstCull = dbsNew.OpenRecordset("SELECT tblBroodLookup.* FROM tblBroodLookup", dbOpenDynaset)
            With rstCull
                If rstCull.RecordCount > 0 Then
                    rstCull.MoveFirst
                    Do Until rstCull.EOF
                        If UCase(rstCull!Mark) = UCase(frmInput.txtTagCur2.Text) Then
                            Msg = rstCull!Comments
                            Style = vbOKOnly + vbExclamation + vbDefaultButton1
                            Title = "Warning"
                            Response = MsgBox(Msg, Style, Title)
                            
                        End If
                        rstCull.MoveNext
                    Loop
                End If
            End With
    
            Set rstTemp = dbsNew.OpenRecordset("SELECT tblBrood.* FROM tblBrood", dbOpenDynaset)
            With rstTemp
                rstTemp.MoveFirst
                Do Until rstTemp.EOF
reTry:
                    If UCase(!Mark) = UCase(frmInput.txtTagCur2.Text) Then
                        If !Gender = "M" Then
                            frmInput.txtSex.Locked = False
                            frmInput.txtSex.Text = "Male"
                            frmInput.txtSex.Locked = True
                        ElseIf !Gender = "F" Then
                            frmInput.txtSex.Locked = False
                            frmInput.txtSex.Text = "Female"
                            frmInput.txtSex.Locked = True
                        Else
enterSex:
                            Msg = "There is no gender entered for this tag, please enter in a sex." & Chr(13) & Chr(13) & "'M' for male, 'F' for female"
                            Title = "Gender Identification"
                            Default = "M"
                            Response = InputBox(Msg, Title, Default)
                            
                            If Response = "M" Or Response = "m" Or Response = "F" Or Response = "f" Then
                                rstTemp.Edit
                                !Gender = UCase(Response)
                                rstTemp.Update
                                GoTo reTry
                            ElseIf Response = "" Then
                                GoTo inputCancel
                            Else
                                Msg = "Please enter in either a 'M' or an 'F' for the gender input." & Chr(13) & Chr(13) & "If you do not know the gender please use a different individual."
                                Style = vbOKOnly + vbInformation + vbDefaultButton1
                                Title = "Invalid Gender"
                                Response = MsgBox(Msg, Style, Title)
                                GoTo reTry
                            End If
                        End If
                        
                        Select Case !Drainage
                            Case 1
                                frmInput.txtDrainage.Text = "Penobscot"
                            Case 11
                                frmInput.txtDrainage.Text = "Sheepscot"
                            Case 3
                                frmInput.txtDrainage.Text = "NARRAGUAGUS"
                            Case 4
                                frmInput.txtDrainage.Text = "Pleasant"
                            Case 5
                                frmInput.txtDrainage.Text = "Machias"
                            Case 6
                                frmInput.txtDrainage.Text = "East Machias"
                            Case 7
                                frmInput.txtDrainage.Text = "Dennys"
                            Case Else
                                frmInput.txtDrainage.Text = !Drainage
                        End Select
                        
                        frmInput.txtYear = rstTemp!CaptureYear
                        
                        GoTo Finished
                    End If
                    rstTemp.MoveNext
                Loop
                Msg = "The tag " & UCase(frmInput.txtTagCur2.Text) & " is not present in the 'tblbrood' table." & Chr(13) & "Do you wish to add it?"
                Style = vbYesNo + vbInformation + vbDefaultButton1
                Title = "Tag Not Present"
                Response = MsgBox(Msg, Style, Title)
                
                If Response = 6 Then 'yes
                    frmAddTag.Show 1
                Else
                    Msg = "Do you still wish to use this tag without entering it into 'tblBrood'?"
                    Style = vbYesNo + vbInformation + vbDefaultButton1
                    Title = "Tag Fate"
                    Response = MsgBox(Msg, Style, Title)
                    
                    If Response = 6 Then 'yes
                        frmPartialAddTag.Show 1
                    Else
inputCancel:
                        Msg = "Please scan another individual."
                        Style = vbOKOnly + vbInformation + vbDefaultButton1
                        Title = "Scan Different Individual"
                        Response = MsgBox(Msg, Style, Title)
                    
                        frmInput.txtTagCur2.Text = ""
                        frmInput.txtTagCur2.SetFocus
                        GoTo Finished2
                    End If
                End If
            End With
            
Finished:
            Set rstTemp = dbsNew.OpenRecordset("SELECT tblBroodMating.* From tblBroodMating ORDER BY tblBroodMating.TakeDate DESC , tblBroodMating.TakeTime DESC", dbOpenDynaset)
            
            With rstTemp
                rstTemp.MoveFirst
                priorSpawn = 0
                frmInput.lstXSpawned.Clear
                If frmInput.txtSex.Text = "Female" Then
                    Do Until rstTemp.EOF
                        If UCase(!Dam) = UCase(frmInput.txtTagCur2.Text) Then
                            priorSpawn = priorSpawn + 1
                            frmInput.lstXSpawned.AddItem priorSpawn & "   " & Format(!TakeDate, "MM/DD/YYYY")
                        End If
                        rstTemp.MoveNext
                    Loop
                End If
                
                If frmInput.txtSex.Text = "Male" Then
                    Do Until rstTemp.EOF
                        If UCase(!Sire) = UCase(frmInput.txtTagCur2.Text) Then
                            priorSpawn = priorSpawn + 1
                            frmInput.lstXSpawned.AddItem priorSpawn & "   " & Format(!TakeDate, "MM/DD/YYYY")
                        End If
                        rstTemp.MoveNext
                    Loop
                End If
            End With
        End With
        
Finished2:
        dbsNew.Close
        wrkJet.Close
    ElseIf tempTag <> UCase(frmInput.txtTagCur2.Text) Then
        frmInput.txtSex.Text = ""
        frmInput.txtDrainage.Text = ""
        frmInput.txtYear.Text = ""
        frmInput.lstXSpawned.Clear
    End If
    
    tempTag = UCase(frmInput.txtTagCur2.Text)
    
End Sub
