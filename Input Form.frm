VERSION 5.00
Begin VB.Form frmInput 
   Caption         =   "Individual Input"
   ClientHeight    =   8535
   ClientLeft      =   915
   ClientTop       =   1245
   ClientWidth     =   13695
   Icon            =   "Input Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFamIDPrefix 
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
      Height          =   390
      Left            =   5040
      MouseIcon       =   "Input Form.frx":048A
      MousePointer    =   99  'Custom
      TabIndex        =   113
      ToolTipText     =   $"Input Form.frx":05DC
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CheckBox chkFamIDPrefix 
      Caption         =   "Use Family ID Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      MouseIcon       =   "Input Form.frx":0697
      MousePointer    =   99  'Custom
      TabIndex        =   112
      ToolTipText     =   $"Input Form.frx":07E9
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton cmdCohortChange 
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
      Height          =   405
      Left            =   2760
      MouseIcon       =   "Input Form.frx":089B
      MousePointer    =   99  'Custom
      TabIndex        =   111
      TabStop         =   0   'False
      ToolTipText     =   "Update the cohort of the current individual"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdPopChange 
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
      Height          =   405
      Left            =   2760
      MouseIcon       =   "Input Form.frx":09ED
      MousePointer    =   99  'Custom
      TabIndex        =   110
      TabStop         =   0   'False
      ToolTipText     =   "Update the population of the current individual"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdClearTemplate 
      Caption         =   "Clear Templates"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      MouseIcon       =   "Input Form.frx":0B3F
      MousePointer    =   99  'Custom
      TabIndex        =   109
      ToolTipText     =   "Clear the list of ID template options"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cmbTemplate 
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
      Height          =   390
      Left            =   120
      MouseIcon       =   "Input Form.frx":0C91
      MousePointer    =   99  'Custom
      TabIndex        =   108
      ToolTipText     =   $"Input Form.frx":0DE3
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CheckBox chkTemplate 
      Caption         =   "Use ID Input Template"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "Input Form.frx":0E9E
      MousePointer    =   99  'Custom
      TabIndex        =   107
      ToolTipText     =   $"Input Form.frx":0FF0
      Top             =   960
      Width           =   2655
   End
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
      Left            =   2760
      MouseIcon       =   "Input Form.frx":10A2
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Remove the entered ID"
      Top             =   480
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
      Height          =   405
      Left            =   2760
      MouseIcon       =   "Input Form.frx":11F4
      MousePointer    =   99  'Custom
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Update the gender of the current individual"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   10
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1346
      MousePointer    =   99  'Custom
      TabIndex        =   102
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   7080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   9
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1498
      MousePointer    =   99  'Custom
      TabIndex        =   101
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   6480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   8
      Left            =   9240
      MouseIcon       =   "Input Form.frx":15EA
      MousePointer    =   99  'Custom
      TabIndex        =   100
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   7
      Left            =   9240
      MouseIcon       =   "Input Form.frx":173C
      MousePointer    =   99  'Custom
      TabIndex        =   99
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   5280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   6
      Left            =   9240
      MouseIcon       =   "Input Form.frx":188E
      MousePointer    =   99  'Custom
      TabIndex        =   98
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   4680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   5
      Left            =   9240
      MouseIcon       =   "Input Form.frx":19E0
      MousePointer    =   99  'Custom
      TabIndex        =   97
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   4080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   4
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1B32
      MousePointer    =   99  'Custom
      TabIndex        =   96
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   3480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   3
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1C84
      MousePointer    =   99  'Custom
      TabIndex        =   95
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   2
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1DD6
      MousePointer    =   99  'Custom
      TabIndex        =   94
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   2280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   9240
      MouseIcon       =   "Input Form.frx":1F28
      MousePointer    =   99  'Custom
      TabIndex        =   93
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   1680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkUse 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   0
      Left            =   9240
      MouseIcon       =   "Input Form.frx":207A
      MousePointer    =   99  'Custom
      TabIndex        =   92
      ToolTipText     =   "If checked, then male is available for creation of mating pairs. If not checked, male will be ignored."
      Top             =   1080
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
      Left            =   120
      MouseIcon       =   "Input Form.frx":21CC
      MousePointer    =   99  'Custom
      TabIndex        =   38
      ToolTipText     =   "Close this form"
      Top             =   7920
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
      Height          =   270
      Left            =   9000
      MouseIcon       =   "Input Form.frx":231E
      MousePointer    =   99  'Custom
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   $"Input Form.frx":2470
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.ListBox lstXSpawned 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "List of prior mating dates and mates for the current individual"
      Top             =   5760
      Width           =   3495
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
      Left            =   4200
      TabIndex        =   22
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   20
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   18
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4200
      TabIndex        =   4
      ToolTipText     =   "Family ID number associated with the female"
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDrainage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Population of the current individual"
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox txtTagCur2 
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Most recently scanned/entered individual ID"
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdAddCurTag 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Current ID"
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
      Left            =   840
      MouseIcon       =   "Input Form.frx":2514
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Add the current ID to the first open female or male slot"
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox txtSex 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Gender of the current individual"
      Top             =   2880
      Width           =   2655
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
      Top             =   480
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
      Left            =   8160
      MouseIcon       =   "Input Form.frx":2666
      MousePointer    =   99  'Custom
      TabIndex        =   37
      ToolTipText     =   "Clears all data for mating individuals"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpt 
      BackColor       =   &H0080FF80&
      Caption         =   "Proceed"
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
      MouseIcon       =   "Input Form.frx":27B8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Create mating pairs based on the 'Optimize Matings' setting"
      Top             =   7920
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":290A
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   7080
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   7080
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":2A5C
      MousePointer    =   99  'Custom
      TabIndex        =   33
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   6480
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   6480
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":2BAE
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   5880
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   5880
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":2D00
      MousePointer    =   99  'Custom
      TabIndex        =   31
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   5280
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   5280
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":2E52
      MousePointer    =   99  'Custom
      TabIndex        =   30
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   4680
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   4680
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":2FA4
      MousePointer    =   99  'Custom
      TabIndex        =   29
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   4080
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   4080
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":30F6
      MousePointer    =   99  'Custom
      TabIndex        =   28
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   3480
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   3480
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":3248
      MousePointer    =   99  'Custom
      TabIndex        =   27
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   2880
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   2880
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":339A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   2280
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   2280
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":34EC
      MousePointer    =   99  'Custom
      TabIndex        =   25
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   1680
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   1680
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
      Left            =   12480
      MouseIcon       =   "Input Form.frx":363E
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   "Remove this male's unique ID"
      Top             =   1080
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating male"
      Top             =   1080
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3790
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   6480
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   6480
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":38E2
      MousePointer    =   99  'Custom
      TabIndex        =   21
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   5880
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   5880
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3A34
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   5280
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   5280
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3B86
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   4680
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   4680
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3CD8
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   4080
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   4080
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3E2A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   3480
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   3480
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":3F7C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   2880
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   2880
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":40CE
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   2280
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   2280
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":4220
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   1680
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   1680
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
      Left            =   7680
      MouseIcon       =   "Input Form.frx":4372
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Remove this female's unique ID"
      Top             =   1080
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID for mating female"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   105
      TabStop         =   0   'False
      ToolTipText     =   "Cohort of the current individual"
      Top             =   4800
      Width           =   2655
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
      Caption         =   "Cohort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   104
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   91
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      X1              =   8805
      X2              =   8805
      Y1              =   600
      Y2              =   6840
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   90
      ToolTipText     =   "3 digit ID number for female-male spawning"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Population"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   89
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Prior Matings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   87
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   84
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
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
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   82
      ToolTipText     =   "PIT tag number for spawning male"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Females"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   81
      ToolTipText     =   "PIT tag number for spawning female"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblMale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
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
      Index           =   10
      Left            =   9720
      TabIndex        =   79
      Top             =   7140
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   9
      Left            =   9720
      TabIndex        =   77
      Top             =   6540
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   8
      Left            =   9720
      TabIndex        =   75
      Top             =   5940
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   7
      Left            =   9720
      TabIndex        =   73
      Top             =   5340
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   6
      Left            =   9720
      TabIndex        =   71
      Top             =   4740
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   5
      Left            =   9720
      TabIndex        =   69
      Top             =   4140
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   67
      Top             =   3540
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   3
      Left            =   9720
      TabIndex        =   65
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   63
      Top             =   2340
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   61
      Top             =   1740
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   59
      Top             =   1140
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   57
      Top             =   6540
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   55
      Top             =   5940
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   53
      Top             =   5340
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   51
      Top             =   4740
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   49
      Top             =   4140
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   47
      Top             =   3540
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   45
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   43
      Top             =   2340
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   41
      Top             =   1740
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3840
      TabIndex        =   39
      Top             =   1140
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   7095
      Left            =   3840
      Top             =   480
      Width           =   9735
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
Dim pFinalMeasure() As Single, strFactorial As String, maleNonOpt() As Integer
Dim pTotal As Single, pFinal As Single, pShared As Long
Dim pMeasure As Single, totAlleles As Long, maleOrder() As Long
Dim priorSpawn As Long, femScored() As Long, maleScored() As Long, locScored As Long
Dim tempTag As String, fPic() As String, fWeight() As Single, mPic() As String
Dim mWeight() As Single, mDrainage(11) As String, fDrainage(10) As String
Dim rstCull As Recordset, flaGG() As Variant, numDrainage As Long, disAbledF As Long
Dim disAbledM As Long, enNumF() As Long, enNumM() As Long, tmpEnNumM() As Long, enableCnt As Long
Dim mYear(11) As Long, fYear(10) As Long, z As Long, enTrans As String
Dim tmpMOrder() As String, pwPSA() As Single, numFem As Long
Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long
Dim femMD As Integer, numF As Integer, numM As Integer, mateNum As Integer, matDes As Integer
Dim maleMD As Integer, lociCnt As Long, firstAllele As Long, idCol As Long, minLoci As Long
Dim pPartial As Single, tmpBi As Integer, popRelVal() As Variant, mateArray() As Variant
Dim reasonArray() As Variant, tmpStr As String, popArray() As Variant
Dim accessApp As Access.Application, dbTemp As Database

Private Sub chkFamIDPrefix_Click()
    If frmInput.cmbFamIDPrefix.Enabled = False Then
        frmInput.cmbFamIDPrefix.Enabled = True
    Else
        frmInput.cmbFamIDPrefix.Enabled = False
    End If
End Sub

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

Private Sub chkTemplate_Click()
    Call frmInput.toggleTemplate
End Sub

Private Sub chkUse_Click(Index As Integer)
    For z = 0 To (CInt(frmMain.cmbMale.Text) - 1)
        If frmInput.chkUse(z).Enabled = False Or frmInput.chkUse(z).Value = 0 Then
            frmInput.lblMale(z).Enabled = False
            frmInput.txtTagMale(z).Enabled = False
            frmInput.cmdClearM(z).Enabled = False
        Else
            frmInput.lblMale(z).Enabled = True
            frmInput.txtTagMale(z).Enabled = True
            frmInput.cmdClearM(z).Enabled = True
        End If
    Next z
End Sub

Private Sub cmbTemplate_LostFocus()
    For i = 0 To frmInput.cmbTemplate.ListCount - 1
        If frmInput.cmbTemplate.Text = frmInput.cmbTemplate.List(i) Then
            Exit Sub
        End If
    Next i
    If frmInput.cmbTemplate.Text <> "" Then
        frmInput.cmbTemplate.AddItem frmInput.cmbTemplate.Text
        frmInput.cmbTemplate.ListIndex = frmInput.cmbTemplate.ListCount - 1
    End If
End Sub

Private Sub cmdAddCurTag_Click()
    If frmInput.chkTemplate.Value = 1 Then
        i = frmInput.checkTemplate
            
        If i = -1 Then
            Msg = "The current ID does not match the input template, please finish entering the ID or change the template."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "ID Error"
            Response = MsgBox(Msg, Style, Title)
            
            frmInput.txtTagCur2.SetFocus
            Exit Sub
        End If
    End If
    
    If UCase(frmInput.txtSex.Text) = "MALE" Then
        For i = 0 To CInt(frmMain.cmbMale.Text) - 1
            If UCase(frmInput.txtTagCur2.Text) = UCase(frmInput.txtTagMale(i).Text) Then
                Msg = "This tag number is already present in male number " & i + 1 & "." & Chr(13) & "Do you wish to add it again?"
                Style = vbYesNo + vbInformation + vbDefaultButton1
                Title = "Duplicate Tag"
                Response = MsgBox(Msg, Style, Title)
                
                If Response = 7 Then Exit Sub
            End If
        Next i
        
        For i = 0 To 10
            If frmInput.txtTagMale(i).Visible = True Then
                If frmInput.txtTagMale(i).Text = "" Then
                    frmInput.txtTagMale(i).Text = UCase(frmInput.txtTagCur2.Text)
                    frmInput.cmdClearM(i).Tag = UCase(frmInput.txtTagCur2.Text)
                    mDrainage(i) = frmInput.txtDrainage.Text
                    If frmInput.txtYear.Text <> "" Then
                        mYear(i) = CInt(frmInput.txtYear.Text)
                    End If
                    GoTo clearData
                End If
            End If
        Next i
        Msg = "There are no empty spaces to add an additional MALE tag number."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Input Error"
        Response = MsgBox(Msg, Style, Title)
        
    ElseIf UCase(frmInput.txtSex.Text) = "FEMALE" Then
        For i = 0 To CInt(frmMain.cmbFem.Text) - 1
            If UCase(frmInput.txtTagCur2.Text) = UCase(frmInput.txtTagFem(i).Text) Then
                Msg = "This tag number is already present in female number " & i + 1 & "." & Chr(13) & "Do you wish to add it again?"
                Style = vbYesNo + vbInformation + vbDefaultButton1
                Title = "Duplicate Tag"
                Response = MsgBox(Msg, Style, Title)
                
                If Response = 7 Then Exit Sub
            End If
        Next i
        
        For i = 0 To 9
            If frmInput.txtTagFem(i).Visible = True Then
                If frmInput.txtTagFem(i).Text = "" Then
                    frmInput.txtTagFem(i).Text = UCase(frmInput.txtTagCur2.Text)
                    frmInput.cmdClearF(i).Tag = UCase(frmInput.txtTagCur2.Text)
                    For b = 0 To mateNum - 1
                        fDrainage(i + b) = frmInput.txtDrainage.Text
                        If frmInput.txtYear.Text <> "" Then
                            fYear(i + b) = CInt(frmInput.txtYear.Text)
                        End If
                    Next b
                    GoTo clearData
                End If
            End If
        Next i
        Msg = "There are no empty spaces to add an additional FEMALE tag number."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
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
    
    'dbsNew.Close
    'accessApp.CloseCurrentDatabase
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
        If frmInput.cmdClearF(i).Enabled = True Then
            frmInput.txtFamID(i).Text = ""
            frmInput.txtTagFem(i).Text = ""
        End If
    Next i
   
    For j = 0 To CInt(frmMain.cmbMale.Text) - 1
        If frmInput.cmdClearM(j).Enabled = True Then
            frmInput.chkUse(j).Value = 1
            frmInput.txtTagMale(j).Text = ""
        End If
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
    
    frmInput.txtTagFem(Index).Text = ""
End Sub

Private Sub cmdClearM_Click(Index As Integer)
    If frmInput.txtTagMale(Index).Text <> "" Then
        Msg = "Would you like to provide a reason for clearing this individual?"
        Style = vbYesNo + vbQuestion + vbDefaultButton1
        Title = "Reason for Clearing"
        Response = MsgBox(Msg, Style, Title)
    
        If Response = 6 Then 'yes
            Dim tmpTag As String
            Dim secondForm As New frmAddReason
            tmpTag = frmInput.txtTagMale(Index).Text
            secondForm.tmpTagReason = tmpTag
            secondForm.Show 1
        End If
    End If
    
    frmInput.txtTagMale(Index).Text = ""
End Sub

Private Sub cmdClearTag_Click()
    frmInput.txtTagCur2.Text = ""
    frmInput.txtTagCur2.SetFocus
End Sub

Public Function getAlleles(rstTemp)
    lociCnt = CLng(frmDataSpec.txtLoci.Text)
    firstAllele = frmDataSpec.cmbFirstLocus.ListIndex
    idCol = frmDataSpec.cmbUniqueId.ListIndex
    
    ReDim fAlleles(femMD - disAbledF, lociCnt, 2)
    ReDim mAlleles(numM - disAbledM, lociCnt, 2)
    ReDim femFinal(femMD - disAbledF)
    ReDim maleFinal(femMD - disAbledF)
    ReDim pTempMeasure(femMD - disAbledF)
    ReDim pFinalMeasure(femMD - disAbledF)
    ReDim maleOrder(numM - disAbledM)
    ReDim femScored(femMD - disAbledF)
    ReDim maleNonOpt(femMD - disAbledF)
    ReDim maleScored(numM - disAbledM)
    
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveLast
        rstTemp.MoveFirst
    End If
    
    If rstTemp.RecordCount > 0 Then
        popArray = rstTemp.GetRows(rstTemp.RecordCount)
        
        'acquiring male and female allele values
        For j = 1 To femMD - disAbledF
            For a = 0 To UBound(popArray, 2)
                If UCase(popArray(idCol, a)) = UCase(fOrder(j)) Then
                    m = 1
                    locScored = 0
                    For k = firstAllele To (firstAllele + (lociCnt * 2) - 1)
                        fAlleles(j, m, 1) = popArray(k, a)
                        If IsNull(popArray(k, a)) = False Then locScored = locScored + 1
                        k = k + 1
                        fAlleles(j, m, 2) = popArray(k, a)
                        If IsNull(popArray(k, a)) = False Then locScored = locScored + 1
                        m = m + 1
                    Next k
                    femScored(j) = locScored / 2
                    Exit For
                End If
            Next a
        Next j
                
        For j = 1 To numM - disAbledM
            For a = 0 To UBound(popArray, 2)
                If UCase(popArray(idCol, a)) = UCase(mOrder(j)) Then
                    m = 1
                    locScored = 0
                    For k = firstAllele To (firstAllele + (lociCnt * 2) - 1)
                        mAlleles(j, m, 1) = popArray(k, a)
                        If IsNull(popArray(k, a)) = False Then locScored = locScored + 1
                        k = k + 1
                        mAlleles(j, m, 2) = popArray(k, a)
                        If IsNull(popArray(k, a)) = False Then locScored = locScored + 1
                        m = m + 1
                    Next k
                    maleScored(j) = locScored / 2
                    Exit For
                End If
            Next a
        Next j
        
        Erase popArray
    End If
End Function

Public Function propSA()
    'calculating proportion of shared alleles
    ReDim pwPSA(femMD - disAbledF, numM - disAbledM)
    
    For f = 1 To femMD - disAbledF
        For m = 1 To numM - disAbledM
            pTotal = 0
            totAlleles = lociCnt * 2
            pShared = 0
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
        If pTotal > pFinal Or pTotal = 0 Then
            pFinal = pTotal
            For n = 1 To femMD - disAbledF
                maleFinal(n) = arr(n - 1)
                pFinalMeasure(n) = Format(Exp(-(pTempMeasure(n))), "0.000")
            Next n
        End If
    Next
    
    For n = 1 To femMD - disAbledF
        femFinal(n) = n
        maleNonOpt(n) = n
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
                        pTotal = pTotal + pwPSA(arrF(n - 1 + j), arr(i - 1))
                        pTempMeasure(n + j) = pwPSA(arrF(n - 1 + j), arr(i - 1))
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
    
    For n = 1 To femMD - disAbledF
        j = 0
        For k = 1 To mateNum
            b = (Int(n / ((mateNum ^ 2) + 1))) + j + (Int(n / ((mateNum ^ 2) + 1) + 1))
            maleNonOpt(n + j) = b
            'pFinalMeasure(n + j) = Format(Exp(-(pwPSA((n + j), b))), "0.000")
            j = j + 1
        Next k
        n = n + (mateNum - 1)
    Next n
End Function

Private Sub cmdClearTemplate_Click()
    Msg = "Would you like to remove all of the ID templates from the dropdown list?"
    Style = vbYesNo + vbQuestion + vbDefaultButton1
    Title = "Clear Template List"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = 6 Then 'yes
        frmInput.cmbTemplate.Clear
    End If
End Sub

Private Sub cmdCohortChange_Click()
    If frmInput.txtTagCur2.Text <> "" Then
        Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "])='" & frmInput.txtTagCur2.Text & "'));", dbOpenDynaset)
        rstTemp.MoveLast
        rstTemp.MoveFirst
        If rstTemp.RecordCount > 0 Then
reenterCohort:
            Msg = "Please enter in a cohort for this ID (must be a number)."
            Title = "Cohort Update"
            Default = Format(Date, "YYYY")
            Response = InputBox(Msg, Title, Default)
                                
            If IsNumeric(Response) = True Then
                rstTemp.Edit
                rstTemp(frmDataSpec.cmbCohort.Text) = CSng(Response)
                rstTemp.Update
                frmInput.txtYear.Text = Response
            ElseIf Response = "" Then
                Exit Sub
            Else
                GoTo reenterCohort
            End If
        Else
            Msg = "The ID '" & UCase(frmInput.txtTagCur2.Text) & "' is not located in the '" & frmDataSpec.cmbGeneticsTable.Text & "' table so it will just be updated here."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "ID Not Found"
            Response = MsgBox(Msg, Style, Title)
        
reenterCohort2:
            Msg = "Please enter in a cohort for this ID."
            Title = "Cohort Update"
            Default = Format(Date, "YYYY")
            Response = InputBox(Msg, Title, Default)
                                
            If IsNumeric(Response) = True Then
                frmInput.txtYear.Text = Response
            ElseIf Response = "" Then
                Exit Sub
            Else
                GoTo reenterCohort2
            End If
        End If
    End If
End Sub

Public Sub cmdOpt_Click()
    Call frmInput.saveTemplates
    
    disAbledF = 0
    disAbledM = 0
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
    
    If matDes = 2 Or matDes = 4 Then
        maleMD = numM * mateNum
    Else
        maleMD = numM
    End If
        
    For i = 0 To femMD - 1
        If frmInput.txtTagFem(i).Visible = True And frmInput.txtTagFem(i).Text = "" Then
            Msg = "There is an empty space in the Female list." + Chr(13) + "Please either add a female or reselect the appropriate number of females."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "Missing Female"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
        
        If frmInput.txtFamID(i).Text = "" Then
            Msg = "There is an empty space in the Family ID list." + Chr(13) + "Please enter in the appropriate family identification number."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "Missing Family ID"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
            
        If frmInput.txtFamID(i).Enabled = False Then
            disAbledF = disAbledF + 1
        End If
    Next i
        
    If disAbledF = femMD Then
        Msg = "There are no enabled female tags to optimize." + Chr(13) + "Please cancel this form and either deselect a mating-pair or reselect the number of females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "No Females to Optimize"
        Response = MsgBox(Msg, Style, Title)
            
        Exit Sub
    End If
                
    For i = 0 To numM - 1
        If frmInput.txtTagMale(i).Text = "" Then
            Msg = "There is an empty space in the Male list." + Chr(13) + "Please either add a male or reselect the appropriate number of males."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "Missing Male"
            Response = MsgBox(Msg, Style, Title)
            
            Exit Sub
        End If
            
        If frmInput.txtTagMale(i).Enabled = False Then
            disAbledM = disAbledM + 1
        End If
    Next i
                    
    If femMD - disAbledF > maleMD - disAbledM Then
        Msg = "There must be at least as many enabled males as enabled females."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Not Enough Enabled Males"
        Response = MsgBox(Msg, Style, Title)
            
        Exit Sub
    End If
    
    ReDim famID(femMD - disAbledF) As String
    ReDim fOrder(femMD - disAbledF) As String
    ReDim enNumF(femMD - disAbledF) As Long
    
    ReDim flaGG(femMD - disAbledF) As Variant
    
    'read in female tags and fam ids
    tmpStr = ""
    j = 1
    For i = 1 To femMD
        For m = 1 To mateNum
            If frmInput.txtFamID((i - 1) + (m - 1)).Enabled = True Then
                enNumF(j) = (i - 1) + (m - 1)
                famID(j) = frmInput.txtFamID((i - 1) + (m - 1)).Text
                fOrder(j) = frmInput.txtTagFem(i - 1).Text
                'fDrainage(j - 1) = fDrainage((i - 1))
                'fYear(j - 1) = fYear((i - 1))
                If tmpStr = "" Then
                    tmpStr = "([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "] = '" & fOrder(j) & "'"
                Else
                    tmpStr = tmpStr & " OR [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "] = '" & fOrder(j) & "'"
                End If
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
            tmpStr = tmpStr & " OR [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "] = '" & mOrder(j) & "'"
            j = j + 1
        End If
    Next i
    
    tmpStr = tmpStr & ")"
    
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE " & tmpStr & ";", dbOpenDynaset)
    
    If frmInput.chkOptimize.Value = 1 Then
        '***Read in mating individuals genotypes
        Call getAlleles(rstTemp)
        
        '***Get proportion of shared alleles (add other relatedness options here)
        Select Case frmMain.cmbRelMetric.Text
            Case "Proportion of Shared Alleles"
                Call propSA
        End Select
        
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
        
        
        m = 1
        If frmFlags.chkNoAlleles.Value = 1 Then
            minLoci = CLng(frmFlags.txtMinLoci.Text)
        End If
        
        For p = 1 To femMD
            If frmInput.txtFamID(p - 1).Enabled = True Then
                frmMain.lblFem(enNumF(m)).Caption = famID(femFinal(m))
                frmMain.txtFem(enNumF(m)).Text = fOrder(femFinal(m))
                frmMain.txtFem(enNumF(m)).Tag = fOrder(m)
                frmMain.txtLocFem(enNumF(m)).Text = femScored(femFinal(m))
                frmMain.txtLocFem(enNumF(m)).Tag = femScored(m)
                frmMain.txtDrainageF(enNumF(m)).Text = fDrainage(enNumF(femFinal(m)))
                frmMain.txtDrainageF(enNumF(m)).Tag = fDrainage(enNumF(m))
                frmMain.txtfYear(enNumF(m)).Text = fYear(enNumF(femFinal(m)))
                frmMain.txtfYear(enNumF(m)).Tag = fYear(enNumF(m))
                
                frmMain.lblMale(enNumF(m)).Caption = enNumM(maleFinal(m)) + 1
                frmMain.txtMale(enNumF(m)).Text = mOrder(maleFinal(m))
                frmMain.txtMale(enNumF(m)).Tag = mOrder(maleNonOpt(m))
                frmMain.txtLocMale(enNumF(m)).Text = maleScored(maleFinal(m))
                frmMain.txtLocMale(enNumF(m)).Tag = maleScored(maleNonOpt(m))
                frmMain.txtDrainageM(enNumF(m)).Text = mDrainage(enNumM(maleFinal(m)))
                frmMain.txtDrainageM(enNumF(m)).Tag = mDrainage(enNumM(maleNonOpt(m)))
                frmMain.txtmYear(enNumF(m)).Text = mYear(enNumM(maleFinal(m)))
                frmMain.txtmYear(enNumF(m)).Tag = mYear(enNumM(maleNonOpt(m)))
                
                If femScored(femFinal(m)) = 0 Or maleScored(maleFinal(m)) = 0 Then
                    frmMain.txtPShare(enNumF(m)).Text = "NA"
                Else
                    frmMain.txtPShare(enNumF(m)).Text = Format(pFinalMeasure(m), "0.000")
                End If
                
                If femScored(femFinal(m)) = 0 Or maleScored(maleNonOpt(m)) = 0 Then
                    frmMain.txtPShare(enNumF(m)).Tag = "NA"
                Else
                    frmMain.txtPShare(enNumF(m)).Tag = Format(Exp(-(pwPSA(m, maleNonOpt(m)))), "0.000")
                End If
                
                If (disAbledF > 0 And disAbledM > 0) Then
                    frmMain.txtOptimized(enNumF(m)) = "REOPTIMIZED"
                Else
                    frmMain.txtOptimized(enNumF(m)) = "YES"
                End If
                
                
                'Determine mating flags
                
                'Exceeds relatedness threshold
                If frmFlags.chkUppQuart.Value = 1 Then
                    If femScored(femFinal(m)) > 0 And maleScored(maleFinal(m)) > 0 Then
                        tmpBi = 0
                        For a = 0 To UBound(popRelVal, 1)
                            If fDrainage(enNumF(femFinal(m))) = popRelVal(a, 0) Then
                                tmpBi = 1
                                Exit For
                            End If
                        Next a
                        
                        If tmpBi = 1 Then
                            If pFinalMeasure(m) > CSng(popRelVal(a, 1)) Then
                                flaGG(m - 1) = flaGG(m - 1) & "The pair-wise genetic relatedness for this mating (" & pFinalMeasure(m) & ") exceeds the threshold value for the " & fDrainage(enNumF(m)) & " population (>" & popRelVal(a, 1) & ")." & Chr(13) & Chr(13)
                            End If
                        End If
                    End If
                End If
                
                'Below minimum number of scored loci
                If frmFlags.chkNoAlleles.Value = 1 Then
                    If femScored(femFinal(m)) < minLoci Then
                        flaGG(m - 1) = flaGG(m - 1) & "This female has been scored for " & femScored(femFinal(m)) & " loci, less than the specified minimum number (" & minLoci & ")." & Chr(13) & Chr(13)
                    End If
                    
                    If maleScored(maleFinal(m)) < minLoci Then
                        flaGG(m - 1) = flaGG(m - 1) & "This male has been scored for " & maleScored(maleFinal(m)) & " loci, less than the specified minimum number (" & minLoci & ")." & Chr(13) & Chr(13)
                    End If
                End If
                
                'Previous mating
                If frmFlags.chkPrevMating.Value = 1 Then
                    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbMatingsTable.Text & "].* FROM [" & frmDataSpec.cmbMatingsTable.Text & "] WHERE ([" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbDamID.Text & "] = '" & fOrder(femFinal(m)) & "' AND [" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbSireID.Text & "] = '" & mOrder(maleFinal(m)) & "');", dbOpenDynaset)
                    
                    If rstTemp.RecordCount > 0 Then
                        rstTemp.MoveLast
                        rstTemp.MoveFirst
                        Do Until rstTemp.EOF
                            flaGG(m - 1) = flaGG(m - 1) & "This female and male pairing was previously mated on " & rstTemp(frmDataSpec.cmbDate.Text) & Chr(13) & Chr(13)
                            rstTemp.MoveNext
                        Loop
                    End If
                    Set rstTemp = Nothing
                End If
                
                'Different pops
                If frmDataSpec.chkPop.Value = 1 And frmFlags.chkDiffPops.Value = 1 Then
                    If UCase(fDrainage(enNumF(femFinal(m)))) <> UCase(mDrainage(enNumM(maleFinal(m)))) Then
                        flaGG(m - 1) = flaGG(m - 1) & "The female (" & fDrainage(enNumF(femFinal(m))) & ") and male (" & mDrainage(enNumM(maleFinal(m))) & ") for this mating pair come from two different populations." & Chr(13) & Chr(13)
                    End If
                End If
                
                'Different cohorts
                If frmDataSpec.chkCohort.Value = 1 And frmFlags.chkDiffCohorts.Value = 1 Then
                    If UCase(fYear(enNumF(femFinal(m)))) <> UCase(mYear(enNumM(maleFinal(m)))) Then
                        flaGG(m - 1) = flaGG(m - 1) & "The female (" & fYear(enNumF(femFinal(m))) & ") and male (" & mYear(enNumM(maleFinal(m))) & ") for this mating pair are from two different cohorts." & Chr(13) & Chr(13)
                    End If
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
    Else '***No optimization, use straight order
        '***Read in mating individuals genotypes
        Call getAlleles(rstTemp)
                
        '***Get proportion of shared alleles (add other relatedness options here)
        Select Case frmMain.cmbRelMetric.Text
            Case "Proportion of Shared Alleles"
                Call propSA
        End Select
                
        '***Adjust mOrder for 2 to 2 and 3 to 3 mating designs
        If matDes = 2 Or matDes = 4 Then
            For n = 1 To femMD - disAbledF
                j = 0
                For k = 1 To mateNum
                    b = (Int(n / ((mateNum ^ 2) + 1))) + j + (Int(n / ((mateNum ^ 2) + 1) + 1))
                    femFinal(n + j) = enNumF(n) + 1
                    maleFinal(n + j) = enNumM(b)
                    pFinalMeasure(n + j) = Format(Exp(-(pwPSA((n + j), b))), "0.000")
                    j = j + 1
                Next k
                n = n + (mateNum - 1)
            Next n
            
            ReDim Preserve mOrder(femMD)
            ReDim Preserve maleScored(femMD)
            tmpEnNumM = enNumM
            ReDim enNumM(femMD)
            For n = femMD To 1 Step -1
                For i = 1 To numM
                    If tmpEnNumM(i) = maleFinal(n) Then
                        enNumM(n) = tmpEnNumM(i)
                        maleFinal(n) = i
                        'mOrder(n) = mOrder(i)
                        'maleScored(n) = maleScored(i)
                        Exit For
                    End If
                Next i
            Next n
        Else
            For n = 1 To femMD - disAbledF
                maleFinal(n) = n
                femFinal(n) = n
                pFinalMeasure(n) = Format(Exp(-(pwPSA(n, n))), "0.000")
            Next n
        End If
        
        
        
        z = 1
        For i = 0 To femMD - 1
            frmMain.lblFem(i).Visible = True
            frmMain.txtFem(i).Visible = True
            frmMain.lblMale(i).Visible = True
            frmMain.txtMale(i).Visible = True
            frmMain.txtPShare(i).Visible = True
            frmMain.cmdComment(i).Visible = True
            frmMain.chkSpawned(i).Visible = True
            frmMain.chkReleaseF(i).Visible = True
            frmMain.chkReleaseM(i).Visible = True
            
            If frmMain.txtFem(i).BackColor <> &HFF& Then
                frmMain.lblFem(i).Caption = famID(z)
                frmMain.txtFem(i).Text = fOrder(z)
                frmMain.txtFem(i).Tag = fOrder(z)
                frmMain.txtLocFem(i).Text = femScored(z)
                frmMain.txtLocFem(i).Tag = femScored(z)
                frmMain.txtDrainageF(i).Text = fDrainage(enNumF(z))
                frmMain.txtDrainageF(i).Tag = fDrainage(enNumF(z))
                frmMain.txtfYear(i).Text = fYear(enNumF(z))
                frmMain.txtfYear(i).Tag = fYear(enNumF(z))
                
                frmMain.lblMale(i).Caption = frmInput.lblMale(maleFinal(z)).Caption
                frmMain.txtMale(i).Text = mOrder(maleFinal(z))
                frmMain.txtMale(i).Tag = mOrder(maleFinal(z))
                frmMain.txtLocMale(i).Text = maleScored(maleFinal(z))
                frmMain.txtLocMale(i).Tag = maleScored(maleFinal(z))
                frmMain.txtDrainageM(i).Text = mDrainage(enNumM(z))
                frmMain.txtDrainageM(i).Tag = mDrainage(enNumM(z))
                frmMain.txtmYear(i).Text = mYear(enNumM(z))
                frmMain.txtmYear(i).Tag = mYear(enNumM(z))
                                
                If femScored(z) = 0 Or maleScored(maleFinal(z)) = 0 Then
                    frmMain.txtPShare(i).Text = "NA"
                    frmMain.txtPShare(i).Tag = "NA"
                Else
                    frmMain.txtPShare(i).Text = Format(pFinalMeasure(z), "0.000")
                    frmMain.txtPShare(i).Tag = Format(pFinalMeasure(z), "0.000")
                End If
                frmMain.txtOptimized(i).Text = "NO"
                z = z + 1
            End If
        Next i
                
        For i = femMD To 9
            frmMain.lblFem(i).Visible = False
            frmMain.txtFem(i).Visible = False
            frmMain.txtFem(i).Text = ""
            frmMain.lblMale(i).Visible = False
            frmMain.txtMale(i).Visible = False
            frmMain.txtMale(i).Text = ""
            frmMain.txtPShare(i).Visible = False
            frmMain.cmdComment(i).Visible = False
            frmMain.imgFlag(i).Visible = False
            frmMain.chkSpawned(i).Visible = False
            frmMain.chkReleaseF(i).Visible = False
            frmMain.chkReleaseM(i).Visible = False
        Next i
                
        'Determine mating flags
        For m = 1 To femMD
            If frmMain.txtFem(m - 1).BackColor <> &HFF& Then  'Check logic*************************************************************************
                'Exceeds relatedness threshold
                If frmFlags.chkUppQuart.Value = 1 Then
                    If femScored(m) > 0 And maleScored(m) > 0 Then
                        tmpBi = 0
                        For a = 0 To UBound(popRelVal, 1)
                            If fDrainage(enNumF(m)) = popRelVal(a, 0) Then
                                tmpBi = 1
                                Exit For
                            End If
                        Next a
                        
                        If tmpBi = 1 Then
                            If pFinalMeasure(m) > CSng(popRelVal(a, 1)) Then
                                flaGG(m - 1) = flaGG(m - 1) & "The pair-wise genetic relatedness for this mating (" & pFinalMeasure(m) & ") exceeds the threshold value for the " & fDrainage(enNumF(m)) & " population (>" & popRelVal(a, 1) & ")." & Chr(13) & Chr(13)
                            End If
                        End If
                    End If
                End If
                
                'Below minimum number of scored loci
                If frmFlags.chkNoAlleles.Value = 1 Then
                    If femScored(m) < minLoci Then
                        flaGG(m - 1) = flaGG(m - 1) & "This female has been scored for " & femScored(m) & " loci, less than the specified minimum number (" & minLoci & ")." & Chr(13) & Chr(13)
                    End If
                    
                    If maleScored(maleFinal(m)) < minLoci Then
                        flaGG(m - 1) = flaGG(m - 1) & "This male has been scored for " & maleScored(maleFinal(m)) & " loci, less than the specified minimum number (" & minLoci & ")." & Chr(13) & Chr(13)
                    End If
                End If
                
                'Previous mating
                If frmFlags.chkPrevMating.Value = 1 Then
                    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbMatingsTable.Text & "].* FROM [" & frmDataSpec.cmbMatingsTable.Text & "] WHERE ([" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbDamID.Text & "] = '" & fOrder(m) & "' AND [" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbDamID.Text & "] = '" & mOrder(maleFinal(m)) & "');", dbOpenDynaset)
                    
                    If rstTemp.RecordCount > 0 Then
                        rstTemp.MoveLast
                        rstTemp.MoveFirst
                        Do Until rstTemp.EOF
                            flaGG(m - 1) = flaGG(m - 1) & "This female and male pairing was previously mated on " & rstTemp(frmDataSpec.cmbDate.Text) & Chr(13) & Chr(13)
                            rstTemp.MoveNext
                        Loop
                    End If
                    Set rstTemp = Nothing
                End If
                
                'Different pops
                If frmDataSpec.chkPop.Value = 1 And frmFlags.chkDiffPops.Value = 1 Then
                    If UCase(fDrainage(enNumF(m))) <> UCase(mDrainage(enNumM(m))) Then
                        flaGG(m - 1) = flaGG(m - 1) & "The female (" & fDrainage(enNumF(m)) & ") and male (" & mDrainage(enNumM(m)) & ") for this mating pair come from two different populations." & Chr(13) & Chr(13)
                    End If
                End If
                
                'Different cohorts
                If frmDataSpec.chkCohort.Value = 1 And frmFlags.chkDiffCohorts.Value = 1 Then
                    If UCase(fYear(enNumF(m))) <> UCase(mYear(enNumM(m))) Then
                        flaGG(m - 1) = flaGG(m - 1) & "The female (" & fYear(enNumF(m)) & ") and male (" & mYear(enNumM(m)) & ") for this mating pair are from two different cohorts." & Chr(13) & Chr(13)
                    End If
                End If
                
                If flaGG(m - 1) <> Empty Then
                    frmMain.imgFlag(m - 1).Visible = True
                    frmMain.txtFlag(m - 1).Text = flaGG(m - 1)
                Else
                    frmMain.imgFlag(m - 1).Visible = False
                End If
            End If
        Next m
    End If
    
    pFinal = 0
    b = 0
    For i = 0 To femMD - 1
        If (CLng(frmMain.txtLocFem(i).Text) > 0 And CLng(frmMain.txtLocMale(i).Text) > 0) Then
            b = b + 1
            pFinal = pFinal + CSng(frmMain.txtPShare(i).Text)
        End If
    Next i
    
    frmMain.lblAvgPShare.Visible = True
    If b = 0 Then
        frmMain.txtAvgPShare.Text = "NA"
    Else
        frmMain.txtAvgPShare.Text = Format((pFinal / b), "0.000")
    End If
    frmMain.txtAvgPShare.Visible = True
    
    'Call frmMain.highlightTag
    
    frmInput.Hide
End Sub

Private Sub cmdPopChange_Click()
    If frmInput.txtTagCur2.Text <> "" Then
        Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "])='" & frmInput.txtTagCur2.Text & "'));", dbOpenDynaset)
        rstTemp.MoveLast
        rstTemp.MoveFirst
        If rstTemp.RecordCount > 0 Then
            
            Msg = "Please enter in a population for this ID."
            Title = "Population Update"
            Default = ""
            Response = InputBox(Msg, Title, Default)
                                
            If UCase(Response) <> "" Then
                rstTemp.Edit
                rstTemp(frmDataSpec.cmbPop.Text) = Response
                rstTemp.Update
                frmInput.txtDrainage.Text = Response
            ElseIf Response = "" Then
                Exit Sub
            End If
        Else
            Msg = "The ID '" & UCase(frmInput.txtTagCur2.Text) & "' is not located in the '" & frmDataSpec.cmbGeneticsTable.Text & "' table so it will just be updated here."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "ID Not Found"
            Response = MsgBox(Msg, Style, Title)
            
            Msg = "Please enter in a population for this ID."
            Title = "Population Update"
            Default = ""
            Response = InputBox(Msg, Title, Default)
            
            If UCase(Response) <> "" Then
                frmInput.txtDrainage.Text = Response
            ElseIf Response = "" Then
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub cmdSexChange_Click()
    If frmInput.txtTagCur2.Text <> "" Then
        Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "])='" & frmInput.txtTagCur2.Text & "'));", dbOpenDynaset)
        rstTemp.MoveLast
        rstTemp.MoveFirst
        If rstTemp.RecordCount > 0 Then
reenterSex:
            Msg = "Please enter in a sex for this ID." & Chr(13) & Chr(13) & "'M' for male, 'F' for female"
            Title = "Gender Update"
            Default = "M"
            Response = InputBox(Msg, Title, Default)
                                
            If UCase(Response) = "M" Or UCase(Response) = "F" Then
                rstTemp.Edit
                rstTemp(frmDataSpec.cmbSex.Text) = UCase(Response)
                rstTemp.Update
                If UCase(Response) = "M" Then
                    frmInput.txtSex.Text = "Male"
                Else
                    frmInput.txtSex.Text = "Female"
                End If
            ElseIf Response = "" Then
                Exit Sub
            Else
                GoTo reenterSex
            End If
        Else
            Msg = "The ID '" & UCase(frmInput.txtTagCur2.Text) & "' is not located in the '" & frmDataSpec.cmbGeneticsTable.Text & "' table so it will just be updated here."
            Style = vbOKOnly + vbInformation + vbDefaultButton1
            Title = "ID Not Found"
            Response = MsgBox(Msg, Style, Title)
            
reenterSex2:
            Msg = "Please enter in a sex for this ID." & Chr(13) & Chr(13) & "'M' for male, 'F' for female"
            Title = "Gender Update"
            Default = "M"
            Response = InputBox(Msg, Title, Default)
                                
            If UCase(Response) = "M" Or UCase(Response) = "F" Then
                If UCase(Response) = "M" Then
                    frmInput.txtSex.Text = "Male"
                Else
                    frmInput.txtSex.Text = "Female"
                End If
            ElseIf Response = "" Then
                Exit Sub
            Else
                GoTo reenterSex2
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If frmDataSpec.chkPop.Value = 0 Then
        frmInput.txtDrainage.Enabled = False
        frmInput.cmdPopChange.Enabled = False
    Else
        frmInput.txtDrainage.Enabled = True
        frmInput.cmdPopChange.Enabled = True
    End If
    
    If frmDataSpec.chkCohort.Value = 0 Then
        frmInput.txtYear.Enabled = False
        frmInput.cmdCohortChange.Enabled = False
    Else
        frmInput.txtYear.Enabled = True
        frmInput.cmdCohortChange.Enabled = True
    End If
    
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
        Else
            frmInput.lblFem(i).Enabled = True
            frmInput.txtFamID(i).Enabled = True
            frmInput.txtTagFem(i).Enabled = True
            frmInput.cmdClearF(i).Enabled = True
        End If
    Next i
    
    For i = numF To 9
        frmInput.lblFem(i).Visible = False
        frmInput.txtFamID(i).Visible = False
        frmInput.txtTagFem(i).Visible = False
        frmInput.cmdClearF(i).Visible = False
    Next i
    
    For i = 0 To CInt(frmMain.cmbMale.Text) - 1
        If frmInput.chkUse(i).Value = 1 Then
            frmInput.chkUse(i).Visible = True
            frmInput.lblMale(i).Visible = True
            frmInput.txtTagMale(i).Visible = True
            frmInput.cmdClearM(i).Visible = True
            
            frmInput.chkUse(i).Enabled = True
            frmInput.lblMale(i).Enabled = True
            frmInput.txtTagMale(i).Enabled = True
            frmInput.cmdClearM(i).Enabled = True
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
        
        frmInput.chkUse(i).Enabled = True
        frmInput.lblMale(i).Enabled = True
        frmInput.txtTagMale(i).Enabled = True
        frmInput.cmdClearM(i).Enabled = True
    Next i
End Sub

Private Sub Form_Load()
    'Set accessApp = CreateObject("Access.Application")
    Call frmInput.importTemplates
End Sub

Private Sub txtFamID_LostFocus(Index As Integer)
    For i = 0 To numF - 1
        If i <> Index Then
            If UCase(frmInput.txtFamID(i).Text) = UCase(frmInput.txtFamID(Index).Text) And frmInput.txtFamID(Index).Text <> "" Then
                Msg = "The family ID " & frmInput.txtFamID(Index).Text & " is already present in female " & i + 1 & "'s family ID." & Chr(13) & "Please enter a different family ID."
                Style = vbOKOnly + vbInformation + vbDefaultButton1
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
    k = Len(frmInput.txtTagCur2.Text)
    
    frmInput.txtTagCur2.Text = UCase(frmInput.txtTagCur2.Text)
    frmInput.txtTagCur2.SelStart = k
    
    If frmInput.chkTemplate.Value = 1 Then
        tmpBi = frmInput.checkTemplate
    Else
        tmpBi = 1
    End If
    
    If frmInput.txtTagCur2.Text <> "" And frmInput.txtTagCur2.Text <> tempTag And tmpBi <> -1 Then 'And ((k = 11 And Mid(txtTagCur2.Text, 4, 1) <> ".") Or Left(Right(txtTagCur2.Text, 11), 1) = ".")
        Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "])='" & frmInput.txtTagCur2.Text & "'));", dbOpenDynaset)
        
        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveLast
            rstTemp.MoveFirst
            If frmDataSpec.chkFlag.Value = 1 Then
                If rstTemp(frmDataSpec.cmbFlag.Text).Value <> "" Then
                    Msg = rstTemp(frmDataSpec.cmbFlag.Text).Value
                    Style = vbOKOnly + vbExclamation + vbDefaultButton1
                    Title = "Flag Present"
                    Response = MsgBox(Msg, Style, Title)
                End If
            End If
        
reTry:
            If UCase(rstTemp(frmDataSpec.cmbSex.Text)) = "M" Or UCase(rstTemp(frmDataSpec.cmbSex.Text)) = "MALE" Then
                frmInput.txtSex.Text = "Male"
            ElseIf UCase(rstTemp(frmDataSpec.cmbSex.Text)) = "F" Or UCase(rstTemp(frmDataSpec.cmbSex.Text)) = "FEMALE" Then
                frmInput.txtSex.Text = "Female"
            Else
enterSex:
                Msg = "There is no gender entered for this tag, please enter in a sex." & Chr(13) & Chr(13) & "'F' for female, 'M' for male"
                Title = "Gender Identification"
                Default = "M"
                Response = InputBox(Msg, Title, Default)
                                    
                If UCase(Response) = "M" Or UCase(Response) = "F" Then
                    rstTemp.Edit
                    rstTemp(frmDataSpec.cmbSex.Text) = UCase(Response)
                    rstTemp.Update
                    GoTo reTry
                ElseIf Response = "" Then
                    GoTo inputCancel
                Else
                    Msg = "Please enter in either an 'F' for female or an 'M' for male for the gender input." & Chr(13) & Chr(13) & "If you do not know the gender please use a different individual."
                    Style = vbOKOnly + vbInformation + vbDefaultButton1
                    Title = "Invalid Gender"
                    Response = MsgBox(Msg, Style, Title)
                    GoTo reTry
                End If
            End If
        
            If frmDataSpec.chkPop.Value = 1 Then
                If IsNull(rstTemp(frmDataSpec.cmbPop.Text)) = False Then
                    frmInput.txtDrainage.Text = rstTemp(frmDataSpec.cmbPop.Text)
                End If
            End If
                                    
            If frmDataSpec.chkCohort.Value = 1 Then
                If IsNull(rstTemp(frmDataSpec.cmbCohort.Text)) = False Then
                    frmInput.txtYear = rstTemp(frmDataSpec.cmbCohort.Text)
                End If
            End If
                                
            GoTo Finished
        Else
            Msg = "The ID '" & UCase(frmInput.txtTagCur2.Text) & "' is not present in the genetics table (" & frmDataSpec.cmbGeneticsTable.Text & ")." & Chr(13) & Chr(13) & "Would you like to add this individual to the genetics table?"
            Style = vbYesNo + vbQuestion + vbDefaultButton1
            Title = "ID Not Present"
            Response = MsgBox(Msg, Style, Title)
                    
            If Response = 6 Then 'yes
                frmAddTag.Show 1
            Else
                Msg = "Do you still wish to use this tag without entering it into the genetics table?"
                Style = vbYesNo + vbQuestion + vbDefaultButton1
                Title = "ID Fate"
                Response = MsgBox(Msg, Style, Title)
                        
                If Response = 6 Then 'yes
                    frmPartialAddTag.Show 1
                Else
inputCancel:
                    Msg = "Please enter another ID."
                    Style = vbOKOnly + vbInformation + vbDefaultButton1
                    Title = "Enter Different Individual"
                    Response = MsgBox(Msg, Style, Title)
                        
                    'frmInput.txtTagCur2.Text = ""
                    frmInput.txtTagCur2.SetFocus
                    GoTo Finished2
                End If
            End If
        End If
                
Finished:
        Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbMatingsTable.Text & "].* FROM [" & frmDataSpec.cmbMatingsTable.Text & "] WHERE ((([" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbDamID.Text & "])='" & frmInput.txtTagCur2.Text & "')) OR ((([" & frmDataSpec.cmbMatingsTable.Text & "].[" & frmDataSpec.cmbSireID.Text & "])='" & frmInput.txtTagCur2.Text & "'));", dbOpenDynaset)
        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveLast
            rstTemp.MoveFirst
            mateArray = rstTemp.GetRows(rstTemp.RecordCount)
        
            frmInput.lstXSpawned.Clear
            For k = 0 To UBound(mateArray, 2)
                If frmInput.txtSex.Text = "Female" Then
                    frmInput.lstXSpawned.AddItem Format(mateArray(frmDataSpec.cmbDate.ListIndex, k), "MM/DD/YYYY") & "  " & mateArray(frmDataSpec.cmbSireID.ListIndex, k)
                ElseIf frmInput.txtSex.Text = "Male" Then
                    frmInput.lstXSpawned.AddItem Format(mateArray(frmDataSpec.cmbDate.ListIndex, k), "MM/DD/YYYY") & "  " & mateArray(frmDataSpec.cmbDamID.ListIndex, k)
                End If
            Next k
        End If
            
Finished2:
    ElseIf tempTag <> UCase(frmInput.txtTagCur2.Text) Then
        frmInput.txtSex.Text = ""
        frmInput.txtDrainage.Text = ""
        frmInput.txtYear.Text = ""
        frmInput.lstXSpawned.Clear
    End If
    
    tempTag = UCase(frmInput.txtTagCur2.Text)
End Sub

Public Function setPopRelVals(tmpCount, tmpArray)
    ReDim popRelVal(tmpCount, 1)
    
    For i = 0 To tmpCount
        popRelVal(i, 0) = tmpArray(i, 0)
        popRelVal(i, 1) = tmpArray(i, 1)
    Next i
End Function

Public Function fillReasons() As Variant
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbGenComments.Text & "] FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] GROUP BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbGenComments.Text & "] HAVING ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbGenComments.Text & "]) Is Not Null)) ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbGenComments.Text & "];", dbOpenDynaset)
    rstTemp.MoveLast
    rstTemp.MoveFirst
    fillReasons = rstTemp.GetRows(rstTemp.RecordCount)
End Function

Public Function addReason(tmpTag As String, tmpTagReason As String)
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbGenComments.Text & "] FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbUniqueId.Text & "])='" & tmpTag & "'));", dbOpenDynaset)
    rstTemp.MoveLast
    rstTemp.MoveFirst
    If rstTemp.RecordCount > 0 Then
        rstTemp.Edit
        If IsNull(rstTemp(frmDataSpec.cmbGenComments.Text)) = True Then
            rstTemp(frmDataSpec.cmbGenComments.Text) = tmpTagReason
        Else
            rstTemp(frmDataSpec.cmbGenComments.Text) = rstTemp(frmDataSpec.cmbGenComments.Text) & ", " & tmpTagReason
        End If
        rstTemp.Update
    Else
        MsgBox "ID '" & tmpTagReason & "' is not present in the '" & frmDataSpec.cmbGeneticsTable.Text & "' table, therefore a reason cannot be added.", vbInformation, "ID Not Found"
    End If
End Function

Public Function toggleTemplate()
    If frmInput.cmbTemplate.Enabled = False Then
        frmInput.cmbTemplate.Enabled = True
        frmInput.cmdClearTemplate.Enabled = True
    Else
        frmInput.cmbTemplate.Enabled = False
        frmInput.cmdClearTemplate.Enabled = False
    End If
End Function

Public Function saveTemplates()
    Open App.Path & "\template_settings.mmf" For Output As #1
        Print #1, CStr(frmInput.chkTemplate.Value)
        Print #1, CStr(frmInput.cmbTemplate.ListCount - 1)
        For i = 0 To frmInput.cmbTemplate.ListCount - 1
            Print #1, frmInput.cmbTemplate.List(i)
        Next i
        Print #1, frmInput.cmbTemplate.Text
    Close #1
End Function

Public Function importTemplates()
    If Dir(App.Path & "\template_settings.mmf") <> "" Then
        Open App.Path & "\template_settings.mmf" For Input As #1
            Input #1, tmpStr
            frmInput.chkTemplate.Value = CLng(tmpStr)
            
            Input #1, tmpStr
            a = CLng(tmpStr)
            For i = 0 To a
                Input #1, tmpStr
                frmInput.cmbTemplate.List(i) = tmpStr
            Next i
            
            Input #1, tmpStr
            For i = 0 To a
                If frmInput.cmbTemplate.List(i) = tmpStr Then
                    frmInput.cmbTemplate.ListIndex = i
                    Exit For
                End If
            Next i
        Close #1
    End If
End Function

Public Function checkTemplate() As Integer
    tmpStr = frmInput.cmbTemplate.Text
    checkTemplate = 1
    
    If frmInput.cmbTemplate.ListIndex <> -1 Then
        If Len(frmInput.txtTagCur2.Text) <> Len(tmpStr) Then
            checkTemplate = -1
        Else
            For i = 1 To Len(tmpStr)
                If Mid(tmpStr, i, 1) <> "#" Then
                    If Mid(frmInput.txtTagCur2.Text, i, 1) <> Mid(tmpStr, i, 1) Then
                        checkTemplate = -1
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
End Function

Public Function getPops(tmpType As String)
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] GROUP BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] HAVING ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "]) Is Not Null)) ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "];", dbOpenDynaset)
    rstTemp.MoveLast
    rstTemp.MoveFirst
    If rstTemp.RecordCount > 0 Then
        popArray = rstTemp.GetRows(rstTemp.RecordCount)
        For i = 0 To UBound(popArray, 2)
            If tmpType = "full" Then
                frmAddTag.cmbPop.List(i) = popArray(0, i)
            Else
                frmPartialAddTag.cmbPop.List(i) = popArray(0, i)
            End If
        Next i
    End If
    Set rstTemp = Nothing
End Function

Public Function addTag()
    Set rstTemp = dbsNew.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* FROM [" & frmDataSpec.cmbGeneticsTable.Text & "];", dbOpenDynaset)
            
    rstTemp.AddNew
        rstTemp(frmDataSpec.cmbUniqueId.Text) = UCase(frmAddTag.txtTagAdd.Text)
        rstTemp(frmDataSpec.cmbSex.Text) = UCase(Left(frmAddTag.cmbGender.Text, 1))
        If frmDataSpec.chkPop.Value = 1 Then
            rstTemp(frmDataSpec.cmbPop.Text) = frmAddTag.cmbPop.Text
        End If
        If frmDataSpec.chkCohort.Value = 1 Then
            rstTemp(frmDataSpec.cmbCohort.Text) = CLng(frmAddTag.txtCohort.Text)
        End If
        rstTemp(frmDataSpec.cmbComments.Text) = frmAddTag.txtComment.Text
    rstTemp.Update
End Function

Public Function setDBase()
    Set accessApp = CreateObject("Access.Application")
    On Error Resume Next
    accessApp.CloseCurrentDatabase
    accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
    On Error GoTo 0
    Set dbsNew = accessApp.CurrentDb
End Function

