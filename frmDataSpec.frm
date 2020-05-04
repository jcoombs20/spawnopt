VERSION 5.00
Begin VB.Form frmDataSpec 
   Caption         =   "Database Specifications"
   ClientHeight    =   8655
   ClientLeft      =   4485
   ClientTop       =   3525
   ClientWidth     =   20250
   Icon            =   "frmDataSpec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   20250
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   13080
      MouseIcon       =   "frmDataSpec.frx":048A
      MousePointer    =   99  'Custom
      Picture         =   "frmDataSpec.frx":05DC
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   1
      ToolTipText     =   "Click to refresh tables listed under the 'Genetics' and 'Matings' table options from the associated database"
      Top             =   120
      Width           =   405
   End
   Begin VB.TextBox txtMatingsNonopt 
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
      Left            =   9720
      TabIndex        =   71
      TabStop         =   0   'False
      Text            =   "matings_nonopt"
      ToolTipText     =   $"frmDataSpec.frx":0FAD
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtMatingsTable 
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
      Left            =   7080
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Table that will be created to store records of mating pairs and their associated information"
      Top             =   1080
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Frame frameMatings 
      Enabled         =   0   'False
      Height          =   6255
      Left            =   7080
      TabIndex        =   48
      Top             =   1560
      Width           =   12975
      Begin VB.ComboBox cmbMatDes 
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
         Left            =   1800
         MouseIcon       =   "frmDataSpec.frx":1036
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "-1"
         ToolTipText     =   "Column containing the genetic relatedness measure for the dam (mother) and sire (father) in the mating pair"
         Top             =   2160
         Width           =   4650
      End
      Begin VB.ComboBox cmbMetric 
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
         Left            =   1800
         MouseIcon       =   "frmDataSpec.frx":1188
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "-1"
         ToolTipText     =   "Column containing the metric used to calculate pairwise genetic relatedness"
         Top             =   1200
         Width           =   4650
      End
      Begin VB.CommandButton cmdClearMatings 
         Caption         =   "Clear Fields"
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
         MouseIcon       =   "frmDataSpec.frx":12DA
         MousePointer    =   99  'Custom
         TabIndex        =   38
         ToolTipText     =   "Click to clear the above fields"
         Top             =   5640
         Width           =   1575
      End
      Begin VB.ComboBox cmbComments 
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
         Left            =   8160
         MouseIcon       =   "frmDataSpec.frx":142C
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "-1"
         ToolTipText     =   "Column containing comments about this mating pair, including any mating flags that are raised by this pairing"
         Top             =   1680
         Width           =   4650
      End
      Begin VB.ComboBox cmbTime 
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
         Left            =   8160
         MouseIcon       =   "frmDataSpec.frx":157E
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "-1"
         ToolTipText     =   "Column containing the individual's gender (e.g. M, F, U)"
         Top             =   1200
         Width           =   4650
      End
      Begin VB.ComboBox cmbDate 
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
         Left            =   8160
         MouseIcon       =   "frmDataSpec.frx":16D0
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "-1"
         ToolTipText     =   "Column containing the date that the dam (mother) and sire (father) were mated"
         Top             =   720
         Width           =   4650
      End
      Begin VB.ComboBox cmbOptimized 
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
         Left            =   8160
         MouseIcon       =   "frmDataSpec.frx":1822
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "-1"
         ToolTipText     =   "Column recording wheter the mating pair were selected by minimizing genetic relatedness among all mating pairs in the batch"
         Top             =   240
         Width           =   4650
      End
      Begin VB.ComboBox cmbRelatedness 
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
         Left            =   1800
         MouseIcon       =   "frmDataSpec.frx":1974
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "-1"
         ToolTipText     =   "Column containing the genetic relatedness measure for the dam (mother) and sire (father) in the mating pair"
         Top             =   1680
         Width           =   4650
      End
      Begin VB.ComboBox cmbSireReleased 
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
         Left            =   8400
         MouseIcon       =   "frmDataSpec.frx":1AC6
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "-1"
         ToolTipText     =   "Column showing whether the sire (father) was released after mating"
         Top             =   5025
         Width           =   4410
      End
      Begin VB.ComboBox cmbSireLoci 
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
         Left            =   8400
         MouseIcon       =   "frmDataSpec.frx":1C18
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Tag             =   "-1"
         ToolTipText     =   "Column containing the number of loci that had allelic information for the sire (father)"
         Top             =   4545
         Width           =   4410
      End
      Begin VB.ComboBox cmbSireCohort 
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
         Left            =   8400
         MouseIcon       =   "frmDataSpec.frx":1D6A
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Tag             =   "-1"
         ToolTipText     =   "Column containing the cohort information (if available) for the sire (father)"
         Top             =   4065
         Width           =   4410
      End
      Begin VB.ComboBox cmbSirePop 
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
         Left            =   8400
         MouseIcon       =   "frmDataSpec.frx":1EBC
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "-1"
         ToolTipText     =   "Column containing the population information (if available) for the sire (father)"
         Top             =   3585
         Width           =   4410
      End
      Begin VB.ComboBox cmbSireID 
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
         Left            =   8400
         MouseIcon       =   "frmDataSpec.frx":200E
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "-1"
         ToolTipText     =   "Column containing the unique identerifier for the sire (father)"
         Top             =   3105
         Width           =   4410
      End
      Begin VB.ComboBox cmbDamReleased 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":2160
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Tag             =   "-1"
         ToolTipText     =   "Column showing whether the dam (mother) was released after mating"
         Top             =   5040
         Width           =   4410
      End
      Begin VB.ComboBox cmbBatch 
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
         Left            =   1800
         MouseIcon       =   "frmDataSpec.frx":22B2
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "-1"
         ToolTipText     =   $"frmDataSpec.frx":2404
         Top             =   240
         Width           =   4650
      End
      Begin VB.ComboBox cmbDamID 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":2491
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "-1"
         ToolTipText     =   "Column containing the unique identerifier for the dam (mother)"
         Top             =   3120
         Width           =   4410
      End
      Begin VB.ComboBox cmbFamID 
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
         Left            =   1800
         MouseIcon       =   "frmDataSpec.frx":25E3
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "-1"
         ToolTipText     =   "Column containing the identifier for all offspring from a mated pair"
         Top             =   720
         Width           =   4650
      End
      Begin VB.ComboBox cmbDamPop 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":2735
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "-1"
         ToolTipText     =   "Column containing the population information (if available) for the dam (mother)"
         Top             =   3600
         Width           =   4410
      End
      Begin VB.ComboBox cmbDamCohort 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":2887
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "-1"
         ToolTipText     =   "Column containing the cohort information (if available) for the dam (mother)"
         Top             =   4080
         Width           =   4410
      End
      Begin VB.ComboBox cmbDamLoci 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":29D9
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "-1"
         ToolTipText     =   "Column containing the number of loci that had allelic information for the dam (mother)"
         Top             =   4560
         Width           =   4410
      End
      Begin VB.Label lblMatDes 
         Alignment       =   1  'Right Justify
         Caption         =   "Mating Design"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label lblMetric 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblComments 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
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
         Left            =   6720
         TabIndex        =   69
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   68
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   67
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label lblOptimized 
         Alignment       =   1  'Right Justify
         Caption         =   "Optimized"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   66
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label lblRelatedness 
         Alignment       =   1  'Right Justify
         Caption         =   "Relatedness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   65
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label lblSireReleased 
         Alignment       =   1  'Right Justify
         Caption         =   "Released"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   64
         Top             =   5055
         Width           =   1815
      End
      Begin VB.Label lblSireLoci 
         Alignment       =   1  'Right Justify
         Caption         =   "Scored Loci"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   63
         Top             =   4575
         Width           =   1815
      End
      Begin VB.Label lblSireInfo 
         Caption         =   "Sire Information                                                      "
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
         Left            =   6840
         TabIndex        =   62
         Top             =   2745
         Width           =   6015
      End
      Begin VB.Label lblSireCohort 
         Alignment       =   1  'Right Justify
         Caption         =   "Cohort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   61
         Top             =   4095
         Width           =   1815
      End
      Begin VB.Label lblSirePop 
         Alignment       =   1  'Right Justify
         Caption         =   "Population"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   60
         Top             =   3615
         Width           =   1815
      End
      Begin VB.Label lblSireID 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   59
         Top             =   3135
         Width           =   1815
      End
      Begin VB.Label lblDamReleased 
         Alignment       =   1  'Right Justify
         Caption         =   "Released"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   5085
         Width           =   1815
      End
      Begin VB.Label lblDamLoci 
         Alignment       =   1  'Right Justify
         Caption         =   "Scored Loci"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   4605
         Width           =   1815
      End
      Begin VB.Label lblDamInfo 
         Caption         =   "Dam Information                                                         "
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
         Left            =   240
         TabIndex        =   56
         Top             =   2760
         Width           =   6255
      End
      Begin VB.Label lblDamCohort 
         Alignment       =   1  'Right Justify
         Caption         =   "Cohort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   4125
         Width           =   1815
      End
      Begin VB.Label lblBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "Batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label lblDamID 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   3165
         Width           =   1815
      End
      Begin VB.Label lblDamPop 
         Alignment       =   1  'Right Justify
         Caption         =   "Population"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   3645
         Width           =   1815
      End
      Begin VB.Label lblFamID 
         Alignment       =   1  'Right Justify
         Caption         =   "Family ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   765
         Width           =   1335
      End
   End
   Begin VB.OptionButton optMatingsTable 
      Caption         =   "Create New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   11760
      MouseIcon       =   "frmDataSpec.frx":2B2B
      MousePointer    =   99  'Custom
      TabIndex        =   16
      ToolTipText     =   "Select to create a new table for recording mated pairs"
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton optMatingsTable 
      Caption         =   "Open Existing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9600
      MouseIcon       =   "frmDataSpec.frx":2C7D
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Tag             =   "True"
      ToolTipText     =   "Select to map the below fields for recording mating pairs to columns present in an existing table"
      Top             =   720
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.ComboBox cmbMatingsTable 
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
      Left            =   7080
      MouseIcon       =   "frmDataSpec.frx":2DCF
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "-1"
      ToolTipText     =   "Table containing records of mating pairs and their associated information"
      Top             =   1080
      Width           =   6390
   End
   Begin VB.Frame frameGenetics 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   240
      TabIndex        =   44
      Top             =   1560
      Width           =   6375
      Begin VB.ComboBox cmbGenComments 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":2F21
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "-1"
         ToolTipText     =   $"frmDataSpec.frx":3073
         Top             =   2640
         Width           =   4170
      End
      Begin VB.CommandButton cmdClearGenetics 
         Caption         =   "Clear Fields"
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
         Left            =   2400
         MouseIcon       =   "frmDataSpec.frx":30FB
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Click to clear the above fields"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.ComboBox cmbUniqueId 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":324D
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "-1"
         ToolTipText     =   "Column containing the unique identifier for an individual (e.g. PIT tag number)"
         Top             =   240
         Width           =   4170
      End
      Begin VB.ComboBox cmbSex 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":339F
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "-1"
         ToolTipText     =   "Column containing the individual's gender (e.g. M, F, U)"
         Top             =   840
         Width           =   4170
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "           Flag"
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
         Left            =   480
         MouseIcon       =   "frmDataSpec.frx":34F1
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Tag             =   "0"
         ToolTipText     =   "Check to set and use a field containing a warning or comment about that individual that is displayed when used for mating"
         Top             =   4485
         Width           =   1455
      End
      Begin VB.ComboBox cmbFlag 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":3643
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "-1"
         ToolTipText     =   "Column in the table containing a warning or comment about the individual that is displayed when used for mating"
         Top             =   4440
         Width           =   4170
      End
      Begin VB.ComboBox cmbCohort 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":3795
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "-1"
         ToolTipText     =   "Column in the table containing the assigned cohort for an individual"
         Top             =   3840
         Width           =   4170
      End
      Begin VB.CheckBox chkCohort 
         Caption         =   "       Cohort"
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
         Left            =   480
         MouseIcon       =   "frmDataSpec.frx":38E7
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Tag             =   "0"
         ToolTipText     =   "Check to set and use a field containing an individual's assigned cohort"
         Top             =   3885
         Width           =   1455
      End
      Begin VB.ComboBox cmbPop 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":3A39
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "-1"
         ToolTipText     =   "Column in the table containing the assigned population for an individual"
         Top             =   3240
         Width           =   4170
      End
      Begin VB.CheckBox chkPop 
         Caption         =   "Population"
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
         Left            =   500
         MouseIcon       =   "frmDataSpec.frx":3B8B
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Tag             =   "0"
         ToolTipText     =   "Check to set and use a field containing an individual's assigned population"
         Top             =   3285
         Width           =   1435
      End
      Begin VB.ComboBox cmbFirstLocus 
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
         Left            =   2040
         MouseIcon       =   "frmDataSpec.frx":3CDD
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "-1"
         ToolTipText     =   $"frmDataSpec.frx":3E2F
         Top             =   2040
         Width           =   4170
      End
      Begin VB.TextBox txtLoci 
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
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "The total number of loci comprising an individual's genotype"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblGenComments 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   2685
         Width           =   1815
      End
      Begin VB.Label lblUniqueId 
         Alignment       =   1  'Right Justify
         Caption         =   "Unique Identifier"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   280
         Width           =   1815
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   880
         Width           =   1815
      End
      Begin VB.Label lblFirstLocus 
         Alignment       =   1  'Right Justify
         Caption         =   "First Locus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   2080
         Width           =   1815
      End
      Begin VB.Label lblLoci 
         Caption         =   "Number of Loci"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   1480
         Width           =   1815
      End
   End
   Begin VB.ComboBox cmbGeneticsTable 
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
      Left            =   240
      MouseIcon       =   "frmDataSpec.frx":3F0E
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "-1"
      ToolTipText     =   "Table containing the unique identifier and genetic information"
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox txtDBFile 
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   8295
   End
   Begin VB.CommandButton cmdDBFile 
      Caption         =   "..."
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
      Left            =   12480
      MouseIcon       =   "frmDataSpec.frx":4060
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Click to select the Access database containing individual genetic information and recorded matings if used previously "
      Top             =   120
      Width           =   495
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
      Left            =   240
      MouseIcon       =   "frmDataSpec.frx":41B2
      MousePointer    =   99  'Custom
      TabIndex        =   40
      ToolTipText     =   "Click to cancel changes made to the form"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FF80&
      Caption         =   "Save"
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
      Left            =   18720
      MouseIcon       =   "frmDataSpec.frx":4304
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Click to save the current database settings"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblMatingsNonopt 
      Alignment       =   2  'Center
      Caption         =   "matings_nonopt table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   72
      Top             =   7920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Matings Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   47
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Genetics Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   43
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mating Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   42
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmDataSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pSavefilename As SAVEFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type SAVEFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Dim OpenFile As OPENFILENAME
Dim SaveFile As SAVEFILENAME
Dim lReturn As Long
Dim sFilter As String

Dim tmpString As String, a As Long, accessApp As Access.Application, j As Long, i As Long
Dim dbTemp As Database, rstTemp As Recordset, tmpText As String, x As Long, tmpBi As Integer
Dim matTab As TableDef, tmpField As Field, tmpProp As Property, chkArray, txtArray
Dim Msg, Style, Title, Response, tmpBool As Boolean

Private Sub chkCohort_Click()
    Call toggleCohort
End Sub

Private Sub chkFlag_Click()
    Call toggleFlag
End Sub

Private Sub chkPop_Click()
    Call togglePop
End Sub

Private Sub cmbBatch_Click()
    If frmDataSpec.cmbBatch.ListIndex > -1 Then
        tmpText = checkDataType("cmbBatch")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbBatch.SetFocus
            End If
            frmDataSpec.cmbBatch.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbBatch.Text, "cmbBatch")
    End If
End Sub

Private Sub cmbCohort_Click()
    If frmDataSpec.cmbCohort.ListIndex > -1 Then
        tmpText = checkDataTypeGen("cmbCohort")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbCohort.SetFocus
            End If
            frmDataSpec.cmbCohort.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmbComments_Click()
    If frmDataSpec.cmbComments.ListIndex > -1 Then
        tmpText = checkDataType("cmbComments")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbComments.SetFocus
            End If
            frmDataSpec.cmbComments.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbComments.Text, "cmbComments")
    End If
End Sub

Private Sub cmbDamCohort_Click()
    If frmDataSpec.cmbDamCohort.ListIndex > -1 Then
        tmpText = checkDataType("cmbDamCohort")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDamCohort.SetFocus
            End If
            frmDataSpec.cmbDamCohort.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDamCohort.Text, "cmbDamCohort")
    End If
End Sub

Private Sub cmbDamID_Click()
    If frmDataSpec.cmbDamID.ListIndex > -1 Then
        tmpText = checkDataType("cmbDamID")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDamID.SetFocus
            End If
            frmDataSpec.cmbDamID.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDamID.Text, "cmbDamID")
    End If
End Sub

Private Sub cmbDamLoci_Click()
    If frmDataSpec.cmbDamLoci.ListIndex > -1 Then
        tmpText = checkDataType("cmbDamLoci")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDamLoci.SetFocus
            End If
            frmDataSpec.cmbDamLoci.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDamLoci.Text, "cmbDamLoci")
    End If
End Sub

Private Sub cmbDamPop_Click()
    If frmDataSpec.cmbDamPop.ListIndex > -1 Then
        tmpText = checkDataType("cmbDamPop")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDamPop.SetFocus
            End If
            frmDataSpec.cmbDamPop.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDamPop.Text, "cmbDamPop")
    End If
End Sub

Private Sub cmbDamReleased_Click()
    If frmDataSpec.cmbDamReleased.ListIndex > -1 Then
        tmpText = checkDataType("cmbDamReleased")
        If (tmpText <> "Yes/No") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be yes/no", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDamReleased.SetFocus
            End If
            frmDataSpec.cmbDamReleased.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDamReleased.Text, "cmbDamReleased")
    End If
End Sub

Private Sub cmbDate_Click()
    If frmDataSpec.cmbDate.ListIndex > -1 Then
        tmpText = checkDataType("cmbDate")
        If (tmpText <> "Date/Time") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be date/time", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbDate.SetFocus
            End If
            frmDataSpec.cmbDate.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbDate.Text, "cmbDate")
    End If
End Sub

Private Sub cmbFamID_Click()
    If frmDataSpec.cmbFamID.ListIndex > -1 Then
        tmpText = checkDataType("cmbFamID")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbFamID.SetFocus
            End If
            frmDataSpec.cmbFamID.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbFamID.Text, "cmbFamID")
    End If
End Sub

Private Sub cmbGeneticsTable_Click()
    Call enableGeneticOptions
End Sub

Private Sub cmbMatDes_Click()
    If frmDataSpec.cmbMatDes.ListIndex > -1 Then
        tmpText = checkDataType("cmbMatDes")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbMatDes.SetFocus
            End If
            frmDataSpec.cmbMatDes.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbMatDes.Text, "cmbMatDes")
    End If
End Sub

Private Sub cmbMatingsTable_Click()
    Call enableMatingsOptions
End Sub

Private Sub cmbMetric_Click()
    If frmDataSpec.cmbMetric.ListIndex > -1 Then
        tmpText = checkDataType("cmbMetric")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbMetric.SetFocus
            End If
            frmDataSpec.cmbMetric.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbMetric.Text, "cmbMetric")
    End If
End Sub

Private Sub cmbOptimized_Click()
    If frmDataSpec.cmbOptimized.ListIndex > -1 Then
        tmpText = checkDataType("cmbOptimized")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbOptimized.SetFocus
            End If
            frmDataSpec.cmbOptimized.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbOptimized.Text, "cmbOptimized")
    End If
End Sub

Private Sub cmbRelatedness_Click()
    If frmDataSpec.cmbRelatedness.ListIndex > -1 Then
        tmpText = checkDataType("cmbRelatedness")
        If (tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbRelatedness.SetFocus
            End If
            frmDataSpec.cmbRelatedness.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbRelatedness.Text, "cmbRelatedness")
    End If
End Sub

Private Sub cmbSireCohort_Click()
    If frmDataSpec.cmbSireCohort.ListIndex > -1 Then
        tmpText = checkDataType("cmbSireCohort")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbSireCohort.SetFocus
            End If
            frmDataSpec.cmbSireCohort.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbSireCohort.Text, "cmbSireCohort")
    End If
End Sub

Private Sub cmbSireID_Click()
    If frmDataSpec.cmbSireID.ListIndex > -1 Then
        tmpText = checkDataType("cmbSireID")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbSireID.SetFocus
            End If
            frmDataSpec.cmbSireID.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbSireID.Text, "cmbSireID")
    End If
End Sub

Private Sub cmbSireLoci_Click()
    If frmDataSpec.cmbSireLoci.ListIndex > -1 Then
        tmpText = checkDataType("cmbSireLoci")
        If (tmpText <> "Long Integer" And tmpText <> "Single" And tmpText <> "Double" And tmpText <> "Decimal") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be long integer, single, double, or decimal", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbSireLoci.SetFocus
            End If
            frmDataSpec.cmbSireLoci.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbSireLoci.Text, "cmbSireLoci")
    End If
End Sub

Private Sub cmbSirePop_Click()
    If frmDataSpec.cmbSirePop.ListIndex > -1 Then
        tmpText = checkDataType("cmbSirePop")
        If (tmpText <> "Text" And tmpText <> "Text (fixed width)") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be text", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbSirePop.SetFocus
            End If
            frmDataSpec.cmbSirePop.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbSirePop.Text, "cmbSirePop")
    End If
End Sub

Private Sub cmbSireReleased_Click()
    If frmDataSpec.cmbSireReleased.ListIndex > -1 Then
        tmpText = checkDataType("cmbSireReleased")
        If (tmpText <> "Yes/No") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be yes/no", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbSireReleased.SetFocus
            End If
            frmDataSpec.cmbSireReleased.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbSireReleased.Text, "cmbSireReleased")
    End If
End Sub

Private Sub cmbTime_Click()
    If frmDataSpec.cmbTime.ListIndex > -1 Then
        tmpText = checkDataType("cmbTime")
        If (tmpText <> "Date/Time") Then
            If frmDataSpec.Visible = True Then
                MsgBox "The data type for the field must be date/time", vbInformation, "Wrong Data Type"
                frmDataSpec.cmbTime.SetFocus
            End If
            frmDataSpec.cmbTime.ListIndex = -1
        End If
        
        Call frmDataSpec.checkFieldDups(frmDataSpec.cmbTime.Text, "cmbTime")
    End If
End Sub

Private Sub cmdCancel_Click()
    Call doCancel
End Sub

Public Function doCancel()
    On Error GoTo cantCancel
    frmDataSpec.Hide
    frmDataSpec.txtDBFile.ToolTipText = frmDataSpec.txtDBFile.Tag
    tmpString = frmDataSpec.txtDBFile.Tag
    For a = Len(tmpString) To 1 Step -1
        If Mid(tmpString, a, 1) = "\" Then
            Exit For
        End If
    Next a
    frmDataSpec.txtDBFile.Text = Mid(tmpString, a + 1, Len(tmpString))
    
    'genetics table
    frmDataSpec.cmbGeneticsTable.ListIndex = frmDataSpec.cmbGeneticsTable.Tag
    frmDataSpec.cmbUniqueId.ListIndex = frmDataSpec.cmbUniqueId.Tag
    frmDataSpec.cmbSex.ListIndex = frmDataSpec.cmbSex.Tag
    frmDataSpec.txtLoci.Text = frmDataSpec.txtLoci.Tag
    frmDataSpec.cmbFirstLocus.ListIndex = frmDataSpec.cmbFirstLocus.Tag
    frmDataSpec.cmbGenComments.ListIndex = frmDataSpec.cmbGenComments.Tag
    frmDataSpec.chkPop.Value = frmDataSpec.chkPop.Tag
    frmDataSpec.cmbPop.ListIndex = frmDataSpec.cmbPop.Tag
    frmDataSpec.chkCohort.Value = frmDataSpec.chkCohort.Tag
    frmDataSpec.cmbCohort.ListIndex = frmDataSpec.cmbCohort.Tag
    frmDataSpec.chkFlag.Value = frmDataSpec.chkFlag.Tag
    frmDataSpec.cmbFlag.ListIndex = frmDataSpec.cmbFlag.Tag
    
    'matings table
    frmDataSpec.optMatingsTable(0).Value = True
    frmDataSpec.cmbMatingsTable.ListIndex = frmDataSpec.cmbMatingsTable.Tag
    txtArray = Array("cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")
                
    For j = 0 To UBound(txtArray)
        For i = 0 To frmDataSpec.cmbBatch.ListCount - 1
            frmDataSpec(txtArray(j)).ListIndex = frmDataSpec(txtArray(j)).Tag
        Next i
    Next j
    
    Call frmDataSpec.formHide(False)
Exit Function
    
cantCancel:
    frmDataSpec.Hide
    If frmDataSpec.txtDBFile.ToolTipText <> "" Then
        accessApp.CloseCurrentDatabase
    End If
End Function

Private Sub cmdClearGenetics_Click()
    Call clearGenetics("visible")
End Sub

Private Sub cmdClearMatings_Click()
    Call clearMatings("visible")
End Sub

Private Sub cmdDBFile_Click()
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmDataSpec.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "Access Databases (*.mdb, *.accdb)" & Chr(0) & "*.mdb;*.accdb" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*"
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    If frmDataSpec.txtDBFile.ToolTipText <> "" Then
        OpenFile.lpstrInitialDir = frmDataSpec.txtDBFile.ToolTipText
    End If
    OpenFile.lpstrTitle = "Select Mating Database"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        Exit Sub
    Else
        tmpString = Trim(OpenFile.lpstrFile)
        For a = Len(tmpString) To 1 Step -1
            If Mid(tmpString, a, 1) <> Chr(0) Then
                Exit For
            End If
        Next a
        
        tmpString = Left(tmpString, a)
                
        For a = Len(tmpString) To 1 Step -1
            If Mid(tmpString, a, 1) = "\" Then
                frmDataSpec.txtDBFile.ToolTipText = tmpString
                frmDataSpec.txtDBFile.Text = Mid(tmpString, a + 1, Len(tmpString))
                Exit For
            End If
        Next a
    End If
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
    tmpBi = saveDbase
    If tmpBi = 7 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    tmpBool = checkMatingsNonopt
    If tmpBool = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Dir(App.Path & "\flag_settings.mmf") = "" Then
        Call formHide(False)
    Else
        Call formHide(True)
    End If
    
    Call frmMain.setDBase
    Call frmInput.setDBase
    Screen.MousePointer = vbDefault
    frmDataSpec.Hide
    
    If Dir(App.Path & "\flag_settings.mmf") = "" Then
        frmMain.lblDbaseConnection.Caption = "Specify 'Flag Settings' under the 'Options' menu"
        frmFlags.Show 1
    Else
        frmMain.cmdInputPIT.Enabled = True
        frmMain.lblDbaseConnection.Visible = False
    End If
End Sub

Public Function formShow()
    If frmDataSpec.txtDBFile.ToolTipText <> "" Then
        accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
    End If
End Function

Public Function formHide(tmpBool As Boolean)
    On Error Resume Next
    accessApp.CloseCurrentDatabase
    
    If tmpBool = True Then
        frmMain.cmdInputPIT.Enabled = True
    End If
End Function

Private Sub Form_Load()
    Set accessApp = CreateObject("Access.Application")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmDataSpec.Visible = True Then
        Cancel = 1
    End If
    Call doCancel
End Sub

Private Sub optMatingsTable_Click(Index As Integer)
    If Index = 0 Then
        frmDataSpec.txtMatingsTable.Visible = False
        Call frmDataSpec.enableMatFrame
    Else
        frmDataSpec.txtMatingsTable.Visible = True
        Call frmDataSpec.disableMatFrame
        If frmDataSpec.Visible = True Then
            frmDataSpec.txtMatingsTable.SetFocus
        End If
    End If
End Sub

Private Sub Picture1_Click()
    If frmDataSpec.txtDBFile.ToolTipText <> "" Then
        Call enableOptions
    End If
End Sub

Private Sub txtDBFile_Change()
    On Error Resume Next
    accessApp.CloseCurrentDatabase
    accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
    
    If frmDataSpec.txtDBFile.ToolTipText = "" Then
        frmDataSpec.cmbGeneticsTable.Clear
        frmDataSpec.cmbGeneticsTable.Enabled = False
        frmDataSpec.cmbMatingsTable.Clear
        frmDataSpec.cmbMatingsTable.Enabled = False
        
        Call frmDataSpec.disableGenFrame
        Call frmDataSpec.disableMatFrame
    Else
        Call enableOptions
    End If
End Sub

Public Function saveDbase() As Integer
    saveDbase = checkDbase("save")
    If saveDbase = 7 Then
        MsgBox "Please fill in all fields", vbInformation, "Missing Data"
        Exit Function
    End If
        
    Open App.Path & "\dbase_settings.mmf" For Output As #1
        Print #1, frmDataSpec.txtDBFile.ToolTipText
        frmDataSpec.txtDBFile.Tag = frmDataSpec.txtDBFile.ToolTipText
        
        Print #1, frmDataSpec.txtDBFile.Text
        
        Print #1, frmDataSpec.cmbGeneticsTable.Text
        frmDataSpec.cmbGeneticsTable.Tag = frmDataSpec.cmbGeneticsTable.ListIndex
        
        Print #1, frmDataSpec.cmbUniqueId.Text
        frmDataSpec.cmbUniqueId.Tag = frmDataSpec.cmbUniqueId.ListIndex
        
        Print #1, frmDataSpec.cmbSex.Text
        frmDataSpec.cmbSex.Tag = frmDataSpec.cmbSex.ListIndex
        
        Print #1, frmDataSpec.txtLoci.Text
        frmDataSpec.txtLoci.Tag = frmDataSpec.txtLoci.Text
        'Change frmFlags min loci value if necessary
        If frmFlags.chkNoAlleles.Value = 1 Then
            If CLng(frmFlags.txtMinLoci.Text) > CLng(frmDataSpec.txtLoci.Text) Then
                frmFlags.txtMinLoci.Text = frmDataSpec.txtLoci.Text
            End If
        End If
        
        Print #1, frmDataSpec.cmbFirstLocus.Text
        frmDataSpec.cmbFirstLocus.Tag = frmDataSpec.cmbFirstLocus.ListIndex
        
        Print #1, frmDataSpec.cmbGenComments.Text
        frmDataSpec.cmbGenComments.Tag = frmDataSpec.cmbGenComments.ListIndex
        
        Print #1, CStr(frmDataSpec.chkPop.Value)
        frmDataSpec.chkPop.Tag = frmDataSpec.chkPop.Value
        If frmDataSpec.chkPop.Value = 1 Then
            Print #1, frmDataSpec.cmbPop.Text
            frmDataSpec.cmbPop.Tag = frmDataSpec.cmbPop.ListIndex
        End If
        
        Print #1, CStr(frmDataSpec.chkCohort.Value)
        frmDataSpec.chkCohort.Tag = frmDataSpec.chkCohort.Value
        If frmDataSpec.chkCohort.Value = 1 Then
            Print #1, frmDataSpec.cmbCohort.Text
            frmDataSpec.cmbCohort.Tag = frmDataSpec.cmbCohort.ListIndex
            If frmFlags.chkUppQuart.Value = 1 Then
                frmFlags.lblYearsBack.Enabled = True
                frmFlags.txtYearsBack.Enabled = True
            End If
        Else
            If frmFlags.chkUppQuart.Value = 1 Then
                frmFlags.lblYearsBack.Enabled = False
                frmFlags.txtYearsBack.Enabled = False
            End If
        End If
        
        Print #1, CStr(frmDataSpec.chkFlag.Value)
        frmDataSpec.chkFlag.Tag = frmDataSpec.chkFlag.Value
        If frmDataSpec.chkFlag.Value = 1 Then
            Print #1, frmDataSpec.cmbFlag.Text
            frmDataSpec.cmbFlag.Tag = frmDataSpec.cmbFlag.ListIndex
        End If
        
        'Matings Table
        If frmDataSpec.optMatingsTable(1).Value = True Then
            saveDbase = makeMatTable
        End If
        
        txtArray = Array("cmbMatingsTable", "cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbMatDes", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")
        
        For j = 0 To UBound(txtArray)
            Print #1, frmDataSpec(txtArray(j)).Text
            frmDataSpec(txtArray(j)).Tag = frmDataSpec(txtArray(j)).ListIndex
        Next j
    Close #1
End Function

Public Function makeMatTable() As Integer
    If frmDataSpec.txtMatingsTable.Text = "" Then
        MsgBox "You must enter a name for the matings table", vbInformation, "Missing Table Name"
        frmDataSpec.txtMatingsTable.SetFocus
        makeMatTable = 7
        Exit Function
    End If
    
    Set dbTemp = accessApp.CurrentDb
    
    For i = 0 To dbTemp.TableDefs.Count - 1
        If dbTemp.TableDefs(i).Name = frmDataSpec.txtMatingsTable.Text Then
            Msg = "The table '" & frmDataSpec.txtMatingsTable.Text & "' already exists, do you want to replace it with a new, empty table?"
            Style = vbYesNo + vbQuestion + vbDefaultButton2
            Title = "Delete Existing Table"
            Response = MsgBox(Msg, Style, Title)
    
            If Response = 7 Then
                makeMatTable = 7
                Exit Function
            Else
                makeMatTable = 6
                dbTemp.TableDefs.Delete frmDataSpec.txtMatingsTable.Text
                Exit For
            End If
        End If
    Next i
    
    Set matTab = dbTemp.CreateTableDef(frmDataSpec.txtMatingsTable.Text)
    With matTab
        .Fields.Append .CreateField("Batch", dbLong)
        .Fields.Append .CreateField("Family", dbText)
        .Fields.Append .CreateField("Dam", dbText)
        .Fields.Append .CreateField("Dam_Pop", dbText)
        .Fields.Append .CreateField("Dam_Cohort", dbLong)
        .Fields.Append .CreateField("Dam_Scored_Loci", dbLong)
        .Fields.Append .CreateField("Dam_Released", dbBoolean)
        .Fields.Append .CreateField("Sire", dbText)
        .Fields.Append .CreateField("Sire_Pop", dbText)
        .Fields.Append .CreateField("Sire_Cohort", dbLong)
        .Fields.Append .CreateField("Sire_Scored_Loci", dbLong)
        .Fields.Append .CreateField("Sire_Released", dbBoolean)
        .Fields.Append .CreateField("Metric", dbText)
        .Fields.Append .CreateField("Relatedness", dbSingle)
        .Fields.Append .CreateField("Mating_Design", dbText)
        .Fields.Append .CreateField("Optimized", dbText)
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("Time", dbDate)
        .Fields.Append .CreateField("Comments", dbText)
    End With
        
    dbTemp.TableDefs.Append matTab
    
    Set tmpField = matTab.Fields("Dam_Released")
    tmpField.DefaultValue = "No"
    Set tmpProp = tmpField.CreateProperty("DisplayControl", 3, 106)
    tmpField.Properties.Append tmpProp
    Set tmpProp = tmpField.CreateProperty("Format", 10, "Yes/No")
    tmpField.Properties.Append tmpProp
    
    Set tmpField = matTab.Fields("Sire_Released")
    tmpField.DefaultValue = "No"
    Set tmpProp = tmpField.CreateProperty("DisplayControl", 3, 106)
    tmpField.Properties.Append tmpProp
    Set tmpProp = tmpField.CreateProperty("Format", 10, "Yes/No")
    tmpField.Properties.Append tmpProp
    
    Set tmpField = matTab.Fields("Date")
    'tmpField.DefaultValue = "=Date()"
    Set tmpProp = tmpField.CreateProperty("Format", dbText, "mm/dd/yyyy")
    tmpField.Properties.Append tmpProp
    
    Set tmpField = matTab.Fields("Time")
    'tmpField.DefaultValue = "=Time()"
    Set tmpProp = tmpField.CreateProperty("Format", dbText, "hh:nn:ss")
    tmpField.Properties.Append tmpProp
    
    dbTemp.Close
       
    If makeMatTable <> 6 Then
        frmDataSpec.cmbGeneticsTable.List(frmDataSpec.cmbGeneticsTable.ListCount) = frmDataSpec.txtMatingsTable.Text
        frmDataSpec.cmbMatingsTable.List(frmDataSpec.cmbMatingsTable.ListCount) = frmDataSpec.txtMatingsTable.Text
    End If
    frmDataSpec.optMatingsTable(0).Value = True
    frmDataSpec.cmbMatingsTable.Text = frmDataSpec.txtMatingsTable.Text
    frmDataSpec.txtMatingsTable.Text = ""
    Call frmDataSpec.enableMatingsOptions
    
    'set form indexes
    frmDataSpec.cmbBatch.ListIndex = 0
    frmDataSpec.cmbFamID.ListIndex = 1
    frmDataSpec.cmbDamID.ListIndex = 2
    frmDataSpec.cmbDamPop.ListIndex = 3
    frmDataSpec.cmbDamCohort.ListIndex = 4
    frmDataSpec.cmbDamLoci.ListIndex = 5
    frmDataSpec.cmbDamReleased.ListIndex = 6
    frmDataSpec.cmbSireID.ListIndex = 7
    frmDataSpec.cmbSirePop.ListIndex = 8
    frmDataSpec.cmbSireCohort.ListIndex = 9
    frmDataSpec.cmbSireLoci.ListIndex = 10
    frmDataSpec.cmbSireReleased.ListIndex = 11
    frmDataSpec.cmbMetric.ListIndex = 12
    frmDataSpec.cmbRelatedness.ListIndex = 13
    frmDataSpec.cmbMatDes.ListIndex = 14
    frmDataSpec.cmbOptimized.ListIndex = 15
    frmDataSpec.cmbDate.ListIndex = 16
    frmDataSpec.cmbTime.ListIndex = 17
    frmDataSpec.cmbComments.ListIndex = 18
End Function

Public Function checkDbase(tmpType As String) As Integer
    If frmDataSpec.txtLoci.Text = "" Then
        checkDbase = 7
        Exit Function
    End If
    
    If frmDataSpec.optMatingsTable(0).Value = True And tmpType = "save" Then
        txtArray = Array("cmbGeneticsTable", "cmbUniqueId", "cmbSex", "cmbFirstLocus", "cmbGenComments", "cmbMatingsTable", "cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")
    Else
        txtArray = Array("cmbGeneticsTable", "cmbUniqueId", "cmbSex", "cmbFirstLocus", "cmbGenComments")
    End If
    
    chkArray = Array("chkPop", "chkCohort", "chkFlag")
    For x = 0 To UBound(chkArray)
        If frmDataSpec(chkArray(x)).Value = 1 Then
            a = UBound(txtArray) + 1
            ReDim Preserve txtArray(a)
            txtArray(a) = "cmb" & Mid(chkArray(x), 4, Len(chkArray(x)))
        End If
    Next x
    
    For x = 0 To UBound(txtArray)
        If frmDataSpec(txtArray(x)).ListIndex = -1 Then
            checkDbase = 7
            Exit Function
        End If
    Next x
End Function

Public Function checkDataTypeGen(control As String) As String
    Set dbTemp = accessApp.CurrentDb
    Set matTab = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text)
    
    Select Case CLng(matTab.Fields(frmDataSpec(control).Text).Type)
        Case dbBoolean: checkDataTypeGen = "Yes/No"            ' 1
        Case dbByte: checkDataTypeGen = "Byte"                 ' 2
        Case dbInteger: checkDataTypeGen = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbAutoIncrField) = 0& Then
                checkDataTypeGen = "Long Integer"
            Else
                checkDataTypeGen = "AutoNumber"
            End If
        Case dbCurrency: checkDataTypeGen = "Currency"         ' 5
        Case dbSingle: checkDataTypeGen = "Single"             ' 6
        Case dbDouble: checkDataTypeGen = "Double"             ' 7
        Case dbDate: checkDataTypeGen = "Date/Time"            ' 8
        Case dbBinary: checkDataTypeGen = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbFixedField) = 0& Then
                checkDataTypeGen = "Text"
            Else
                checkDataTypeGen = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: checkDataTypeGen = "OLE Object"     '11
        Case dbMemo                                     '12
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbHyperlinkField) = 0& Then
                checkDataTypeGen = "Memo"
            Else
                checkDataTypeGen = "Hyperlink"
            End If
        Case dbGUID: checkDataTypeGen = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: checkDataTypeGen = "Big Integer"        '16
        Case dbVarBinary: checkDataTypeGen = "VarBinary"       '17
        Case dbChar: checkDataTypeGen = "Char"                 '18
        Case dbNumeric: checkDataTypeGen = "Numeric"           '19
        Case dbDecimal: checkDataTypeGen = "Decimal"           '20
        Case dbFloat: checkDataTypeGen = "Float"               '21
        Case dbTime: checkDataTypeGen = "Time"                 '22
        Case dbTimeStamp: checkDataTypeGen = "Time Stamp"      '23
    End Select
    
    dbTemp.Close
End Function

Public Function checkDataType(control As String) As String
    Set dbTemp = accessApp.CurrentDb
    Set matTab = dbTemp.TableDefs(frmDataSpec.cmbMatingsTable.Text)
    
    Select Case CLng(matTab.Fields(frmDataSpec(control).Text).Type)
        Case dbBoolean: checkDataType = "Yes/No"            ' 1
        Case dbByte: checkDataType = "Byte"                 ' 2
        Case dbInteger: checkDataType = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbAutoIncrField) = 0& Then
                checkDataType = "Long Integer"
            Else
                checkDataType = "AutoNumber"
            End If
        Case dbCurrency: checkDataType = "Currency"         ' 5
        Case dbSingle: checkDataType = "Single"             ' 6
        Case dbDouble: checkDataType = "Double"             ' 7
        Case dbDate: checkDataType = "Date/Time"            ' 8
        Case dbBinary: checkDataType = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbFixedField) = 0& Then
                checkDataType = "Text"
            Else
                checkDataType = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: checkDataType = "OLE Object"     '11
        Case dbMemo                                     '12
            If (matTab.Fields(frmDataSpec(control).Text).Attributes And dbHyperlinkField) = 0& Then
                checkDataType = "Memo"
            Else
                checkDataType = "Hyperlink"
            End If
        Case dbGUID: checkDataType = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: checkDataType = "Big Integer"        '16
        Case dbVarBinary: checkDataType = "VarBinary"       '17
        Case dbChar: checkDataType = "Char"                 '18
        Case dbNumeric: checkDataType = "Numeric"           '19
        Case dbDecimal: checkDataType = "Decimal"           '20
        Case dbFloat: checkDataType = "Float"               '21
        Case dbTime: checkDataType = "Time"                 '22
        Case dbTimeStamp: checkDataType = "Time Stamp"      '23
    End Select
    
    dbTemp.Close
End Function

Public Function checkFieldDups(tmpText As String, tmpControl As String)
    txtArray = Array("cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbMatDes", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")
    
    For a = 0 To UBound(txtArray)
        If txtArray(a) <> tmpControl Then
            If frmDataSpec(txtArray(a)).ListIndex = frmDataSpec(tmpControl).ListIndex And frmDataSpec(txtArray(a)).ListIndex > -1 Then
                MsgBox "Cannot select a column that is already selected for another field", vbInformation, "Column Already Being Used"
                frmDataSpec(tmpControl).ListIndex = -1
                frmDataSpec(tmpControl).SetFocus
                Exit Function
            End If
        End If
    Next a
End Function

Public Function enableOptions()
    Set dbTemp = accessApp.CurrentDb
    
    frmDataSpec.cmbGeneticsTable.Enabled = True
    frmDataSpec.cmbGeneticsTable.Clear
    frmDataSpec.cmbMatingsTable.Enabled = True
    frmDataSpec.cmbMatingsTable.Clear
    j = 0
    For i = 0 To dbTemp.TableDefs.Count - 1
        If Left(dbTemp.TableDefs(i).Name, 3) <> "MSy" Then
            frmDataSpec.cmbGeneticsTable.List(j) = dbTemp.TableDefs(i).Name
            frmDataSpec.cmbMatingsTable.List(j) = dbTemp.TableDefs(i).Name
            j = j + 1
        End If
    Next i
    dbTemp.Close
    
    Call enableGeneticOptions
    Call enableMatingsOptions
End Function

Public Function enableGeneticOptions()
    Call clearGenetics("all")
    
    If frmDataSpec.cmbGeneticsTable.Text = frmDataSpec.cmbMatingsTable.Text And frmDataSpec.cmbGeneticsTable <> "" Then
        MsgBox "The table selected for the genetics input needs to be different than the table selected for matings output", vbInformation, "Table Already Selected"
        frmDataSpec.cmbGeneticsTable.ListIndex = -1
        frmDataSpec.cmbGeneticsTable.SetFocus
        Exit Function
    End If
        
    If frmDataSpec.cmbGeneticsTable.ListIndex <> -1 Then
        Set dbTemp = accessApp.CurrentDb
        Call frmDataSpec.enableGenFrame
        
        For j = 0 To dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields.Count - 1
            frmDataSpec.cmbUniqueId.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbSex.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbFirstLocus.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbGenComments.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbPop.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbCohort.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
            frmDataSpec.cmbFlag.List(j) = dbTemp.TableDefs(frmDataSpec.cmbGeneticsTable.Text).Fields(j).Name
        Next j
        dbTemp.Close
    Else
        Call frmDataSpec.disableGenFrame
    End If
End Function

Public Function enableMatingsOptions()
    Call clearMatings("all")
    
    If frmDataSpec.cmbMatingsTable.Text = frmDataSpec.cmbGeneticsTable.Text And frmDataSpec.cmbMatingsTable <> "" Then
        MsgBox "The table selected for the matings output needs to be different than the table selected for genetics input", vbInformation, "Table Already Selected"
        frmDataSpec.cmbMatingsTable.ListIndex = -1
        frmDataSpec.cmbMatingsTable.SetFocus
        Exit Function
    End If
            
    If frmDataSpec.cmbMatingsTable.ListIndex <> -1 Then
        Set dbTemp = accessApp.CurrentDb
        Call frmDataSpec.enableMatFrame
        
        For j = 0 To dbTemp.TableDefs(frmDataSpec.cmbMatingsTable.Text).Fields.Count - 1
            For x = 0 To UBound(txtArray)
                frmDataSpec(txtArray(x)).List(j) = dbTemp.TableDefs(frmDataSpec.cmbMatingsTable.Text).Fields(j).Name
            Next x
        Next j
        dbTemp.Close
    Else
        Call frmDataSpec.disableMatFrame
    End If
End Function

Public Function clearGenetics(tmpType As String)
    If tmpType = "all" Then
        frmDataSpec.cmbUniqueId.Clear
        frmDataSpec.cmbSex.Clear
        frmDataSpec.txtLoci.Text = ""
        frmDataSpec.cmbFirstLocus.Clear
        frmDataSpec.cmbGenComments.Clear
        frmDataSpec.chkPop.Value = 0
        frmDataSpec.cmbPop.Clear
        frmDataSpec.chkCohort.Value = 0
        frmDataSpec.cmbCohort.Clear
        frmDataSpec.chkFlag.Value = 0
        frmDataSpec.cmbFlag.Clear
    Else
        frmDataSpec.cmbUniqueId.ListIndex = -1
        frmDataSpec.cmbSex.ListIndex = -1
        frmDataSpec.txtLoci.Text = ""
        frmDataSpec.cmbFirstLocus.ListIndex = -1
        frmDataSpec.cmbGenComments.ListIndex = -1
        frmDataSpec.cmbPop.ListIndex = -1
        frmDataSpec.cmbCohort.ListIndex = -1
        frmDataSpec.cmbFlag.ListIndex = -1
    End If
End Function

Public Function clearMatings(tmpType As String)
    txtArray = Array("cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbMatDes", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")

    If tmpType = "all" Then
        For j = 0 To UBound(txtArray)
            frmDataSpec(txtArray(j)).Clear
        Next j
    Else
        For j = 0 To UBound(txtArray)
            frmDataSpec(txtArray(j)).ListIndex = -1
        Next j
    End If
End Function

Public Function importDbaseSettings() As Boolean
    On Error GoTo finishUp
    
    Open App.Path & "\dbase_settings.mmf" For Input As #1
        'check to see if path to database exists
        Input #1, tmpText
        If Dir(tmpText) <> "" Then
            frmDataSpec.txtDBFile.ToolTipText = tmpText
            accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
            frmDataSpec.txtDBFile.Tag = tmpText
            Input #1, tmpText
            frmDataSpec.txtDBFile.Text = tmpText
            
            'Genetics Table
            Input #1, tmpText
            tmpBi = 0
            For i = 0 To frmDataSpec.cmbGeneticsTable.ListCount - 1
                If frmDataSpec.cmbGeneticsTable.List(i) = tmpText Then
                    frmDataSpec.cmbGeneticsTable.ListIndex = i
                    frmDataSpec.cmbGeneticsTable.Tag = i
                    Call enableGeneticOptions
                    tmpBi = 1
                    Exit For
                End If
            Next i
            
            If tmpBi = 1 Then
                'unique id field
                Input #1, tmpText
                For i = 0 To frmDataSpec.cmbUniqueId.ListCount - 1
                    If frmDataSpec.cmbUniqueId.List(i) = tmpText Then
                        frmDataSpec.cmbUniqueId.ListIndex = i
                        frmDataSpec.cmbUniqueId.Tag = i
                        Exit For
                    End If
                Next i
                
                 'sex field
                Input #1, tmpText
                For i = 0 To frmDataSpec.cmbSex.ListCount - 1
                    If frmDataSpec.cmbSex.List(i) = tmpText Then
                        frmDataSpec.cmbSex.ListIndex = i
                        frmDataSpec.cmbSex.Tag = i
                        Exit For
                    End If
                Next i
               
                'Number of loci
                Input #1, tmpText
                frmDataSpec.txtLoci.Text = tmpText
                frmDataSpec.txtLoci.Tag = tmpText
                
                'First locus field
                Input #1, tmpText
                For i = 0 To frmDataSpec.cmbFirstLocus.ListCount - 1
                    If frmDataSpec.cmbFirstLocus.List(i) = tmpText Then
                        frmDataSpec.cmbFirstLocus.ListIndex = i
                        frmDataSpec.cmbFirstLocus.Tag = i
                        Exit For
                    End If
                Next i
                
                'Comments field
                Input #1, tmpText
                For i = 0 To frmDataSpec.cmbGenComments.ListCount - 1
                    If frmDataSpec.cmbGenComments.List(i) = tmpText Then
                        frmDataSpec.cmbGenComments.ListIndex = i
                        frmDataSpec.cmbGenComments.Tag = i
                        Exit For
                    End If
                Next i
                
                'Check Pop
                Input #1, tmpText
                If tmpText = "1" Then
                    frmDataSpec.chkPop.Value = CLng(tmpText)
                    frmDataSpec.chkPop.Tag = tmpText
                    Input #1, tmpText
                    For i = 0 To frmDataSpec.cmbPop.ListCount - 1
                        If frmDataSpec.cmbPop.List(i) = tmpText Then
                            frmDataSpec.cmbPop.ListIndex = i
                            frmDataSpec.cmbPop.Tag = i
                            Exit For
                        End If
                    Next i
                End If
                    
                'Check Cohort
                Input #1, tmpText
                If tmpText = "1" Then
                    frmDataSpec.chkCohort.Value = CLng(tmpText)
                    frmDataSpec.chkCohort.Tag = tmpText
                    Input #1, tmpText
                    For i = 0 To frmDataSpec.cmbCohort.ListCount - 1
                        If frmDataSpec.cmbCohort.List(i) = tmpText Then
                            frmDataSpec.cmbCohort.ListIndex = i
                            frmDataSpec.cmbCohort.Tag = i
                            Exit For
                        End If
                    Next i
                End If
                
                'Check Flag
                Input #1, tmpText
                If tmpText = "1" Then
                    frmDataSpec.chkFlag.Value = CLng(tmpText)
                    frmDataSpec.chkFlag.Tag = tmpText
                    Input #1, tmpText
                    For i = 0 To frmDataSpec.cmbFlag.ListCount - 1
                        If frmDataSpec.cmbFlag.List(i) = tmpText Then
                            frmDataSpec.cmbFlag.ListIndex = i
                            frmDataSpec.cmbFlag.Tag = i
                            Exit For
                        End If
                    Next i
                End If
            End If
            
            'Verify all fields are filled and if not raise an error
            tmpBi = frmDataSpec.checkDbase("check")
            If tmpBi = 7 Then
                Err.Raise (0)
            End If
                
            'Matings Table
            Input #1, tmpText
            tmpBi = 0
            For i = 0 To frmDataSpec.cmbMatingsTable.ListCount - 1
                If frmDataSpec.cmbMatingsTable.List(i) = tmpText Then
                    frmDataSpec.cmbMatingsTable.ListIndex = i
                    frmDataSpec.cmbMatingsTable.Tag = i
                    tmpBi = 1
                    Exit For
                End If
            Next i
            
            Call enableMatingsOptions
            If tmpBi = 0 Then
                frmDataSpec.optMatingsTable(1).Value = True
                Err.Raise (1)
            Else
                frmDataSpec.optMatingsTable(0).Value = True
                
                txtArray = Array("cmbBatch", "cmbFamID", "cmbDamID", "cmbDamPop", "cmbDamCohort", "cmbDamLoci", "cmbDamReleased", "cmbSireID", "cmbSirePop", "cmbSireCohort", "cmbSireLoci", "cmbSireReleased", "cmbMetric", "cmbRelatedness", "cmbMatDes", "cmbOptimized", "cmbDate", "cmbTime", "cmbComments")
                
                For j = 0 To UBound(txtArray)
                    Input #1, tmpText
                    For i = 0 To frmDataSpec.cmbBatch.ListCount - 1
                        If frmDataSpec(txtArray(j)).List(i) = tmpText Then
                            frmDataSpec(txtArray(j)).ListIndex = i
                            frmDataSpec(txtArray(j)).Tag = i
                            Exit For
                        End If
                    Next i
                Next j
            End If
            
            'Check to see if the non-optimized matings table is present and has correct fields and data types
            Call checkMatingsNonopt
            
            accessApp.CloseCurrentDatabase
            importDbaseSettings = True
            
            Call frmMain.setDBase
            Call frmInput.setDBase
        End If
    Close #1
Exit Function

finishUp:
    Close #1
    importDbaseSettings = False
    Screen.MousePointer = vbDefault
    'Set dbTemp = accessApp.CurrentDb
    'frmDataSpec.Show 1
End Function

Public Function togglePop()
    If frmDataSpec.chkPop.Value = 1 Then
        frmDataSpec.cmbPop.Enabled = True
    Else
        frmDataSpec.cmbPop.Enabled = False
    End If
End Function

Public Function toggleCohort()
    If frmDataSpec.chkCohort.Value = 1 Then
        frmDataSpec.cmbCohort.Enabled = True
    Else
        frmDataSpec.cmbCohort.Enabled = False
    End If
End Function

Public Function toggleFlag()
    If frmDataSpec.chkFlag.Value = 1 Then
        frmDataSpec.cmbFlag.Enabled = True
    Else
        frmDataSpec.cmbFlag.Enabled = False
    End If
End Function

Public Function enableGenFrame()
    frmDataSpec.frameGenetics.Enabled = True
    frmDataSpec.lblUniqueId.Enabled = True
    frmDataSpec.cmbUniqueId.Enabled = True
    frmDataSpec.lblSex.Enabled = True
    frmDataSpec.cmbSex.Enabled = True
    frmDataSpec.lblLoci.Enabled = True
    frmDataSpec.txtLoci.Enabled = True
    frmDataSpec.lblFirstLocus.Enabled = True
    frmDataSpec.cmbFirstLocus.Enabled = True
    frmDataSpec.chkPop.Enabled = True
    'frmDataSpec.cmbPop.Enabled = True
    frmDataSpec.chkCohort.Enabled = True
    'frmDataSpec.cmbCohort.Enabled = True
    frmDataSpec.chkFlag.Enabled = True
    'frmDataSpec.cmbFlag.Enabled = True
End Function

Public Function disableGenFrame()
    frmDataSpec.frameGenetics.Enabled = False
    frmDataSpec.lblUniqueId.Enabled = False
    frmDataSpec.cmbUniqueId.Enabled = False
    frmDataSpec.lblSex.Enabled = False
    frmDataSpec.cmbSex.Enabled = False
    frmDataSpec.lblLoci.Enabled = False
    frmDataSpec.txtLoci.Enabled = False
    frmDataSpec.lblFirstLocus.Enabled = False
    frmDataSpec.cmbFirstLocus.Enabled = False
    frmDataSpec.chkPop.Enabled = False
    frmDataSpec.cmbPop.Enabled = False
    frmDataSpec.chkCohort.Enabled = False
    frmDataSpec.cmbCohort.Enabled = False
    frmDataSpec.chkFlag.Enabled = False
    frmDataSpec.cmbFlag.Enabled = False
End Function

Public Function enableMatFrame()
    frmDataSpec.frameMatings.Enabled = True
    frmDataSpec.lblBatch.Enabled = True
    frmDataSpec.cmbBatch.Enabled = True
    frmDataSpec.lblFamID.Enabled = True
    frmDataSpec.cmbFamID.Enabled = True
    frmDataSpec.lblDamInfo.Enabled = True
    frmDataSpec.lblDamID.Enabled = True
    frmDataSpec.cmbDamID.Enabled = True
    frmDataSpec.lblDamPop.Enabled = True
    frmDataSpec.cmbDamPop.Enabled = True
    frmDataSpec.lblDamCohort.Enabled = True
    frmDataSpec.cmbDamCohort.Enabled = True
    frmDataSpec.lblDamLoci.Enabled = True
    frmDataSpec.cmbDamLoci.Enabled = True
    frmDataSpec.lblDamReleased.Enabled = True
    frmDataSpec.cmbDamReleased.Enabled = True
    frmDataSpec.lblSireInfo.Enabled = True
    frmDataSpec.lblSireID.Enabled = True
    frmDataSpec.cmbSireID.Enabled = True
    frmDataSpec.lblSirePop.Enabled = True
    frmDataSpec.cmbSirePop.Enabled = True
    frmDataSpec.lblSireCohort.Enabled = True
    frmDataSpec.cmbSireCohort.Enabled = True
    frmDataSpec.lblSireLoci.Enabled = True
    frmDataSpec.cmbSireLoci.Enabled = True
    frmDataSpec.lblSireReleased.Enabled = True
    frmDataSpec.cmbSireReleased.Enabled = True
    frmDataSpec.lblMetric.Enabled = True
    frmDataSpec.cmbMetric.Enabled = True
    frmDataSpec.lblRelatedness.Enabled = True
    frmDataSpec.cmbRelatedness.Enabled = True
    frmDataSpec.lblMatDes.Enabled = True
    frmDataSpec.cmbMatDes.Enabled = True
    frmDataSpec.lblOptimized.Enabled = True
    frmDataSpec.cmbOptimized.Enabled = True
    frmDataSpec.lblDate.Enabled = True
    frmDataSpec.cmbDate.Enabled = True
    frmDataSpec.lblTime.Enabled = True
    frmDataSpec.cmbTime.Enabled = True
    frmDataSpec.lblComments.Enabled = True
    frmDataSpec.cmbComments.Enabled = True
End Function

Public Function disableMatFrame()
    frmDataSpec.frameMatings.Enabled = False
    frmDataSpec.frameMatings.Enabled = False
    frmDataSpec.lblBatch.Enabled = False
    frmDataSpec.cmbBatch.Enabled = False
    frmDataSpec.lblFamID.Enabled = False
    frmDataSpec.cmbFamID.Enabled = False
    frmDataSpec.lblDamInfo.Enabled = False
    frmDataSpec.lblDamID.Enabled = False
    frmDataSpec.cmbDamID.Enabled = False
    frmDataSpec.lblDamPop.Enabled = False
    frmDataSpec.cmbDamPop.Enabled = False
    frmDataSpec.lblDamCohort.Enabled = False
    frmDataSpec.cmbDamCohort.Enabled = False
    frmDataSpec.lblDamLoci.Enabled = False
    frmDataSpec.cmbDamLoci.Enabled = False
    frmDataSpec.lblDamReleased.Enabled = False
    frmDataSpec.cmbDamReleased.Enabled = False
    frmDataSpec.lblSireInfo.Enabled = False
    frmDataSpec.lblSireID.Enabled = False
    frmDataSpec.cmbSireID.Enabled = False
    frmDataSpec.lblSirePop.Enabled = False
    frmDataSpec.cmbSirePop.Enabled = False
    frmDataSpec.lblSireCohort.Enabled = False
    frmDataSpec.cmbSireCohort.Enabled = False
    frmDataSpec.lblSireLoci.Enabled = False
    frmDataSpec.cmbSireLoci.Enabled = False
    frmDataSpec.lblSireReleased.Enabled = False
    frmDataSpec.cmbSireReleased.Enabled = False
    frmDataSpec.lblMetric.Enabled = False
    frmDataSpec.cmbMetric.Enabled = False
    frmDataSpec.lblMatDes.Enabled = False
    frmDataSpec.cmbMatDes.Enabled = False
    frmDataSpec.lblRelatedness.Enabled = False
    frmDataSpec.cmbRelatedness.Enabled = False
    frmDataSpec.lblOptimized.Enabled = False
    frmDataSpec.cmbOptimized.Enabled = False
    frmDataSpec.lblDate.Enabled = False
    frmDataSpec.cmbDate.Enabled = False
    frmDataSpec.lblTime.Enabled = False
    frmDataSpec.cmbTime.Enabled = False
    frmDataSpec.lblComments.Enabled = False
    frmDataSpec.cmbComments.Enabled = False
End Function

Public Function checkMatingsNonopt() As Boolean
    If frmDataSpec.Visible = True Then
        On Error GoTo badData
    End If
    
    Set dbTemp = accessApp.CurrentDb
    
    tmpBool = checkTable
    
    If tmpBool = False Then
        Call makeMatNonoptTable
    Else
        Set matTab = dbTemp.TableDefs(frmDataSpec.txtMatingsNonopt.Text)
        
        For i = 0 To matTab.Fields.Count - 1
            tmpString = checkField(matTab.Fields(i))
            Select Case i
                Case 0, 4, 5, 8, 9 'Long Integer
                    If (tmpString <> "Long Integer" And tmpString <> "Single" And tmpString <> "Double" And tmpString <> "Decimal") Then
                        MsgBox "For the 'matings_nonopt' table, the data type for '" & matTab.Fields(i).Name & "' must be long integer. To resolve this issue, your options are" & Chr(13) & Chr(13) & "1) Rename the table to something besides 'mating_nonopt'" & Chr(13) & "2) Change the data type of this field in Access" & Chr(13) & "3) Delete the 'matings_nonopt' table from the Access database", vbInformation, "Wrong Data Type"
                        Err.Raise (1)
                    End If
                Case 1, 2, 3, 6, 7, 10, 12, 13 'Text
                    If (tmpString <> "Text" And tmpString <> "Text (fixed width)") Then
                        MsgBox "For the 'matings_nonopt' table, the data type for '" & matTab.Fields(i).Name & "' must be text. To resolve this issue, your options are" & Chr(13) & Chr(13) & "1) Rename the table to something besides 'mating_nonopt'" & Chr(13) & "2) Change the data type of this field in Access" & Chr(13) & "3) Delete the 'matings_nonopt' table from the Access database", vbInformation, "Wrong Data Type"
                        Err.Raise (1)
                    End If
                Case 14, 15 'Date Time
                    If (tmpString <> "Date/Time") Then
                        MsgBox "For the 'matings_nonopt' table, the data type for '" & matTab.Fields(i).Name & "' must be date/time. To resolve this issue, your options are" & Chr(13) & Chr(13) & "1) Rename the table to something besides 'mating_nonopt'" & Chr(13) & "2) Change the data type of this field in Access" & Chr(13) & "3) Delete the 'matings_nonopt' table from the Access database", vbInformation, "Wrong Data Type"
                        Err.Raise (1)
                    End If
                Case 11 'Single
                    If (tmpString <> "Single" And tmpString <> "Double" And tmpString <> "Decimal") Then
                        MsgBox "For the 'matings_nonopt' table, the data type for '" & matTab.Fields(i).Name & "' must be long single. To resolve this issue, your options are" & Chr(13) & Chr(13) & "1) Rename the table to something besides 'mating_nonopt'" & Chr(13) & "2) Change the data type of this field in Access" & Chr(13) & "3) Delete the 'matings_nonopt' table from the Access database", vbInformation, "Wrong Data Type"
                        Err.Raise (1)
                    End If
            End Select
        Next i
    End If

    dbTemp.Close
    checkMatingsNonopt = True
Exit Function
    
badData:
    checkMatingsNonopt = False
End Function

Public Function checkTable() As Boolean
    On Error GoTo notThere
    Set matTab = dbTemp.TableDefs(frmDataSpec.txtMatingsNonopt.Text)
    checkTable = True
    Exit Function
    
notThere:
    checkTable = False
End Function

Public Function checkField(tmpField As Variant) As String
    Select Case CLng(tmpField.Type)
        Case dbBoolean: checkField = "Yes/No"            ' 1
        Case dbByte: checkField = "Byte"                 ' 2
        Case dbInteger: checkField = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (tmpField.Attributes And dbAutoIncrField) = 0& Then
                checkField = "Long Integer"
            Else
                checkField = "AutoNumber"
            End If
        Case dbCurrency: checkField = "Currency"         ' 5
        Case dbSingle: checkField = "Single"             ' 6
        Case dbDouble: checkField = "Double"             ' 7
        Case dbDate: checkField = "Date/Time"            ' 8
        Case dbBinary: checkField = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (tmpField.Attributes And dbFixedField) = 0& Then
                checkField = "Text"
            Else
                checkField = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: checkField = "OLE Object"     '11
        Case dbMemo                                     '12
            If (tmpField.Attributes And dbHyperlinkField) = 0& Then
                checkField = "Memo"
            Else
                checkField = "Hyperlink"
            End If
        Case dbGUID: checkField = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: checkField = "Big Integer"        '16
        Case dbVarBinary: checkField = "VarBinary"       '17
        Case dbChar: checkField = "Char"                 '18
        Case dbNumeric: checkField = "Numeric"           '19
        Case dbDecimal: checkField = "Decimal"           '20
        Case dbFloat: checkField = "Float"               '21
        Case dbTime: checkField = "Time"                 '22
        Case dbTimeStamp: checkField = "Time Stamp"      '23
    End Select
End Function

Public Function makeMatNonoptTable() As Integer
    'Set dbTemp = accessApp.CurrentDb
    Set matTab = dbTemp.CreateTableDef(frmDataSpec.txtMatingsNonopt.Text)
    With matTab
        .Fields.Append .CreateField("Batch", dbLong)            '0
        .Fields.Append .CreateField("Family", dbText)           '1
        .Fields.Append .CreateField("Dam", dbText)              '2
        .Fields.Append .CreateField("Dam_Pop", dbText)          '3
        .Fields.Append .CreateField("Dam_Cohort", dbLong)       '4
        .Fields.Append .CreateField("Dam_Scored_Loci", dbLong)  '5
        .Fields.Append .CreateField("Sire", dbText)             '6
        .Fields.Append .CreateField("Sire_Pop", dbText)         '7
        .Fields.Append .CreateField("Sire_Cohort", dbLong)      '8
        .Fields.Append .CreateField("Sire_Scored_Loci", dbLong) '9
        .Fields.Append .CreateField("Metric", dbText)           '10
        .Fields.Append .CreateField("Relatedness", dbSingle)    '11
        .Fields.Append .CreateField("Mating_Design", dbText)    '12
        .Fields.Append .CreateField("Optimized", dbText)        '13
        .Fields.Append .CreateField("Date", dbDate)             '14
        .Fields.Append .CreateField("Time", dbDate)             '15
    End With
        
    dbTemp.TableDefs.Append matTab
        
    Set tmpField = matTab.Fields("Date")
    'tmpField.DefaultValue = "=Date()"
    Set tmpProp = tmpField.CreateProperty("Format", dbText, "mm/dd/yyyy")
    tmpField.Properties.Append tmpProp
    
    Set tmpField = matTab.Fields("Time")
    'tmpField.DefaultValue = "=Time()"
    Set tmpProp = tmpField.CreateProperty("Format", dbText, "hh:nn:ss")
    tmpField.Properties.Append tmpProp
End Function

Private Sub txtLoci_Change()
    If frmDataSpec.txtLoci.Text <> "" Then
        If IsNumeric(frmDataSpec.txtLoci.Text) = True Then
            If CLng(frmDataSpec.txtLoci.Text) < 0 Then
                MsgBox "Please enter an integer greater than or equal to 0", vbInformation, "Input Error"
                frmDataSpec.txtLoci.Text = ""
                frmDataSpec.txtLoci.SetFocus
            Else
                frmDataSpec.txtLoci.Text = CLng(frmDataSpec.txtLoci.Text)
            End If
        Else
            MsgBox "Please enter an integer greater than or equal to 0", vbInformation, "Input Error"
            frmDataSpec.txtLoci.Text = ""
            frmDataSpec.txtLoci.SetFocus
        End If
    End If
End Sub

Private Sub txtLoci_LostFocus()
    'adjust flag min loci value if necessary
    If frmFlags.chkNoAlleles.Value = 1 And frmFlags.txtMinLoci.Text <> "" Then
        If CLng(frmDataSpec.txtLoci.Text) < CLng(frmFlags.txtMinLoci.Text) Then
            MsgBox "The number of loci specified here is lower than the number specified under 'Flag Settings'. If this number is correct than the number under 'Flag Settings' will automatically be reduced to this value upon saving these settings. You may also need to recalculate the relatedness distribution if that warning is also checked.", vbInformation, "Loci Incongruence"
        End If
    End If
End Sub

