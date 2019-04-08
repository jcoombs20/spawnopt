VERSION 5.00
Begin VB.Form frmAddTag 
   Caption         =   "Add Individual"
   ClientHeight    =   3465
   ClientLeft      =   2925
   ClientTop       =   2565
   ClientWidth     =   8415
   Icon            =   "Add Tag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Comment for individual to be added to the genetcs table"
      Top             =   1440
      Width           =   7935
   End
   Begin VB.CommandButton cmdAddTag 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Individual"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Adds the above individual to the genetics table"
      Top             =   2760
      Width           =   1695
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
      TabIndex        =   6
      ToolTipText     =   "Closes form without adding the current individual to the genetcs table"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cmbGender 
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
      ItemData        =   "Add Tag.frx":048A
      Left            =   3000
      List            =   "Add Tag.frx":0494
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Gender of individual to be added to the genetcs table"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtCohort 
      Alignment       =   2  'Center
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
      Left            =   7200
      TabIndex        =   3
      ToolTipText     =   "Cohort of individual to be added to the genetcs table"
      Top             =   480
      Width           =   975
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
      ItemData        =   "Add Tag.frx":04A6
      Left            =   4800
      List            =   "Add Tag.frx":04BF
      TabIndex        =   2
      ToolTipText     =   "Population of individual to be added to the genetcs table"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtTagAdd 
      Alignment       =   2  'Center
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
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Unique identifier of individual that is being added to the genetics table"
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCohort 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cohort"
      Enabled         =   0   'False
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
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblPop 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Population"
      Enabled         =   0   'False
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
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long, wrkJet, dbsNew As Database, rstTemp As Recordset
Dim Msg, Style, Title, Response, empID(50) As Long, empName(50) As String, q As Long

Private Sub cmdAddTag_Click()
    If frmAddTag.cmbGender.Text = "" Then
        Msg = "Please fill in the 'Sex' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Gender"
        Response = MsgBox(Msg, Style, Title)
        
        frmAddTag.cmbGender.SetFocus
        Exit Sub
    End If
    
    If frmAddTag.cmbPop.Enabled = True And frmAddTag.cmbPop.Text = "" Then
        Msg = "Please fill in the 'Population' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Population"
        Response = MsgBox(Msg, Style, Title)
        
        frmAddTag.cmbPop.SetFocus
        Exit Sub
    End If

    If frmAddTag.txtCohort.Enabled = True And frmAddTag.txtCohort.Text = "" Then
        Msg = "Please fill in the 'Cohort' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Cohort"
        Response = MsgBox(Msg, Style, Title)
        
        frmAddTag.txtCohort.SetFocus
        Exit Sub
    End If
    
    Call frmInput.addTag
        
    frmInput.txtSex.Text = frmAddTag.cmbGender.Text
    frmInput.txtDrainage.Text = frmAddTag.cmbPop.Text
    frmInput.txtYear.Text = frmAddTag.txtCohort.Text
    
    Unload frmAddTag
End Sub

Private Sub cmdCancel_Click()
    Unload frmAddTag
End Sub

Private Sub Form_Load()
    frmAddTag.txtTagAdd.Locked = False
    frmAddTag.txtTagAdd.Text = frmInput.txtTagCur2.Text
    frmAddTag.txtTagAdd.Locked = True
    
    If frmDataSpec.chkPop.Value = 1 Then
        frmAddTag.lblPop.Enabled = True
        frmAddTag.cmbPop.Enabled = True
        Call frmInput.getPops("full")
    End If
    
    If frmDataSpec.chkCohort.Value = 1 Then
        frmAddTag.lblCohort.Enabled = True
        frmAddTag.txtCohort.Enabled = True
    End If
End Sub
