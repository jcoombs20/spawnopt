VERSION 5.00
Begin VB.Form frmPartialAddTag 
   Caption         =   "Add Individual"
   ClientHeight    =   1890
   ClientLeft      =   3510
   ClientTop       =   2730
   ClientWidth     =   8415
   Icon            =   "frmPartialAddTag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
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
      ToolTipText     =   "Cohort of individual to be added"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Closes form without adding the current individual's information to the 'Input Form'"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddTag 
      BackColor       =   &H0080FF80&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Adds the above information to the 'Input Form'"
      Top             =   1200
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
      ItemData        =   "frmPartialAddTag.frx":048A
      Left            =   3000
      List            =   "frmPartialAddTag.frx":0494
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Gender of individual to be added"
      Top             =   480
      Width           =   1455
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
      TabStop         =   0   'False
      ToolTipText     =   "Unique ID of individual that is currently being added"
      Top             =   480
      Width           =   2415
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
      ItemData        =   "frmPartialAddTag.frx":04A6
      Left            =   4800
      List            =   "frmPartialAddTag.frx":04BF
      TabIndex        =   2
      ToolTipText     =   "Population of the individual to be added"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblCohort 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cohort"
      Enabled         =   0   'False
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
      Left            =   6720
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblPop 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Population"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Population of individual to be added"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPartialAddTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg, Style, Title, Response

Private Sub cmdAddTag_Click()
    If frmPartialAddTag.cmbGender.Text = "" Then
        Msg = "Please fill in the 'Sex' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Gender"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.cmbGender.SetFocus
        Exit Sub
    End If
    
    If frmPartialAddTag.cmbPop.Enabled = True And frmPartialAddTag.cmbPop.Text = "" Then
        Msg = "Please fill in the 'Population' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Population"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.cmbPop.SetFocus
        Exit Sub
    End If

    If frmPartialAddTag.txtCohort.Enabled = True And frmPartialAddTag.txtCohort.Text = "" Then
        Msg = "Please fill in the 'Cohort' field."
        Style = vbOKOnly + vbInformation + vbDefaultButton1
        Title = "Missing Cohort"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.txtCohort.SetFocus
        Exit Sub
    End If
            
    frmInput.txtSex.Text = frmPartialAddTag.cmbGender.Text
    frmInput.txtDrainage.Text = frmPartialAddTag.cmbPop.Text
    frmInput.txtYear.Text = frmPartialAddTag.txtCohort.Text
    
    Unload frmPartialAddTag
End Sub

Private Sub cmdCancel_Click()
    Unload frmPartialAddTag
End Sub

Private Sub Form_Load()
    frmPartialAddTag.txtTagAdd.Locked = False
    frmPartialAddTag.txtTagAdd.Text = frmInput.txtTagCur2.Text
    frmPartialAddTag.txtTagAdd.Locked = True
    
    If frmDataSpec.chkPop.Value = 1 Then
        frmPartialAddTag.lblPop.Enabled = True
        frmPartialAddTag.cmbPop.Enabled = True
        Call frmInput.getPops("partial")
    End If
    
    If frmDataSpec.chkCohort.Value = 1 Then
        frmPartialAddTag.lblCohort.Enabled = True
        frmPartialAddTag.txtCohort.Enabled = True
    End If

End Sub
