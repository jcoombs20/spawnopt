VERSION 5.00
Begin VB.Form frmPartialAddTag 
   Caption         =   "Tag Information"
   ClientHeight    =   1965
   ClientLeft      =   3510
   ClientTop       =   2730
   ClientWidth     =   8265
   Icon            =   "frmPartialAddTag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtYear 
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
      Height          =   465
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "PIT tag of individual that is currently being spawned"
      Top             =   600
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
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Closes form without adding the current individual's information to the 'Input Form'"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddTag 
      Caption         =   "Add Tag"
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
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Adds the above information to the 'Input Form'"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cmbGender 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPartialAddTag.frx":0CCA
      Left            =   4560
      List            =   "frmPartialAddTag.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Gender of individual that is currently being spawned"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtTagAdd 
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
      Height          =   465
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "PIT tag of individual that is currently being spawned"
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbDrainage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPartialAddTag.frx":0CE6
      Left            =   2400
      List            =   "frmPartialAddTag.frx":0CFF
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Drainage of origin of individual that is currently being spawned"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PIT Tag"
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
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Year"
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
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Drainage"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
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
    If frmPartialAddTag.cmbDrainage.Text = "" Then
        Msg = "Please fill in the 'Drainage' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Drainage"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.cmbDrainage.SetFocus
        Exit Sub
    End If

    If frmPartialAddTag.cmbGender.Text = "" Then
        Msg = "Please fill in the 'Gender' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Gender"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.cmbGender.SetFocus
        Exit Sub
    End If
    
    If frmPartialAddTag.txtYear.Text = "" Then
        Msg = "Please fill in the 'Capture Year' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Year"
        Response = MsgBox(Msg, Style, Title)
        
        frmPartialAddTag.txtYear.SetFocus
        Exit Sub
    End If
    
    frmInput.txtDrainage.Text = frmPartialAddTag.cmbDrainage.Text
    
    frmInput.txtSex.Text = frmPartialAddTag.cmbGender.Text
    
    frmInput.txtYear.Text = frmPartialAddTag.txtYear.Text
    
    Unload frmPartialAddTag
End Sub

Private Sub cmdCancel_Click()
    Unload frmPartialAddTag
End Sub

Private Sub Form_Load()
    frmPartialAddTag.txtTagAdd.Locked = False
    frmPartialAddTag.txtTagAdd.Text = frmInput.txtTagCur2.Text
    frmPartialAddTag.txtTagAdd.Locked = True
End Sub
