VERSION 5.00
Begin VB.Form frmAddReason 
   Caption         =   "Add Reason"
   ClientHeight    =   1395
   ClientLeft      =   4695
   ClientTop       =   2475
   ClientWidth     =   5955
   Icon            =   "frmAddReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbReason 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmAddReason.frx":048A
      Left            =   240
      List            =   "frmAddReason.frx":0494
      TabIndex        =   0
      ToolTipText     =   "Reason for removal of indivual from the mating pool"
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdAddReason 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Reason"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Adds reason for removing individual from the mating pool to its record in the genetics table"
      Top             =   840
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
      Height          =   465
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Closes the form without adding the reason"
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg, Style, Title, Response, accessApp As Access.Application, dbsNew As Database, rstTemp As Recordset
Public tmpTagReason As String, tmpReason As Variant, i As Long, j As Long, tmpReasons As Variant

Private Sub cmdAddReason_Click()
    Call frmInput.addReason(tmpTagReason, Me.cmbReason.Text)
        
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    tmpReasons = frmInput.fillReasons
    
    Me.cmbReason.Clear
    For i = 0 To UBound(tmpReasons, 2)
        Me.cmbReason.List(i) = tmpReasons(0, i)
    Next i
End Sub
