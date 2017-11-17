VERSION 5.00
Begin VB.Form frmAddReason 
   Caption         =   "Add Reason"
   ClientHeight    =   1395
   ClientLeft      =   4695
   ClientTop       =   2475
   ClientWidth     =   4575
   Icon            =   "frmAddReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4575
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
      ItemData        =   "frmAddReason.frx":0CCA
      Left            =   240
      List            =   "frmAddReason.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdAddReason 
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
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Adds comment to corresponding spawning pair"
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
      TabIndex        =   1
      ToolTipText     =   "Closes the form without saving the comment"
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
Dim Msg, Style, Title, Response, wrkJet As Workspace, dbsNew As Database, rstTemp As Recordset
Public tmpTagReason As String

Private Sub cmdAddReason_Click()
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    With dbsNew
        Set rstTemp = dbsNew.OpenRecordset("SELECT tblBrood.Comments From tblBrood WHERE (((tblBrood.Mark)='" & tmpTagReason & "'));", dbOpenDynaset)
        With rstTemp
            rstTemp.MoveFirst
            rstTemp.Edit
            If IsNull(!Comments) = True Then
                !Comments = Me.cmbReason.Text
            Else
                !Comments = !Comments & ", " & Me.cmbReason.Text
            End If
            rstTemp.Update
        End With
    End With
        
    dbsNew.Close
    wrkJet.Close

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
