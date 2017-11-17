VERSION 5.00
Begin VB.Form frmDataSpec 
   Caption         =   "Database Specifications"
   ClientHeight    =   7425
   ClientLeft      =   8295
   ClientTop       =   3000
   ClientWidth     =   9645
   Icon            =   "frmDataSpec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9645
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
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Closes form without adding the current individual to 'tblbrood'"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   8040
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmDataSpec.Hide
End Sub

Private Sub cmdSave_Click()
    frmDataSpec.Hide
End Sub
