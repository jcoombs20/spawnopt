VERSION 5.00
Begin VB.Form frmComment 
   Caption         =   "Add Comment"
   ClientHeight    =   2655
   ClientLeft      =   4695
   ClientTop       =   2475
   ClientWidth     =   5775
   Icon            =   "Comment Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Comment"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Closes the form without saving the comment"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Comment to add to corresponding spawning pair"
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdAddComment 
      Caption         =   "Add Comment"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "Adds comment to corresponding spawning pair"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Closes the form without saving the comment"
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg, Style, Title, Response

Private Sub cmdAddComment_Click()
    If frmComment.txtComment.Text = "" Then
        Msg = "There is no comment to add." & Chr(13) & Chr(13) & "Please either enter a comment or cancel the form."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "No Comment"
        Response = MsgBox(Msg, Style, Title)
        
        frmComment.txtComment.SetFocus
        Exit Sub
    End If
    Call frmMain.addComment
    Unload frmComment
End Sub

Private Sub cmdCancel_Click()
    Unload frmComment
End Sub

Private Sub cmdRemove_Click()
    Call frmMain.removeComment
    Unload frmComment
End Sub

Private Sub Form_Load()
    Call frmMain.loadComment
End Sub
