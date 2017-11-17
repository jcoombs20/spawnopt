VERSION 5.00
Begin VB.Form frmAddTag 
   Caption         =   "Add Tag to 'tblbrood'"
   ClientHeight    =   5175
   ClientLeft      =   2925
   ClientTop       =   2565
   ClientWidth     =   8250
   Icon            =   "Add Tag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbHatchery 
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
      ItemData        =   "Add Tag.frx":0CCA
      Left            =   6360
      List            =   "Add Tag.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   "Hatchery location where fish was tagged"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Comment for individual to be added"
      Top             =   3360
      Width           =   7935
   End
   Begin VB.CommandButton cmdAddTag 
      Caption         =   "Add Tag"
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
      Left            =   5880
      TabIndex        =   13
      ToolTipText     =   "Adds the above information to 'tblbrood' for the current individual"
      Top             =   4560
      Width           =   1335
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
      Left            =   960
      TabIndex        =   14
      ToolTipText     =   "Closes form without adding the current individual to 'tblbrood'"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
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
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Date that individual to be added was tagged"
      Top             =   2040
      Width           =   1815
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
      ItemData        =   "Add Tag.frx":0CF1
      Left            =   3480
      List            =   "Add Tag.frx":0CFB
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Gender of individual to be added"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cmbOrigin 
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Origin of individual to be added"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtCapYear 
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
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Year that the individual to be added was captured"
      Top             =   720
      Width           =   975
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
      ItemData        =   "Add Tag.frx":0D0D
      Left            =   2400
      List            =   "Add Tag.frx":0D26
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Drainage of origin of individual to be added"
      Top             =   720
      Width           =   1815
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
      TabIndex        =   0
      ToolTipText     =   "PIT tag of individual that is being added to 'tblbrood' table"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hatchery"
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
      Left            =   6480
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
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
      Left            =   3360
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   1440
      Width           =   855
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
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Origin"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
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
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   2175
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
      TabIndex        =   2
      Top             =   120
      Width           =   1575
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
      TabIndex        =   1
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
    If frmAddTag.cmbDrainage.Text = "" Then
        Msg = "Please fill in the 'Drainage' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Drainage"
        Response = MsgBox(Msg, Style, Title)
        
        frmAddTag.cmbDrainage.SetFocus
        Exit Sub
    End If

    If frmAddTag.cmbGender.Text = "" Then
        Msg = "Please fill in the 'Gender' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Gender"
        Response = MsgBox(Msg, Style, Title)
        
        frmAddTag.cmbGender.SetFocus
        Exit Sub
    End If
    
    If frmAddTag.cmbHatchery.Text = "" Then
        Msg = "Please fill in the 'Hatchery' field."
        Style = vbOKOnly + vbCritical + vbDefaultButton1
        Title = "Missing Hatchery"
        Response = MsgBox(Msg, Style, Title)
    
        frmAddTag.cmbHatchery.SetFocus
        Exit Sub
    End If
    
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    With dbsNew
        Set rstTemp = dbsNew.OpenRecordset("SELECT tblBrood.* FROM tblBrood", dbOpenDynaset)
        With rstTemp
            rstTemp.AddNew
                !BroodID = UCase(frmAddTag.txtTagAdd.Text) & "-" & frmAddTag.txtCapYear.Text
                
                Select Case frmAddTag.cmbDrainage.ListIndex + 1
                    Case 2
                        !Drainage = 11
                    Case Else
                        !Drainage = frmAddTag.cmbDrainage.ListIndex + 1
                End Select
                
                Select Case frmAddTag.cmbHatchery.ListIndex + 1
                    Case 0
                        !Hatchery = "CB"
                    Case 1
                        !Hatchery = "GL"
                End Select
                
                !CaptureYear = CLng(frmAddTag.txtCapYear.Text)
                !OriginID = frmAddTag.cmbOrigin.ListIndex + 1
                !MarkDate = Format(frmAddTag.txtDate.Text, "MM/DD/YY")
                !Mark = UCase(frmInput.txtTagCur2.Text)
                !MarkCode = "PIT"
                If frmAddTag.cmbGender.Text = "Male" Then
                    !Gender = "M"
                ElseIf frmAddTag.cmbGender.Text = "Female" Then
                    !Gender = "F"
                Else
                    !Gender = "U"
                End If
                !Active = "YES"
                !LossTypeID = Null
                !Comments = frmAddTag.txtComment.Text
                
                'q = 0
                'Do Until empName(q) = frmAddTag.cmbName.Text
                '    q = q + 1
                'Loop
                '!TableEditor = empID(q)
                
                !TableEditDate = Date
            rstTemp.Update
        End With
    End With
    
    dbsNew.Close
    wrkJet.Close
    
    frmInput.txtSex.Text = frmAddTag.cmbGender.Text
    frmInput.txtDrainage.Text = frmAddTag.cmbDrainage.Text
    frmInput.txtYear.Text = frmAddTag.txtCapYear.Text
    
    Unload frmAddTag
    
End Sub

Private Sub cmdCancel_Click()
    Unload frmAddTag
End Sub

Private Sub Form_Load()
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsNew = wrkJet.OpenDatabase("C:\Databases\MaineBroodstock.mdb")
    
    With dbsNew
        'Set rstTemp = dbsNew.OpenRecordset("SELECT CodeEmployee.* From CodeEmployee WHERE (((CodeEmployee.LastName)='King')) OR (((CodeEmployee.LastName)='Craig')) OR (((CodeEmployee.LastName)='Buckley')) OR (((CodeEmployee.LastName)='Tozier')) OR (((CodeEmployee.LastName)='Thies'))", dbOpenDynaset)
        'With rstTemp
        '    i = 0
        '    rstTemp.MoveFirst
        '    Do Until rstTemp.EOF
        '        frmAddTag.cmbName.AddItem !LastName & ", " & !FirstName
        '        empID(i) = rstTemp!EmployeeID
        '        empName(i) = !LastName & ", " & !FirstName
        '        i = i + 1
        '        rstTemp.MoveNext
        '    Loop
        'End With
        
        Set rstTemp = dbsNew.OpenRecordset("SELECT CodeOrigin.* FROM CodeOrigin", dbOpenDynaset)
        With rstTemp
            rstTemp.MoveFirst
            Do Until rstTemp.EOF
                frmAddTag.cmbOrigin.AddItem !Origin
                rstTemp.MoveNext
            Loop
        End With
    End With

    dbsNew.Close
    wrkJet.Close
    
    frmAddTag.txtTagAdd.Locked = False
    frmAddTag.txtTagAdd.Text = frmInput.txtTagCur2.Text
    frmAddTag.txtTagAdd.Locked = True
    
    frmAddTag.txtCapYear = Format(Date, "YYYY")
    frmAddTag.txtDate = Date
End Sub
