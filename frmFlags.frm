VERSION 5.00
Begin VB.Form frmFlags 
   Caption         =   "Warning Flag Settings"
   ClientHeight    =   5910
   ClientLeft      =   7455
   ClientTop       =   4470
   ClientWidth     =   14070
   Icon            =   "frmFlags.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   14070
   Begin VB.CommandButton cmdViewQuarts 
      Caption         =   "View Values"
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
      Height          =   495
      Left            =   12000
      MouseIcon       =   "frmFlags.frx":048A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Click to view relatedness values corresponding to the 'Maximum Percentile' for each population"
      Top             =   1140
      Width           =   1575
   End
   Begin VB.TextBox txtMaxPerc 
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
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   $"frmFlags.frx":05DC
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtMinLoci 
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
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "The minimum number of loci that an individual must be scored for in order to avoid being flagged"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtYearsBack 
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
      Left            =   6960
      TabIndex        =   2
      ToolTipText     =   $"frmFlags.frx":0673
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcDist 
      Caption         =   "Calculate Relatedness Values"
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
      Height          =   495
      Left            =   8160
      MouseIcon       =   "frmFlags.frx":06FD
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   $"frmFlags.frx":084F
      Top             =   1140
      Width           =   3375
   End
   Begin VB.CheckBox chkDiffCohorts 
      Caption         =   "Individuals in a mating pair are from different cohorts"
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
      Left            =   600
      MouseIcon       =   "frmFlags.frx":08E1
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Tag             =   "0"
      ToolTipText     =   "Check to flag a mating where the female and male are from different cohorts"
      Top             =   4440
      Width           =   5775
   End
   Begin VB.CheckBox chkNoAlleles 
      Caption         =   "One or both individuals in a mating pair have allelic information for less then the minimum number of loci specified"
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
      Left            =   600
      MouseIcon       =   "frmFlags.frx":0A33
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "0"
      ToolTipText     =   "Check to flag a mating where either the female and male have allelic information for less than the 'Minimum Number of Loci'"
      Top             =   1800
      Width           =   12495
   End
   Begin VB.CheckBox chkPrevMating 
      Caption         =   "Individuals in a mating pair have previously reproduced"
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
      Left            =   600
      MouseIcon       =   "frmFlags.frx":0B85
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Tag             =   "0"
      ToolTipText     =   "Check to flag a mating where the female and male have previously been mated"
      Top             =   3000
      Width           =   6015
   End
   Begin VB.CheckBox chkUppQuart 
      Caption         =   "Individuals in a mating pair have a genetic relatedness above the specified percentile among alll mating pairs in the population"
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
      Left            =   600
      MouseIcon       =   "frmFlags.frx":0CD7
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   $"frmFlags.frx":0E29
      Top             =   600
      Width           =   13335
   End
   Begin VB.CheckBox chkDiffPops 
      Caption         =   "Individuals in a mating pair are from different populations"
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
      Left            =   600
      MouseIcon       =   "frmFlags.frx":0EEE
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Tag             =   "0"
      ToolTipText     =   "Check to flag a mating where the female and male are from different populations"
      Top             =   3720
      Width           =   6255
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
      Left            =   12480
      MouseIcon       =   "frmFlags.frx":1040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click to save current warning settings"
      Top             =   5280
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
      Left            =   240
      MouseIcon       =   "frmFlags.frx":1192
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Click to cancel changes made to the form"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblMaxPerc 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Percentile"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   1245
      Width           =   2175
   End
   Begin VB.Label lblMinLoci 
      Alignment       =   1  'Right Justify
      Caption         =   "Minimum Number of Loci"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   2445
      Width           =   2655
   End
   Begin VB.Label lblYearsBack 
      Alignment       =   1  'Right Justify
      Caption         =   "Cohorts to go Back"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1245
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Warn me if:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpBi As Long, tmpText As String, accessApp As Access.Application, j As Long, i As Long
Dim k As Long, m As Long, allNum As Long, tempAll1 As String, tempAll2 As String
Dim temp2All1 As String, temp2All2 As String, pShared As Long, tmpCnt As Long
Dim dbTemp As Database, rstTemp As Recordset, rstTemp2 As Recordset, rstTblTemp As Recordset
Dim tblTemp As TableDef, tmpField As Field, tmpProp As Property, a As Long, b As Long
Dim maxCohort() As Variant, popRelVal() As Variant, firstLocus As Long, popColumn As Long
Dim alleleCnt As Long, rstArray As Variant, rstArray2 As Variant, percVal As Double
Dim tmpVal As Double, tmpRec As Variant, tmpStr As String, strArray As Variant
Dim Msg, Style, Title, Response, tmpBool As Boolean

Private Sub chkNoAlleles_Click()
    If frmDataSpec.txtLoci.Text <> "" Then
        If frmFlags.chkNoAlleles.Value = 1 Then
            frmFlags.lblMinLoci.Enabled = True
            frmFlags.txtMinLoci.Enabled = True
        Else
            frmFlags.lblMinLoci.Enabled = False
            frmFlags.txtMinLoci.Enabled = False
        End If
    Else
        If frmFlags.chkNoAlleles.Value = 1 Then
            MsgBox "Please specify the number of loci under 'Database Settings' to enable use of this warning", vbInformation, "Loci Number Not Set"
            frmFlags.chkNoAlleles.Value = 0
        End If
    End If
End Sub

Private Sub chkUppQuart_Click()
    If frmFlags.chkUppQuart.Value = 1 Then
        frmFlags.lblMaxPerc.Enabled = True
        frmFlags.txtMaxPerc.Enabled = True
        If frmDataSpec.chkCohort.Value = 1 Then
            frmFlags.lblYearsBack.Enabled = True
            frmFlags.txtYearsBack.Enabled = True
        End If
        Call frmFlags.checkDistEnable("enable")
    Else
        frmFlags.lblMaxPerc.Enabled = False
        frmFlags.txtMaxPerc.Enabled = False
        frmFlags.lblYearsBack.Enabled = False
        frmFlags.txtYearsBack.Enabled = False
        frmFlags.cmdCalcDist.Enabled = False
        frmFlags.cmdViewQuarts.Enabled = False
    End If
End Sub

Private Sub cmdCalcDist_Click()
    Screen.MousePointer = vbHourglass
    Call frmFlags.calcRelVals
    Screen.MousePointer = vbDefault
    frmFlags.cmdViewQuarts.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Call doCancel
    frmFlags.Hide
End Sub

Private Sub cmdSave_Click()
    tmpBi = frmFlags.saveFlags
    If tmpBi <> 7 Then
        frmFlags.Hide
        If Dir(App.Path & "\dbase_settings.mmf") = "" Then
            frmMain.lblDbaseConnection.Caption = "Specify 'Database Settings' under the 'Options' menu"
            frmDataSpec.Show 1
        Else
            frmMain.cmdInputPIT.Enabled = True
            frmMain.lblDbaseConnection.Visible = False
        End If
    End If
End Sub

Private Sub cmdViewQuarts_Click()
    Call frmFlags.viewPercs
End Sub

Private Sub txtMaxPerc_Change()
    If (frmFlags.txtMaxPerc Is Me.ActiveControl) Then
        Call frmFlags.checkDistEnable("change")
    Else
        Call frmFlags.checkDistEnable("cancel")
    End If
End Sub

Private Sub txtMinLoci_Change()
    If (frmFlags.txtMinLoci Is Me.ActiveControl) Then
        Call frmFlags.checkMinLoci("change")
    Else
        Call frmFlags.checkMinLoci("cancel")
    End If
End Sub

Private Sub txtYearsBack_Change()
    If (frmFlags.txtYearsBack Is Me.ActiveControl) Then
        Call frmFlags.checkDistEnable("change")
    Else
        Call frmFlags.checkDistEnable("cancel")
    End If
End Sub

Public Function checkMinLoci(tmpType As String)
    If frmFlags.txtMinLoci.Text <> "" Then
        If IsNumeric(frmFlags.txtMinLoci.Text) = True Then
            If CLng(frmFlags.txtMinLoci.Text) < 0 Or CLng(frmFlags.txtMinLoci) > CLng(frmDataSpec.txtLoci.Text) Then
                MsgBox "Please enter an integer between 0 and " & frmDataSpec.txtLoci.Text & " inclusive", vbInformation, "Input Error"
                frmFlags.txtMinLoci.Text = ""
                frmFlags.txtMinLoci.SetFocus
            Else
                frmFlags.txtMinLoci.Text = CLng(frmFlags.txtMinLoci.Text)
            End If
        Else
            MsgBox "Please enter an integer between 0 and " & frmDataSpec.txtLoci.Text & " inclusive", vbInformation, "Input Error"
            frmFlags.txtMinLoci.Text = ""
            frmFlags.txtMinLoci.SetFocus
        End If
    End If
End Function

Public Function checkDistEnable(tmpType As String)
    If frmFlags.txtMaxPerc.Text <> "" Then
        If IsNumeric(frmFlags.txtMaxPerc.Text) = True Then
            If CSng(frmFlags.txtMaxPerc.Text) < 0 Or CSng(frmFlags.txtMaxPerc.Text) > 100 Then
                If tmpType = "change" Then
                    MsgBox "Please enter a number between 0 and 100 inclusive", vbInformation, "Input Error"
                    frmFlags.cmdCalcDist.Enabled = False
                    frmFlags.txtMaxPerc.Text = ""
                    frmFlags.txtMaxPerc.SetFocus
                End If
                Exit Function
            Else
                tmpBi = 1
            End If
        Else
            If tmpType = "change" Then
                MsgBox "Please enter a number between 0 and 100 inclusive", vbInformation, "Input Error"
                frmFlags.cmdCalcDist.Enabled = False
                frmFlags.txtMaxPerc.Text = ""
                frmFlags.txtMaxPerc.SetFocus
            End If
            Exit Function
        End If
    End If
    
    If frmFlags.txtYearsBack.Text <> "" And tmpBi = 1 Then
        If IsNumeric(frmFlags.txtYearsBack.Text) = True Then
            If CSng(frmFlags.txtYearsBack.Text) < 0 Then
                If tmpType = "change" Then
                    MsgBox "Please enter a number greater than or equal to 0", vbInformation, "Input Error"
                    frmFlags.cmdCalcDist.Enabled = False
                    frmFlags.txtYearsBack.Text = ""
                    frmFlags.txtYearsBack.SetFocus
                End If
                Exit Function
            Else
                tmpBi = frmDataSpec.checkDbase("quart")
                If tmpBi <> 7 Then
                    frmFlags.cmdCalcDist.Enabled = True
                    If Not Not popRelVal Then
                        frmFlags.cmdViewQuarts.Enabled = True
                    Else
                        frmFlags.cmdViewQuarts.Enabled = False
                    End If
                End If
            End If
        Else
            If tmpType = "change" Then
                MsgBox "Please enter a number greater than or equal to 0", vbInformation, "Input Error"
                frmFlags.cmdCalcDist.Enabled = False
                frmFlags.txtYearsBack.Text = ""
                frmFlags.txtYearsBack.SetFocus
            End If
            Exit Function
        End If
    End If

End Function

Public Function doCancel()
    frmFlags.chkUppQuart.Value = frmFlags.chkUppQuart.Tag
    frmFlags.txtMaxPerc.Text = frmFlags.txtMaxPerc.Tag
    frmFlags.txtYearsBack.Text = frmFlags.txtYearsBack.Tag
    frmFlags.chkNoAlleles.Value = frmFlags.chkNoAlleles.Tag
    frmFlags.txtMinLoci.Text = frmFlags.txtMinLoci.Tag
    frmFlags.chkPrevMating.Value = frmFlags.chkPrevMating.Tag
    frmFlags.chkDiffPops.Value = frmFlags.chkDiffPops.Tag
    frmFlags.chkDiffCohorts.Value = frmFlags.chkDiffCohorts.Tag
End Function

Public Function saveFlags() As Integer
    saveFlags = checkFlags("save")
    If saveFlags = 7 Then
        Exit Function
    End If
        
    Open App.Path & "\flag_settings.mmf" For Output As #1
        Print #1, CStr(frmFlags.chkUppQuart.Value)
        frmFlags.chkUppQuart.Tag = frmFlags.chkUppQuart.Value
        If frmFlags.chkUppQuart.Value = 1 Then
            Print #1, frmFlags.txtMaxPerc.Text
            frmFlags.txtMaxPerc.Tag = frmFlags.txtMaxPerc.Text
            
            Print #1, frmFlags.txtYearsBack.Text
            frmFlags.txtYearsBack.Tag = frmFlags.txtYearsBack.Text
        End If
        
        Print #1, CStr(frmFlags.chkNoAlleles.Value)
        frmFlags.chkNoAlleles.Tag = frmFlags.chkNoAlleles.Value
        If frmFlags.chkNoAlleles.Value = 1 Then
            Print #1, frmFlags.txtMinLoci.Text
            frmFlags.txtMinLoci.Tag = frmFlags.txtMinLoci.Text
        End If
        
        Print #1, CStr(frmFlags.chkPrevMating.Value)
        frmFlags.chkPrevMating.Tag = frmFlags.chkPrevMating.Value
            
        Print #1, CStr(frmFlags.chkDiffPops.Value)
        frmFlags.chkDiffPops.Tag = frmFlags.chkDiffPops.Value
            
        Print #1, CStr(frmFlags.chkDiffCohorts.Value)
        frmFlags.chkDiffCohorts.Tag = frmFlags.chkDiffCohorts.Value
    Close #1
    
    If frmFlags.chkUppQuart.Value = 1 Then
        Open App.Path & "\perc_settings.mmf" For Output As #1
            For k = 0 To UBound(popRelVal)
                Print #1, popRelVal(k, 0) & Chr(9) & popRelVal(k, 1)
            Next k
        Close #1
        Call frmInput.setPopRelVals(UBound(popRelVal), popRelVal)
    End If

End Function

Public Function checkFlags(tmpType As String) As Integer
    checkFlags = 6
    If frmFlags.chkUppQuart.Value = 1 Then
        If frmFlags.txtMaxPerc.Text = "" Then
            checkFlags = 7
            If tmpType = "save" Then
                MsgBox "Please enter a 'Maximum Percentile' value", vbInformation, "Missing Data"
                frmFlags.txtMaxPerc.SetFocus
            End If
            Exit Function
        End If
        
        If frmFlags.txtYearsBack.Text = "" And frmDataSpec.chkCohort.Value = 1 Then
            checkFlags = 7
            If tmpType = "save" Then
                MsgBox "Please enter a 'Years to go Back' value", vbInformation, "Missing Data"
                frmFlags.txtYearsBack.SetFocus
            End If
            Exit Function
        End If
        
        If ((Not Not popRelVal) = 0) Then
            checkFlags = 7
            If tmpType = "save" Then
                MsgBox "Please calculate relatedness values or uncheck the genetic relatedness flag", vbInformation, "Missing Data"
                frmFlags.cmdCalcDist.SetFocus
            End If
        End If
    End If
    
    If frmFlags.chkNoAlleles.Value = 1 Then
        If frmFlags.txtMinLoci.Text = "" Then
            checkFlags = 7
            If tmpType = "save" Then
                MsgBox "Please enter a 'Min Number of Loci' value", vbInformation, "Missing Data"
                frmFlags.txtMinLoci.SetFocus
            End If
            Exit Function
        End If
    End If
End Function

Public Function importFlagSettings() As Boolean
    On Error GoTo finishUp
    
    Open App.Path & "\flag_settings.mmf" For Input As #1
        Input #1, tmpText
        frmFlags.chkUppQuart.Value = CLng(tmpText)
        frmFlags.chkUppQuart.Tag = CLng(tmpText)
        If frmFlags.chkUppQuart.Value = 1 Then
            Input #1, tmpText
            frmFlags.txtMaxPerc.Text = tmpText
            frmFlags.txtMaxPerc.Tag = tmpText
        
            Input #1, tmpText
            frmFlags.txtYearsBack.Text = tmpText
            frmFlags.txtYearsBack.Tag = tmpText
        End If
        
        Input #1, tmpText
        frmFlags.chkNoAlleles.Value = CLng(tmpText)
        frmFlags.chkNoAlleles.Tag = CLng(tmpText)
        If frmFlags.chkNoAlleles.Value = 1 Then
            Input #1, tmpText
            frmFlags.txtMinLoci.Text = tmpText
            frmFlags.txtMinLoci.Tag = tmpText
        End If

        Input #1, tmpText
        frmFlags.chkPrevMating.Value = CLng(tmpText)
        frmFlags.chkPrevMating.Tag = CLng(tmpText)
        
        Input #1, tmpText
        frmFlags.chkDiffPops.Value = CLng(tmpText)
        frmFlags.chkDiffPops.Tag = CLng(tmpText)
        
        Input #1, tmpText
        frmFlags.chkDiffCohorts.Value = CLng(tmpText)
        frmFlags.chkDiffCohorts.Tag = CLng(tmpText)
    Close #1
    importFlagSettings = True
Exit Function

finishUp:
    Close #1
    importFlagSettings = False
End Function

Public Function importPercSettings() As Boolean
    On Error GoTo finishUp
    
    i = -1
    Open App.Path & "\perc_settings.mmf" For Input As #1
        Do Until EOF(1)
            Input #1, tmpText
            i = i + 1
        Loop
    Close #1
    
    ReDim popRelVal(i, 1)
    i = 0
    Open App.Path & "\perc_settings.mmf" For Input As #1
        Do Until EOF(1)
            Input #1, tmpText
            strArray = Split(tmpText, Chr(9))
            If UBound(strArray) = 0 Then
                popRelVal(i, 1) = strArray(0)
            Else
                popRelVal(i, 0) = strArray(0)
                popRelVal(i, 1) = strArray(1)
            End If
            i = i + 1
        Loop
    Close #1
    
    Call frmInput.setPopRelVals(UBound(popRelVal), popRelVal)
    If Not Not popRelVal Then
     frmFlags.cmdViewQuarts.Enabled = True
    End If
    'Call frmFlags.checkDistEnable("import")
    importPercSettings = True
Exit Function

finishUp:
    Close #1
    importPercSettings = False
End Function

Public Function viewPercs()
    If UBound(popRelVal) >= 0 Then
        tmpText = ""
        For i = 0 To UBound(popRelVal)
            tmpText = tmpText & Format(popRelVal(i, 1), "0.000") & Chr(9) & popRelVal(i, 0) & Chr(13)
        Next i
        MsgBox tmpText, , "Relatedness Thresholds"
    Else
        MsgBox "There are currently no calculated values present to view", vbInformation, "No Values Present"
    End If
End Function

Public Function calcRelVals()
    If Dir(App.Path & "\perc_settings.mmf") <> "" Then
        Msg = "Population specific relatedness threshold values already exist, do you want to overwrite them?"
        Style = vbYesNo + vbQuestion + vbDefaultButton2
        Title = "Overwrite Existing Values"
        Response = MsgBox(Msg, Style, Title)
    
        If Response = 7 Then
            Exit Function
        End If
    End If
    
    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase frmDataSpec.txtDBFile.ToolTipText, False
    Set dbTemp = accessApp.CurrentDb
    
    'Get max cohort
    If frmDataSpec.chkCohort.Value = 1 Then
        If frmDataSpec.chkPop.Value = 1 Then
            Set rstTemp = dbTemp.OpenRecordset("SELECT Max([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "]) AS MaxOf" & frmDataSpec.cmbCohort.Text & ", [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] GROUP BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] HAVING (((Max([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "])) Is Not Null)) ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "];", dbOpenDynaset)
        Else
            Set rstTemp = dbTemp.OpenRecordset("SELECT Max([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "]) AS MaxOf" & frmDataSpec.cmbCohort.Text & " FROM [" & frmDataSpec.cmbGeneticsTable.Text & "] HAVING (((Max([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "])) Is Not Null));", dbOpenDynaset)
        End If
        
        rstTemp.MoveLast
        rstTemp.MoveFirst
        maxCohort = rstTemp.GetRows(rstTemp.RecordCount)
        Set rstTemp = Nothing
    End If
    
    'Get populations
    If frmDataSpec.chkPop.Value = 1 Then
        Set rstTemp = dbTemp.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] From [" & frmDataSpec.cmbGeneticsTable.Text & "] GROUP BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] HAVING ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "]) Is Not Null)) ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "];", dbOpenDynaset)
        rstTemp.MoveLast
        rstTemp.MoveFirst
        ReDim popRelVal(rstTemp.RecordCount - 1, 1)
        i = 0
        Do Until rstTemp.EOF
            popRelVal(i, 0) = rstTemp.Fields(frmDataSpec.cmbPop.Text).Value
            i = i + 1
            rstTemp.MoveNext
        Loop
        Set rstTemp = Nothing
    Else
        ReDim popRelVal(0, 1)
    End If
    
    Select Case frmMain.cmbRelMetric.Text
        Case "Proportion of Shared Alleles"
            Call frmFlags.calcPSA
    End Select


    dbTemp.Close
    accessApp.CloseCurrentDatabase
    'Set accessApp = Nothing
    
    Msg = "The new quartile values have been calculated for all drainages."
    Style = vbOKOnly + vbInformation + vbDefaultButton1
    Title = "Quartiles Calculated"
    Response = MsgBox(Msg, Style, Title)
End Function

Public Function calcPSA()
    If frmDataSpec.chkPop.Value = 1 And frmDataSpec.chkCohort.Value = 1 Then
        tmpText = ""
        For k = 0 To UBound(maxCohort, 2)
            If tmpText = "" Then
                tmpText = "((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "]) = '" & maxCohort(1, k) & "') AND (([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "]) >= " & (maxCohort(0, k) - CLng(frmFlags.txtYearsBack.Text)) & "))"
            Else
                tmpText = tmpText & " OR ((([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "]) = '" & maxCohort(1, k) & "') AND (([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "]) >= " & (maxCohort(0, k) - CLng(frmFlags.txtYearsBack.Text)) & "))"
            End If
        Next k
        
        Set rstTemp = dbTemp.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* From [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE " & tmpText & " ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "], [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "];", dbOpenDynaset)
    ElseIf frmDataSpec.chkPop.Value = 1 Then
        Set rstTemp = dbTemp.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* From [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "] <> '') ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbPop.Text & "];", dbOpenDynaset)
    ElseIf frmDataSpec.chkCohort.Value = 1 Then
        Set rstTemp = dbTemp.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* From [" & frmDataSpec.cmbGeneticsTable.Text & "] WHERE ([" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "] >=" & (maxCohort(0, 0) - CLng(frmFlags.txtYearsBack.Text)) & ") ORDER BY [" & frmDataSpec.cmbGeneticsTable.Text & "].[" & frmDataSpec.cmbCohort.Text & "];", dbOpenDynaset)
    Else
        Set rstTemp = dbTemp.OpenRecordset("SELECT [" & frmDataSpec.cmbGeneticsTable.Text & "].* From [" & frmDataSpec.cmbGeneticsTable.Text & "];", dbOpenDynaset)
    End If
    
    If rstTemp.RecordCount > 0 Then
    rstTemp.MoveLast
    rstTemp.MoveFirst
    
    rstArray = rstTemp.GetRows(rstTemp.RecordCount)  'first dimension is columns, second is rows
    rstTemp.MoveFirst
    rstArray2 = rstTemp.GetRows(rstTemp.RecordCount)
       
    On Error Resume Next
    dbTemp.TableDefs.Delete "Temp"
    On Error GoTo 0

    Set tblTemp = dbTemp.CreateTableDef("Temp")
    With tblTemp
        tblTemp.Fields.Append .CreateField("Population", dbText)
        tblTemp.Fields.Append .CreateField("PropShared", dbSingle)
    End With
    dbTemp.TableDefs.Append tblTemp
        
    Set rstTblTemp = dbTemp.OpenRecordset("SELECT Temp.* FROM Temp", dbOpenDynaset)
        
    firstLocus = frmDataSpec.cmbFirstLocus.ListIndex
    alleleCnt = CLng(frmDataSpec.txtLoci.Text)
    popColumn = frmDataSpec.cmbPop.ListIndex
    
    For k = 0 To rstTemp.RecordCount - 1
        j = 0
        For i = firstLocus To (firstLocus + ((alleleCnt * 2) - 1))
            If IsNull(rstArray(i, k)) = False Then
                j = j + 1
            End If
        Next i
                
        If j > 0 Then
            For a = k + 1 To rstTemp.RecordCount - 1
                If frmDataSpec.chkPop.Value = 1 Then
                    If rstArray(popColumn, k) = rstArray2(popColumn, a) Then
                        tmpBi = 1
                    Else
                        tmpBi = 0
                    End If
                Else
                    tmpBi = 1
                End If
                
                If tmpBi = 1 Then
                    m = 0
                    For i = firstLocus To (firstLocus + ((alleleCnt * 2) - 1))
                        If IsNull(rstArray2(i, a)) = False Then
                            m = m + 1
                        End If
                    Next i
                            
                    If m > 0 Then
                        'calculate relatedness
                        allNum = 0: pShared = 0
                        For i = firstLocus To (firstLocus + ((alleleCnt * 2) - 1))
                            If IsNull(rstArray(i, k)) = False Then
                                tempAll1 = rstArray(i, k)
                                i = i + 1
                                tempAll2 = rstArray(i, k)
                                i = i - 1
                                If IsNull(rstArray2(i, a)) = False Then
                                    temp2All1 = rstArray2(i, a)
                                    i = i + 1
                                    temp2All2 = rstArray2(i, a)
                                                
                                    allNum = allNum + 2
                                                
                                    If tempAll1 = tempAll2 Then
                                        If tempAll1 = temp2All1 And tempAll1 = temp2All2 Then
                                            pShared = pShared + 2
                                        ElseIf tempAll1 = temp2All1 Or tempAll1 = temp2All2 Then
                                            pShared = pShared + 1
                                        End If
                                    ElseIf temp2All1 = temp2All2 Then
                                        If tempAll1 = temp2All1 And tempAll2 = temp2All1 Then
                                            pShared = pShared + 2
                                        ElseIf tempAll1 = temp2All1 Or tempAll2 = temp2All1 Then
                                            pShared = pShared + 1
                                        End If
                                    Else
                                        If tempAll1 = temp2All1 Or tempAll1 = temp2All2 Then
                                            pShared = pShared + 1
                                        End If
                        
                                        If tempAll2 = temp2All1 Or tempAll2 = temp2All2 Then
                                            pShared = pShared + 1
                                        End If
                                    End If
                                Else
                                    i = i + 1
                                End If
                            Else
                                i = i + 1
                            End If
                        Next i
                                    
                        rstTblTemp.AddNew
                        If frmDataSpec.chkPop.Value = 1 Then
                            rstTblTemp!Population = rstArray(popColumn, k)
                        End If
                            
                        If allNum > 0 Then
                            rstTblTemp!PropShared = Format(pShared / allNum, "0.000")
                        End If
                        rstTblTemp.Update
                    End If
                Else
                    Exit For
                End If
            Next a
        End If
    Next k
    
    percVal = CDbl(frmFlags.txtMaxPerc.Text) / 100
    For k = 0 To UBound(popRelVal)
        If frmDataSpec.chkPop.Value = 1 Then
            Set rstTemp = dbTemp.OpenRecordset("SELECT Temp.PropShared From Temp Where (((Temp.Population) = '" & popRelVal(k, 0) & "')) ORDER BY Temp.PropShared;", dbOpenDynaset)
        Else
            Set rstTemp = dbTemp.OpenRecordset("SELECT Temp.PropShared From Temp ORDER BY Temp.PropShared;", dbOpenDynaset)
        End If
        If rstTemp.RecordCount > 1 Then
            rstTemp.MoveLast
            rstTemp.MoveFirst
            tmpRec = rstTemp.GetRows(rstTemp.RecordCount)
            tmpVal = Round(percVal * rstTemp.RecordCount, 0)
            popRelVal(k, 1) = tmpRec(0, tmpVal)
        Else
            popRelVal(k, 1) = 0
        End If
    Next k
    
    Set rstTemp = Nothing
    Set rstTblTemp = Nothing
    dbTemp.TableDefs.Delete "Temp"
    End If
End Function

