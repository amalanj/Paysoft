VERSION 5.00
Begin VB.Form frmGradeMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grade Master Maintenance"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Cef_frame 
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   2160
      Width           =   2655
      Begin VB.OptionButton optPerDay 
         Caption         =   "Per Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optPerMonth 
         Caption         =   "Per Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboSno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboDesignation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox txtCEF 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtDA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtBasic 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtGrade 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1261
      TabIndex        =   8
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2341
      TabIndex        =   9
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3421
      TabIndex        =   10
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&l"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4523
      TabIndex        =   11
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5581
      TabIndex        =   12
      Top             =   3150
      Width           =   1095
   End
   Begin VB.TextBox txtDesignation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      MaxLength       =   100
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label6 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   2325
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "CEF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2325
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Dearness Allownace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1845
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Basic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Grade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   405
      Width           =   855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   1235
      Top             =   3120
      Width           =   5475
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   240
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   885
      Width           =   1215
   End
End
Attribute VB_Name = "frmGradeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_grade_m As ADODB.Recordset
Dim ls_grade As String
Dim ls_designation As String

Private Sub cboDesignation_Click()
    If cboDesignation.Text <> "" Then
        cboSno.ListIndex = cboDesignation.ListIndex
        wf_set_data
        If cmdEdit.Caption = "&Update" Then
            txtGrade.SetFocus
        End If
        If cmdDelete.Caption = "Confir&m" Then
            txtGrade.Locked = True
            txtDesignation.Locked = True
            txtBasic.Locked = True
            txtDA.Locked = True
            txtCEF.Locked = True
            Cef_frame.Enabled = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    wf_clear_data
    cboDesignation.Visible = False
    cboSno.Visible = False
    ls_grade = ""
    ls_designation = ""
    txtGrade.Locked = False
    txtDesignation.Locked = False
    txtBasic.Locked = False
    txtDA.Locked = False
    txtCEF.Locked = False
    Cef_frame.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Caption = "&Save"
    cmdEdit.Caption = "&Edit"
    cmdDelete.Caption = "&Delete"
    gf_mdi_message "Ready"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub wf_clear_data()
    txtGrade.Text = ""
    txtDesignation.Text = ""
    txtBasic.Text = ""
    txtDA.Text = ""
    txtCEF.Text = ""
    optPerDay.Value = False
    optPerMonth.Value = False
End Sub

Private Sub cmdDelete_Click()
    If cmdDelete.Caption = "&Delete" Then
        wf_clear_data
        cboDesignation.Visible = True
        cboDesignation.Clear
        cboSno.Clear
        cmdEdit.Enabled = False
        cmdSave.Enabled = False
        cmdDelete.Caption = "Confir&m"
        Set rs_grade_m = New ADODB.Recordset
        rs_grade_m.Open "select * from grademast order by grade,designation", cn, adOpenKeyset, adLockOptimistic
        If rs_grade_m.RecordCount > 0 Then
            rs_grade_m.MoveFirst
            Do While Not rs_grade_m.EOF
                cboDesignation.AddItem rs_grade_m![designation] & "-" & rs_grade_m![grade]
                cboSno.AddItem rs_grade_m![sno]
                rs_grade_m.MoveNext
            Loop
        End If
    ElseIf cmdDelete.Caption = "Confir&m" And Len(Trim(txtGrade.Text)) > 0 Then
        If MsgBox("Do you want to delete?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            gf_execute_sql ("delete from grademast where sno=" & Val(cboSno.Text))
            gf_mdi_message "Deleted"
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        wf_clear_data
        cboDesignation.Visible = True
        cboDesignation.Clear
        cboSno.Clear
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        cmdEdit.Caption = "&Update"
        Set rs_grade_m = New ADODB.Recordset
        rs_grade_m.Open "select * from grademast order by grade,designation", cn, adOpenKeyset, adLockOptimistic
        If rs_grade_m.RecordCount > 0 Then
            rs_grade_m.MoveFirst
            Do While Not rs_grade_m.EOF
                cboDesignation.AddItem rs_grade_m![designation] & "-" & rs_grade_m![grade]
                cboSno.AddItem rs_grade_m![sno]
                rs_grade_m.MoveNext
            Loop
        End If
        cboDesignation.SetFocus
    ElseIf cmdEdit.Caption = "&Update" And Len(Trim(txtGrade.Text)) > 0 Then
        If MsgBox("Do you wish to update?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            If wf_valid_data = False Then Exit Sub
            
            Set rs_grade_m = New ADODB.Recordset
            rs_grade_m.Open "select * from grademast where sno=" & Val(cboSno.Text), cn, adOpenKeyset, adLockOptimistic
            If rs_grade_m.RecordCount > 0 Then
                rs_grade_m![grade] = txtGrade.Text
                rs_grade_m![designation] = txtDesignation.Text
                rs_grade_m![basic] = Val(txtBasic.Text)
                rs_grade_m![da] = Val(txtDA.Text)
                rs_grade_m![cefamount] = Val(txtCEF.Text)
                rs_grade_m.Update
                gf_mdi_message "Updated"
                cmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If wf_valid_data = False Then Exit Sub
    
    Set rs_grade_m = New ADODB.Recordset
    rs_grade_m.Open "select * from grademast where grade='" & Trim(txtGrade.Text) & "' and designation = '" & Trim(txtDesignation.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    If rs_grade_m.RecordCount > 0 Then
        MsgBox ("Cannot save!..Grade with Designation already exists")
        txtGrade.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Do you wish to save?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
        Set rs_grade_m = New ADODB.Recordset
        rs_grade_m.Open "select * from grademast", cn, adOpenKeyset, adLockOptimistic
        
        rs_grade_m.AddNew
        rs_grade_m![grade] = txtGrade.Text
        rs_grade_m![designation] = txtDesignation.Text
        rs_grade_m![basic] = Val(txtBasic.Text)
        rs_grade_m![da] = Val(txtDA.Text)
        If optPerDay.Value Then
            rs_grade_m![cefoption] = "D"
        ElseIf optPerMonth.Value Then
            rs_grade_m![cefoption] = "M"
        End If
        If txtCEF.Text <> 0 Then
            rs_grade_m![cefamount] = Val(txtCEF.Text)
        Else
            rs_grade_m![cefamount] = 0
        End If
        rs_grade_m.Update
        gf_mdi_message "Saved"
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    gf_mdi_message "Ready"
End Sub

Sub wf_set_data()
    Set rs_grade_m = New ADODB.Recordset
    rs_grade_m.Open "select * from grademast where sno=" & cboSno.Text, cn, adOpenKeyset, adLockOptimistic
    If rs_grade_m.RecordCount > 0 Then
        txtGrade.Text = rs_grade_m![grade]
        txtDesignation.Text = rs_grade_m![designation]
        txtBasic.Text = rs_grade_m![basic]
        txtDA.Text = rs_grade_m![da]
        If rs_grade_m![cefoption] = "D" Then
            optPerDay.Value = True
        ElseIf rs_grade_m![cefoption] = "M" Then
            optPerMonth.Value = True
        End If
        txtCEF.Text = rs_grade_m![cefamount]
    End If
    cboDesignation.Visible = False
End Sub

Private Sub txtBasic_KeyPress(KeyAscii As Integer)
    gf_accept_only_number KeyAscii
End Sub

Private Sub txtCEF_KeyPress(KeyAscii As Integer)
    gf_accept_only_number KeyAscii
End Sub

Private Sub txtDA_KeyPress(KeyAscii As Integer)
    gf_accept_only_number KeyAscii
End Sub

Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    gf_accept_alpha_numeric KeyAscii
End Sub

Function wf_valid_data() As Boolean
    If Len(Trim(txtGrade.Text)) <= 0 Then
        MsgBox "Grade should not be blank", vbInformation, "PAYSOFT"
        txtGrade.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If Len(Trim(txtDesignation.Text)) <= 0 Then
        MsgBox "Designation should not be blank", vbInformation, "PAYSOFT"
        txtDesignation.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If Len(Trim(txtBasic.Text)) <= 0 Then
        MsgBox "Basic should not be blank", vbInformation, "PAYSOFT"
        txtBasic.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If Len(Trim(txtDA.Text)) <= 0 Then
        MsgBox "DA should not be blank", vbInformation, "PAYSOFT"
        txtDA.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    wf_valid_data = True
End Function
