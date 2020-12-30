VERSION 5.00
Begin VB.Form frmBranchMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Branch Master Maintenance"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboBranchCode 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
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
      Left            =   5400
      TabIndex        =   12
      Top             =   3240
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
      Left            =   4320
      TabIndex        =   11
      Top             =   3240
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
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
      Left            =   2160
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtBranchCode 
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
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtAddress3 
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
      Left            =   2040
      MaxLength       =   250
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox txtAddress2 
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
      Left            =   2040
      MaxLength       =   250
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox txtAddress1 
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
      Left            =   2040
      MaxLength       =   250
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtBranchName 
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
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   1050
      Top             =   3210
      Width           =   5480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   240
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "Branch Code / Short Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Branch Address"
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
      TabIndex        =   7
      Top             =   1605
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Branch Name"
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
      TabIndex        =   5
      Top             =   1125
      Width           =   1215
   End
End
Attribute VB_Name = "frmBranchMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_branch_m As ADODB.Recordset
Dim ls_branch_code As String

Private Sub cboBranchCode_Click()
    If Len(Trim(cboBranchCode.Text)) > 0 Then
        wf_set_data
        If cmdEdit.Caption = "&Update" Then
            txtBranchCode.Locked = True
            txtBranchName.SetFocus
        End If
        If cmdDelete.Caption = "Confir&m" Then
            txtBranchCode.Locked = True
            txtBranchName.Locked = True
            txtAddress1.Locked = True
            txtAddress2.Locked = True
            txtAddress3.Locked = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    ls_branch_code = ""
    wf_clear_data
    txtBranchCode.Locked = False
    txtBranchName.Locked = False
    txtAddress1.Locked = False
    txtAddress2.Locked = False
    txtAddress3.Locked = False
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Caption = "&Save"
    cmdEdit.Caption = "&Edit"
    cmdDelete.Caption = "&Delete"
    cboBranchCode.Visible = False
    gf_mdi_message "Ready"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmdDelete.Caption = "&Delete" Then
        wf_clear_data
        cboBranchCode.Visible = True
        cboBranchCode.Clear
        cmdEdit.Enabled = False
        cmdSave.Enabled = False
        cmdDelete.Caption = "Confir&m"
        Set rs_branch_m = New ADODB.Recordset
        rs_branch_m.Open "select * from branchmast order by branchname", cn, adOpenKeyset, adLockOptimistic
        If rs_branch_m.RecordCount > 0 Then
            rs_branch_m.MoveFirst
            Do While Not rs_branch_m.EOF
                cboBranchCode.AddItem rs_branch_m![branchname] & "-" & rs_branch_m![branchcode]
                rs_branch_m.MoveNext
            Loop
        End If
    ElseIf cmdDelete.Caption = "Confir&m" And Len(Trim(txtBranchCode.Text)) > 0 Then
        If MsgBox("Do you want to delete?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            gf_execute_sql ("delete from branchmast where branchcode='" & txtBranchCode.Text & "'")
            gf_mdi_message "Deleted"
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        wf_clear_data
        cboBranchCode.Visible = True
        cboBranchCode.Clear
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        cmdEdit.Caption = "&Update"
        Set rs_branch_m = New ADODB.Recordset
        rs_branch_m.Open "select * from branchmast order by branchname", cn, adOpenKeyset, adLockOptimistic
        If rs_branch_m.RecordCount > 0 Then
            rs_branch_m.MoveFirst
            Do While Not rs_branch_m.EOF
                cboBranchCode.AddItem rs_branch_m![branchname] & "-" & rs_branch_m![branchcode]
                rs_branch_m.MoveNext
            Loop
        End If
        cboBranchCode.SetFocus
    ElseIf cmdEdit.Caption = "&Update" And Len(Trim(txtBranchCode.Text)) > 0 Then
        If MsgBox("Do you wish to update?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            If wf_valid_data = False Then Exit Sub
            Set rs_branch_m = New ADODB.Recordset
            rs_branch_m.Open "select * from branchmast where branchcode='" & txtBranchCode.Text & "'", cn, adOpenKeyset, adLockOptimistic
            If rs_branch_m.RecordCount > 0 Then
                rs_branch_m![branchname] = txtBranchName.Text
                rs_branch_m![address1] = txtAddress1.Text
                rs_branch_m![address2] = txtAddress2.Text
                rs_branch_m![address3] = txtAddress3.Text
                rs_branch_m![companyshortname] = gs_comp_short_name
                rs_branch_m.Update
                gf_mdi_message "Updated"
                cmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If wf_valid_data = False Then Exit Sub
    Set rs_branch_m = New ADODB.Recordset
    rs_branch_m.Open "select * from branchmast where branchcode='" & Trim(txtBranchCode.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    If rs_branch_m.RecordCount > 0 Then
        MsgBox ("Cannot save!..Branch code already exists")
        txtBranchCode.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Do you wish to save?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
        Set rs_branch_m = New ADODB.Recordset
        rs_branch_m.Open "select * from branchmast", cn, adOpenKeyset, adLockOptimistic
        
        rs_branch_m.AddNew
        rs_branch_m![branchcode] = txtBranchCode.Text
        rs_branch_m![branchname] = txtBranchName.Text
        rs_branch_m![address1] = txtAddress1.Text
        rs_branch_m![address2] = txtAddress2.Text
        rs_branch_m![address3] = txtAddress3.Text
        rs_branch_m![companyshortname] = gs_comp_short_name
        rs_branch_m.Update
        gf_mdi_message "Saved"
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    gf_mdi_message "Ready"
End Sub

Sub wf_set_data()
    ls_branch_code = Mid(cboBranchCode.Text, InStr(1, cboBranchCode.Text, "-") + 1)
    Set rs_branch_m = New ADODB.Recordset
    rs_branch_m.Open "select * from branchmast where branchcode='" & ls_branch_code & "'", cn, adOpenKeyset, adLockOptimistic
    If rs_branch_m.RecordCount > 0 Then
        txtBranchCode.Text = ls_branch_code
        txtBranchName.Text = rs_branch_m![branchname]
        txtAddress1.Text = rs_branch_m![address1]
        txtAddress2.Text = rs_branch_m![address2]
        txtAddress3.Text = rs_branch_m![address3]
    End If
    cboBranchCode.Visible = False
End Sub

Sub wf_clear_data()
    txtBranchCode.Text = ""
    txtBranchName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddress3.Text = ""
End Sub

Private Sub txtBranchCode_KeyPress(KeyAscii As Integer)
    gf_accept_alpha_numeric KeyAscii
End Sub

Private Sub txtBranchName_KeyPress(KeyAscii As Integer)
    gf_accept_alpha_numeric KeyAscii
End Sub

Function wf_valid_data() As Boolean
    If Len(Trim(txtBranchCode.Text)) <= 0 Then
        MsgBox "Branch code should not be blank", vbInformation, "PAYSOFT"
        txtBranchCode.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If Len(Trim(txtBranchName.Text)) <= 0 Then
        MsgBox "Branch Name should not be blank", vbInformation, "PAYSOFT"
        txtBranchName.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    wf_valid_data = True
End Function
