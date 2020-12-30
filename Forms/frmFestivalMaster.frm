VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFestivalMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Festival Master Maintenance"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
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
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboFestivalName 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txtFestivalName 
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
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   0
      Top             =   480
      Width           =   4215
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
      Left            =   600
      TabIndex        =   3
      Top             =   1920
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1920
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
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
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
      Left            =   4920
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskFestivalDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/MM/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Festival Name"
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
      TabIndex        =   9
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Festival Date"
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
      TabIndex        =   8
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   240
      Top             =   240
      Width           =   6135
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   570
      Top             =   1890
      Width           =   5475
   End
End
Attribute VB_Name = "frmFestivalMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_festival_m As ADODB.Recordset

Private Sub cboFestivalName_Click()
    If Len(Trim(cboFestivalName.Text)) > 0 Then
        cboSno.ListIndex = cboFestivalName.ListIndex
        wf_set_data
        If cmdEdit.Caption = "&Update" Then
            txtFestivalName.SetFocus
        End If
        If cmdDelete.Caption = "Confir&m" Then
            txtFestivalName.Locked = True
            mskFestivalDate.Enabled = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    wf_clear_data
    txtFestivalName.Locked = False
    mskFestivalDate.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Caption = "&Save"
    cmdEdit.Caption = "&Edit"
    cmdDelete.Caption = "&Delete"
    cboFestivalName.Visible = False
    cboSno.Visible = False
    gf_mdi_message "Ready"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmdDelete.Caption = "&Delete" Then
        wf_clear_data
        cboFestivalName.Visible = True
        cboFestivalName.Clear
        cboSno.Clear
        cmdEdit.Enabled = False
        cmdSave.Enabled = False
        cmdDelete.Caption = "Confir&m"
        Set rs_festival_m = New ADODB.Recordset
        rs_festival_m.Open "select * from festivalmast order by sno", cn, adOpenKeyset, adLockOptimistic
        If rs_festival_m.RecordCount > 0 Then
            rs_festival_m.MoveFirst
            Do While Not rs_festival_m.EOF
                cboFestivalName.AddItem rs_festival_m![festivalname] & "-" & rs_festival_m![festivaldate]
                cboSno.AddItem rs_festival_m![sno]
                rs_festival_m.MoveNext
            Loop
        End If
    ElseIf cmdDelete.Caption = "Confir&m" And Len(Trim(txtFestivalName.Text)) > 0 Then
        If MsgBox("Do you want to delete?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            gf_execute_sql ("delete from festivalmast where sno=" & Val(cboSno.Text))
            gf_mdi_message "Deleted"
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        wf_clear_data
        cboFestivalName.Visible = True
        cboFestivalName.Clear
        cboSno.Clear
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        cmdEdit.Caption = "&Update"
        Set rs_festival_m = New ADODB.Recordset
        rs_festival_m.Open "select * from festivalmast order by sno", cn, adOpenKeyset, adLockOptimistic
        If rs_festival_m.RecordCount > 0 Then
            rs_festival_m.MoveFirst
            Do While Not rs_festival_m.EOF
                cboFestivalName.AddItem rs_festival_m![festivalname] & "-" & rs_festival_m![festivaldate]
                cboSno.AddItem rs_festival_m![sno]
                rs_festival_m.MoveNext
            Loop
        End If
        cboFestivalName.SetFocus
    ElseIf cmdEdit.Caption = "&Update" And Len(Trim(txtFestivalName.Text)) > 0 Then
        If MsgBox("Do you wish to update?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
            If wf_valid_data = False Then Exit Sub
            
            Set rs_festival_m = New ADODB.Recordset
            rs_festival_m.Open "select * from festivalmast where sno=" & Val(cboSno.Text), cn, adOpenKeyset, adLockOptimistic
            If rs_festival_m.RecordCount > 0 Then
                rs_festival_m![festivalname] = txtFestivalName.Text
                rs_festival_m![festivaldate] = mskFestivalDate.Text
                rs_festival_m.Update
                gf_mdi_message "Updated"
                cmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If wf_valid_data = False Then Exit Sub
    
    Set rs_festival_m = New ADODB.Recordset
    rs_festival_m.Open "select * from festivalmast where festivalname='" & Trim(txtFestivalName.Text) & "'", cn, adOpenKeyset, adLockOptimistic
    If rs_festival_m.RecordCount > 0 Then
        MsgBox ("Cannot save!..Festival Name already exists")
        txtFestivalName.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Do you wish to save?", vbYesNo + vbQuestion, "PAYSOFT") = vbYes Then
        Set rs_festival_m = New ADODB.Recordset
        rs_festival_m.Open "select * from festivalmast", cn, adOpenKeyset, adLockOptimistic
        
        rs_festival_m.AddNew
        rs_festival_m![festivalname] = txtFestivalName.Text
        rs_festival_m![festivaldate] = mskFestivalDate.Text
        rs_festival_m.Update
        gf_mdi_message "Saved"
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    gf_mdi_message "Ready"
End Sub

Private Sub txtFestivalName_KeyPress(KeyAscii As Integer)
    gf_accept_alpha_numeric KeyAscii
End Sub

Sub wf_clear_data()
    txtFestivalName.Text = ""
    mskFestivalDate.Text = "  /  /    "
End Sub

Sub wf_set_data()
    Set rs_festival_m = New ADODB.Recordset
    rs_festival_m.Open "select * from festivalmast where sno=" & cboSno.Text, cn, adOpenKeyset, adLockOptimistic
    If rs_festival_m.RecordCount > 0 Then
        txtFestivalName.Text = rs_festival_m![festivalname]
        mskFestivalDate.Text = Format(rs_festival_m![festivaldate], "dd/MM/yyyy")
    End If
    cboFestivalName.Visible = False
End Sub

Function wf_valid_data() As Boolean
    If Len(Trim(txtFestivalName.Text)) <= 0 Then
        MsgBox "Festival Name should not be blank", vbInformation, "PAYSOFT"
        txtFestivalName.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If mskFestivalDate.Text = "__/__/____" Then
        MsgBox "Festival Date should not be blank", vbInformation, "PAYSOFT"
        mskFestivalDate.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    
    If IsDate(mskFestivalDate.Text) = False Then
        MsgBox "Invalid date", vbInformation, "PAYSOFT"
        mskFestivalDate.SetFocus
        wf_valid_data = False
        Exit Function
    End If
    wf_valid_data = True
End Function
