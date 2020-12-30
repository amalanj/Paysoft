VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll and Labour Software for Hotel Industry"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14655
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   19052
            MinWidth        =   19052
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
            Object.ToolTipText     =   "Branch"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "Master"
      Begin VB.Menu mnuBranchm 
         Caption         =   "Branch Master"
      End
      Begin VB.Menu mnuGradem 
         Caption         =   "Grade Master"
      End
      Begin VB.Menu mnuFestivalm 
         Caption         =   "Festival Master"
      End
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    stBar.Panels.Item(3).Text = Format(Date, "dd-mmm-yyyy")
    gs_app_path = App.Path
    gs_db_path = Mid(gs_app_path, 1, InStr(1, UCase(gs_app_path), "VBP") - 1)
    gf_CreateDSN "Paysoft", "Microsoft Access Driver (*.mdb)", gs_db_path & "Reports\paysoftdb.mdb"
    gf_mdi_message "Creating Data Source"
    Call gf_db_connection
    gf_mdi_message "Database Connected"
End Sub

Private Sub mnuBranchm_Click()
    gf_mdi_message "Opening Branch Master"
    frmBranchMaster.Show
End Sub

Private Sub mnuFestivalm_Click()
    gf_mdi_message "Opening Festival Master"
    frmFestivalMaster.Show
End Sub

Private Sub mnuGradem_Click()
    gf_mdi_message "Opening Grade Master"
    frmGradeMaster.Show
End Sub

Private Sub mnuLogout_Click()
    End
End Sub
