Attribute VB_Name = "gf_database_connection"
Option Explicit
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal parent As Integer, _
ByVal request As Integer, ByVal driver As String, ByVal attributes As String) As Integer
Private Const ODBC_ADD_SYS_DSN = 1
Private Const ODBC_REMOVE_DSN = 3

Public Sub gf_valid_login()
    If gs_login_valid = True Then
        Call gf_db_connection
'        mdiMain.mnuReport.Enabled = True
'        mdiMain.mnuWorker.Enabled = True
'        mdiMain.mnubranch.Enabled = True
'        mdiMain.mnuGrdFes.Enabled = True
'        mdiMain.mnuTrans.Enabled = True
    End If
    mdiMain.Show
End Sub

Public Function gf_CreateDSN(strDSN As String, strDriver As String, strDBName As String)
Dim sAttributes As String
Dim retVal As Integer

sAttributes = ("DSN=" & strDSN)
sAttributes = sAttributes & Chr(0)
sAttributes = sAttributes & "DBQ="
sAttributes = sAttributes & strDBName
sAttributes = sAttributes & Chr(0)
sAttributes = sAttributes & Chr(0)

retVal = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, strDriver, sAttributes)

End Function

Public Sub gf_db_connection()
    Dim lstr_conn As String
    
    Set cn = New ADODB.Connection
    lstr_conn = "DSN=Paysoft;uid=;pwd=;"
    cn.Open lstr_conn
End Sub
