Attribute VB_Name = "gf_global_functions"
Public Function gf_valid_data(ls_sql As String) As Long

Set gadodc_valid_recordset = New ADODB.Recordset
gadodc_valid_recordset.Open ls_sql, cn, adOpenKeyset, adLockOptimistic

gf_valid_data = gadodc_valid_recordset.RecordCount
End Function

Public Function gf_execute_sql(ls_sql As String)

cn.Execute CStr(ls_sql)

End Function

Public Sub CenterForm(X As Form, Optional vParent, Optional vShowMode)

    Dim oParent As Object
    Dim iMode%
    
    Screen.MousePointer = vbHourglass
    If IsMissing(vParent) Then
        Set oParent = Screen
    Else
        Set oParent = vParent
    End If
        
    If IsMissing(vShowMode) Then iMode = vbModeless Else _
    iMode = Abs(vShowMode) Mod 2
    X.Move (oParent.Width \ 2 - X.Width \ 2), ((oParent.Height \ 2) * 0.85 - X.Height \ 2)
    Load X
    X.Show iMode
    Screen.MousePointer = vbNormal

End Sub

Public Sub gf_mdi_message(p_str_msg As String)
    mdiMain.stBar.Panels.Item(1).Text = p_str_msg
End Sub

Public Function gf_accept_only_text(p_key As Integer) As Integer
    If p_key >= 48 And p_key <= 57 Then
        p_key = 0
    Else
        p_key = Asc(UCase(Chr(p_key)))
    End If
    gf_accept_only_text = p_key
End Function

Public Function gf_accept_only_number(p_key As Integer)
    If Not (p_key >= 48 And p_key <= 57 Or p_key = 46 Or p_key = 8) Then
        p_key = 0
    End If
    gf_accept_only_number = p_key
End Function

Public Function gf_accept_alpha_numeric(p_key As Integer) As Integer
    If p_key >= 97 And p_key <= 122 Then
        p_key = p_key - 32
    End If
    If Not ((p_key >= 65 And p_key <= 90) Or p_key = 8 Or p_key = 32 Or p_key = 46 Or (p_key >= 48 And p_key <= 57) Or p_key = 44) Then
        p_key = 1
    End If
    gf_accept_only_alphanumeric = p_key
End Function
