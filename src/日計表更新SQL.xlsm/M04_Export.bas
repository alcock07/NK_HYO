Attribute VB_Name = "M04_Export"
Option Explicit

'作業ﾃｰﾌﾞﾙから累計ﾃｰﾌﾞﾙへｺﾋﾟｰ
Sub Set_R(strDt As String)

Dim cnW    As ADODB.Connection
Dim rsU    As ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strNT  As String
Dim strSQL As String


    Set cnW = New ADODB.Connection
    Set rsU = New ADODB.Recordset
    strNT = "Initial Catalog=process_os;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    Set Cmd.ActiveConnection = cnW
    '今回更新月のデータを削除
    strSQL = ""
    strSQL = strSQL & "DELETE FROM NK_URI"
    strSQL = strSQL & "            WHERE SMADT = " & strDt
    Cmd.CommandText = strSQL
    Set rsU = Cmd.Execute
    '作業テーブルのデータをコピー
    strSQL = ""
    strSQL = strSQL & "INSERT INTO NK_URI"
    strSQL = strSQL & "       SELECT * FROM W_NKC"
    Cmd.CommandText = strSQL
    Set rsU = Cmd.Execute
    '今回更新月のデータを削除
    strSQL = ""
    strSQL = strSQL & "DELETE FROM NK_SRE"
    strSQL = strSQL & "            WHERE SMADT = " & strDt
    Cmd.CommandText = strSQL
    Set rsU = Cmd.Execute
    '作業テーブルのデータをコピー
    strSQL = ""
    strSQL = strSQL & "INSERT INTO NK_SRE"
    strSQL = strSQL & "       SELECT * FROM W_NKS"
    Cmd.CommandText = strSQL
    Set rsU = Cmd.Execute
    
Exit_DB:

    If Not rsU Is Nothing Then
        If rsU.State = adStateOpen Then rsU.Close
        Set rsU = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub
