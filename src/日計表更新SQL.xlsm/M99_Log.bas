Attribute VB_Name = "M99_Log"
Option Explicit

Public Const APP_NAME = "日計表更新SQL"

Sub Open_Log(strCPN As String)
Dim strLOG As String
Dim boolA  As Boolean
    strLOG = Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " -" & strCPN & "- " & APP_NAME & "：Start"
    boolA = AddText("X:\admin\alcock.Log", strLOG)
End Sub

Sub Close_Log(strCPN As String)
Dim strLOG As String
Dim boolA As Boolean
    strLOG = Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " -" & strCPN & "- " & APP_NAME & "：End"
    boolA = AddText("X:\admin\alcock.Log", strLOG)
End Sub

Public Function AddText(FName As String, txt As String) As Boolean
'=============================
'ﾃｷｽﾄﾌｧｲﾙ追加
'FName : 出力ﾌｧｲﾙ名
'txt   : 出力ﾃｷｽﾄ
'=============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open FName For Append As iFNW
        Print #iFNW, txt
    Close iFNW
End Function

Public Function WriteText(FName As String, txt As String) As Boolean
'==============================
'ﾃｷｽﾄﾌｧｲﾙ書込み
'FName : 出力ﾌｧｲﾙ名
'txt   : 出力ﾃｷｽﾄ
'==============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open FName For Output As iFNW
        Print #iFNW, txt
    Close iFNW
End Function

