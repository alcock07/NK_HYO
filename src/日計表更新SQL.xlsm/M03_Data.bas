Attribute VB_Name = "M03_Data"
Option Explicit

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD = "Password=ALCadmin!;"

'売上データ処理
'(担当者毎データ集計)
Sub Get_TAN_Data(ByVal strDate As String)

    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strNT  As String
    Dim strSQL As String
    Dim strX   As String
    Dim strB   As String
    
    strNT = "Initial Catalog=process_os;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = "SELECT * FROM W_NKC"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    
    '担当者ﾚｺｰﾄﾞ作成(月初に実績のない担当者もいるのでマスタからレコードを作成する）
    Set Cmd.ActiveConnection = cnA
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "        'SELECT TANCD, "
    strSQL = strSQL & "                TANNM, "
    strSQL = strSQL & "                TANBMNCD "
    strSQL = strSQL & "         FROM TANMTA"
    strSQL = strSQL & "         WHERE DATKB = ''1''"
    strSQL = strSQL & "         ') as TAN "
    strSQL = strSQL & "WHERE TANBMNCD"
    strSQL = strSQL & "      IN (SELECT BMNCD"
    strSQL = strSQL & "          FROM 部門"
    strSQL = strSQL & "          WHERE kbn_code IN('01','02','05')"
    strSQL = strSQL & "         )"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    rsA.MoveFirst
    Do Until rsA.EOF
        rsW.AddNew
        rsW.Fields("SMADT") = strDate
        rsW.Fields("TANCD") = Trim(rsA.Fields(0))            '担当者ｺｰﾄﾞ
        strX = Trim(rsA.Fields(1))
        rsW.Fields("TANNM") = strX                           '担当者名
        strB = Trim(rsA.Fields(2))
        rsW.Fields("TANCLAID") = Left(strB, 2)               '名称区分A
        rsW.Fields("TANCLBID") = Mid(strB, 3, 2)             '名称区分B
        rsW.Fields("TANCLCID") = Right(strB, 2)              '名称区分C
        rsW.Fields("TANBMNCD") = strB                        '部門ｺｰﾄﾞ
        rsW.Fields("URIKNM") = 0
        rsW.Fields("ARAKNM") = 0
        If strX = "X" Then
        Else
            rsW.Update
        End If
        rsA.MoveNext
    Loop
    rsW.Close
    
    '表示順、支店名更新（部門区分）
    strSQL = ""
    strSQL = strSQL & "UPDATE W_NKC"
    strSQL = strSQL & "            Set W_NKC.TANCLBNM = 支店.stn_name"
    strSQL = strSQL & "            FROM 支店"
    strSQL = strSQL & "            WHERE W_NKC.TANCLBID = 支店.stn_code"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '部門情報取得（名称ﾏｽﾀ検索）
    strSQL = ""
    strSQL = strSQL & "UPDATE W_NKC"
    strSQL = strSQL & "             SET TANCLANM = MST.CLSNM"
    strSQL = strSQL & "             FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "             'SELECT CLSKB,"
    strSQL = strSQL & "                     CLSID,"
    strSQL = strSQL & "                     CLSNM"
    strSQL = strSQL & "              FROM   CLSMTA"
    strSQL = strSQL & "              WHERE CLSKB =''3''"
    strSQL = strSQL & "             ') as MST"
    strSQL = strSQL & "             INNER JOIN W_NKC"
    strSQL = strSQL & "             ON MST.CLSID = TANCLAID"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
        
    strSQL = ""
    strSQL = strSQL & "UPDATE W_NKC"
    strSQL = strSQL & "             SET TANCLCNM = MST.BMNNM"
    strSQL = strSQL & "             FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "             'SELECT BMNCD,"
    strSQL = strSQL & "                     BMNNM"
    strSQL = strSQL & "              FROM   BMNMTA"
    strSQL = strSQL & "              WHERE  DATKB = ''1''"
    strSQL = strSQL & "             ') as MST"
    strSQL = strSQL & "             INNER JOIN W_NKC"
    strSQL = strSQL & "             ON MST.BMNCD = TANBMNCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '担当者集計データ読込み
'    strSQL = ""
'    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
'    strSQL = strSQL & "                         'SELECT TANCD"
'    strSQL = strSQL & "                                ,Sum(URIKN)"
'    strSQL = strSQL & "                                ,Sum(GNKKN)"
'    strSQL = strSQL & "                                ,Sum(ZKMUZEKN)"
'    strSQL = strSQL & "                          FROM UDNTRA"
'    strSQL = strSQL & "                          WHERE SMADT = ''" & strDate & "''"
'    strSQL = strSQL & "                          And  (DENKB=''2''"
'    strSQL = strSQL & "                          Or    DENKB=''3'')"
'    strSQL = strSQL & "                          And   LINNO < ''990''"
'    strSQL = strSQL & "                          GROUP BY TANCD')"
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                         'SELECT TANCD"
    strSQL = strSQL & "                                ,Sum(URIKN)"
    strSQL = strSQL & "                                ,Sum(GNKKN)"
    strSQL = strSQL & "                                ,Sum(ZKMUZEKN)"
    strSQL = strSQL & "                          FROM V_UDNTRA"
    strSQL = strSQL & "                          WHERE SMADT = ''" & strDate & "''"
    strSQL = strSQL & "                          And  (DENKB=''2''"
    strSQL = strSQL & "                          Or    DENKB=''3'')"
    strSQL = strSQL & "                          GROUP BY TANCD')"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If
    Do Until rsA.EOF
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM W_NKC"
        strSQL = strSQL & "         WHERE TANCD='" & rsA(0) & "'"
        rsW.Open strSQL, cnA, adOpenStatic, adLockPessimistic
        If rsW.EOF = False Then
            rsW.Fields("URIKNM") = rsA(1) - rsA(3)
            rsW.Fields("ARAKNM") = rsA(1) - rsA(2) - rsA(3)
            rsW.Fields("WDT") = Format(Now(), "yyyymmdd")
            rsW.Fields("WTM") = Format(Now(), "hhmmss")
            rsW.Update
        End If
        rsW.Close
        rsA.MoveNext
    Loop
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If

End Sub

'計画データ処理
Sub Get_Plan(ByVal strDate As String)

    Dim cnW    As ADODB.Connection
    Dim rsW    As ADODB.Recordset
    Dim rsP    As ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strNT  As String
    Dim strCD  As String '担当者コード
    Dim lngD   As Long   '月
    
    
    lngD = CLng(Mid(strDate, 5, 2))
    
    '日計作業テーブル(SQLServer)
    Set cnW = New ADODB.Connection
    Set rsW = New ADODB.Recordset
    Set rsP = New ADODB.Recordset
    strNT = "Initial Catalog=process_os;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    Set Cmd.ActiveConnection = cnW
    strSQL = "SELECT * FROM W_NKC"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    rsW.MoveFirst
    If rsW.EOF = False Then rsW.MoveFirst
    Do Until rsW.EOF
        strCD = rsW.Fields("TANCD")
        strSQL = ""
        strSQL = strSQL & "SELECT 担当者コード,"
        strSQL = strSQL & "       Sum(売上01),"
        strSQL = strSQL & "       Sum(売上02),"
        strSQL = strSQL & "       Sum(売上03),"
        strSQL = strSQL & "       Sum(売上04),"
        strSQL = strSQL & "       Sum(売上05),"
        strSQL = strSQL & "       Sum(売上06),"
        strSQL = strSQL & "       Sum(売上07),"
        strSQL = strSQL & "       Sum(売上08),"
        strSQL = strSQL & "       Sum(売上09),"
        strSQL = strSQL & "       Sum(売上10),"
        strSQL = strSQL & "       Sum(売上11),"
        strSQL = strSQL & "       Sum(売上12),"
        strSQL = strSQL & "       Sum(粗利01),"
        strSQL = strSQL & "       Sum(粗利02),"
        strSQL = strSQL & "       Sum(粗利03),"
        strSQL = strSQL & "       Sum(粗利04),"
        strSQL = strSQL & "       Sum(粗利05),"
        strSQL = strSQL & "       Sum(粗利06),"
        strSQL = strSQL & "       Sum(粗利07),"
        strSQL = strSQL & "       Sum(粗利08),"
        strSQL = strSQL & "       Sum(粗利09),"
        strSQL = strSQL & "       Sum(粗利10),"
        strSQL = strSQL & "       Sum(粗利11),"
        strSQL = strSQL & "       Sum(粗利12) "
        strSQL = strSQL & "       FROM 年度計画"
        strSQL = strSQL & "       GROUP BY 担当者コード"
        strSQL = strSQL & "       HAVING ((担当者コード)='" & strCD & "')"
        Cmd.CommandText = strSQL
        Set rsP = Cmd.Execute
        If rsP.EOF Then
            rsW.Fields("URINP") = 0
            rsW.Fields("ARANP") = 0
        Else
            rsW.Fields("URINP") = rsP.Fields(lngD) * 10000
            rsW.Fields("ARANP") = rsP.Fields(lngD + 12) * 10000
        End If
        rsW.Update
        rsP.Close
        strSQL = ""
        strSQL = strSQL & "SELECT 担当者コード,"
        strSQL = strSQL & "       Sum(売上01),"
        strSQL = strSQL & "       Sum(売上02),"
        strSQL = strSQL & "       Sum(売上03),"
        strSQL = strSQL & "       Sum(売上04),"
        strSQL = strSQL & "       Sum(売上05),"
        strSQL = strSQL & "       Sum(売上06),"
        strSQL = strSQL & "       Sum(売上07),"
        strSQL = strSQL & "       Sum(売上08),"
        strSQL = strSQL & "       Sum(売上09),"
        strSQL = strSQL & "       Sum(売上10),"
        strSQL = strSQL & "       Sum(売上11),"
        strSQL = strSQL & "       Sum(売上12),"
        strSQL = strSQL & "       Sum(粗利01),"
        strSQL = strSQL & "       Sum(粗利02),"
        strSQL = strSQL & "       Sum(粗利03),"
        strSQL = strSQL & "       Sum(粗利04),"
        strSQL = strSQL & "       Sum(粗利05),"
        strSQL = strSQL & "       Sum(粗利06),"
        strSQL = strSQL & "       Sum(粗利07),"
        strSQL = strSQL & "       Sum(粗利08),"
        strSQL = strSQL & "       Sum(粗利09),"
        strSQL = strSQL & "       Sum(粗利10),"
        strSQL = strSQL & "       Sum(粗利11),"
        strSQL = strSQL & "       Sum(粗利12) "
        strSQL = strSQL & "       FROM 修正計画"
        strSQL = strSQL & "       GROUP BY 担当者コード"
        strSQL = strSQL & "       HAVING ((担当者コード)='" & strCD & "')"
        Cmd.CommandText = strSQL
        Set rsP = Cmd.Execute
        If rsP.EOF Then
            rsW.Fields("URISP") = 0
            rsW.Fields("ARASP") = 0
        Else
            rsW.Fields("URISP") = rsP.Fields(lngD) * 10000
            rsW.Fields("ARASP") = rsP.Fields(lngD + 12) * 10000
        End If
        rsW.Update
        rsP.Close
        rsW.MoveNext
    Loop

Exit_DB:

    If Not rsP Is Nothing Then
        If rsP.State = adStateOpen Then rsP.Close
        Set rsP = Nothing
    End If
    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub

'受注残データ処理
Sub Get_JZAN(ByVal strDate As String)

    Dim cnA    As ADODB.Connection
    Dim rsA    As ADODB.Recordset
    Dim rsJ    As ADODB.Recordset
    Dim strNT  As String
    Dim strCD  As String
    
    
    Set cnA = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    Set rsJ = New ADODB.Recordset
    
    strNT = "Initial Catalog=process_os;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = "SELECT * FROM W_NKC"
    rsA.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    rsA.MoveFirst
    Do Until rsA.EOF
        '担当者ごとの受注残データ取得
        strCD = rsA.Fields("TANCD")
        strSQL = ""
        strSQL = strSQL & "SELECT Sum(zankn),"
        strSQL = strSQL & "       Sum(gnkkn)"
        strSQL = strSQL & "       FROM JUZTBZ_Hybrid"
        strSQL = strSQL & "       WHERE tancd ='" & strCD & "'"
        strSQL = strSQL & "         AND nokdt <='" & strDate & "'"
        rsJ.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsJ.EOF Then
            rsA.Fields("JUZAN") = 0
            rsA.Fields("JUZANB") = 0
        Else
            If IsNull(rsJ.Fields(0)) Then
                rsA.Fields("JUZAN") = 0
            Else
                rsA.Fields("JUZAN") = rsJ.Fields(0)
            End If
            If IsNull(rsJ.Fields(1)) Then
                rsA.Fields("JUZANB") = 0
            Else
                rsA.Fields("JUZANB") = rsJ.Fields(1)
            End If
        End If
        rsJ.Close
        rsA.Update
        rsA.MoveNext
    Loop
    
Exit_DB:
    
    If Not rsJ Is Nothing Then
        If rsJ.State = adStateOpen Then rsJ.Close
        Set rsJ = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub

'当日データ処理
'(売上ﾄﾗﾝから当日分の売上ﾃﾞｰﾀ取得)
Sub Get_URI(strDate As String, strTime As String)

    Dim cnW     As ADODB.Connection
    Dim rsA     As ADODB.Recordset
    Dim rsU     As ADODB.Recordset
    Dim strNT   As String
    Dim strCODE As String
    Dim lngKIN  As Long
    Dim lngARA  As Long
    
    Set cnW = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    Set rsU = New ADODB.Recordset
    
    strNT = "Initial Catalog=process_os;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_NKC"
    rsA.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields("TANCD")
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
        strSQL = strSQL & "             'SELECT UDNDT,"
        strSQL = strSQL & "                     TANCD,"
        strSQL = strSQL & "                     URIKN,"
        strSQL = strSQL & "                     GNKKN,"
        strSQL = strSQL & "                     ZKMUZEKN"
        strSQL = strSQL & "              FROM V_UDNTRA"
        strSQL = strSQL & "                     WHERE UDNDT = ''" & strDate & "''"
        strSQL = strSQL & "                     AND TANCD = ''" & strCODE & "''')"
        rsU.Open strSQL, cnW, adOpenStatic, adLockReadOnly
        lngKIN = 0
        lngARA = 0
        If rsU.EOF = False Then rsU.MoveFirst
        Do Until rsU.EOF
            lngKIN = lngKIN + rsU.Fields(2) - rsU.Fields(4)
            lngARA = lngARA + rsU.Fields(2) - rsU.Fields(3) - rsU.Fields(4)
            rsU.MoveNext
        Loop
        rsU.Close
        rsA.Fields("URIKND") = lngKIN
        rsA.Fields("ARAKND") = lngARA
        rsA.Fields("WDT") = strDate
        rsA.Fields("WTM") = strTime
        rsA.Update
        rsA.MoveNext
    Loop
    
Exit_DB:

    If Not rsU Is Nothing Then
        If rsU.State = adStateOpen Then rsU.Close
        Set rsU = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub

Sub DEL_Nothing()

    Dim cnW     As ADODB.Connection
    Dim rsA     As ADODB.Recordset
    Dim Cmd     As New ADODB.Command
    Dim strNT   As String
    Dim start_time As Double
    Dim end_time   As Double
        
    Set cnW = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    
    start_time = Timer
    strNT = "Initial Catalog=process_os;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    
    strSQL = ""
    strSQL = strSQL & "DELETE FROM W_NKC"
    strSQL = strSQL & "                          WHERE URIKNM = 0"
    strSQL = strSQL & "                          And   URIKND = 0"
    strSQL = strSQL & "                          And   JUZAN = 0"
    strSQL = strSQL & "                          And   URINP = 0"
    strSQL = strSQL & "                          And   URISP = 0"
    Set Cmd.ActiveConnection = cnW
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    end_time = Timer
    Debug.Print "DELETE W_NKC " & (end_time - start_time)
    
    start_time = Timer
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub

'仕入データ処理
'(仕入ﾄﾗﾝから仕入ﾃﾞｰﾀ取得)
Sub Get_SIRE(ByVal strDate As String)

    Dim cnW     As ADODB.Connection
    Dim rsA     As ADODB.Recordset
    Dim rsS     As ADODB.Recordset
    Dim Cmd     As New ADODB.Command
    Dim strNT   As String
    Dim strCODE As String
    Dim lngKIN  As Long
    Dim start_time As Double
    Dim end_time   As Double
        
    Set cnW = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    Set rsS = New ADODB.Recordset
    
    start_time = Timer
    strNT = "Initial Catalog=process_os;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_NKS"
    rsA.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                         'SELECT TANCD"
    strSQL = strSQL & "                                ,TANNM"
    strSQL = strSQL & "                          FROM TANMTA"
    strSQL = strSQL & "                          WHERE DATKB = ''1''"
    strSQL = strSQL & "                          And TANCD < ''00000700''"
    strSQL = strSQL & "                          And (TANCLCID > ''89''"
    strSQL = strSQL & "                            Or TANCLCID IN (''85'',''87'',''88''))"
    strSQL = strSQL & "                          ORDER BY TANCD"
    strSQL = strSQL & "                          ')"
    Set Cmd.ActiveConnection = cnW
    Cmd.CommandText = strSQL
    Set rsS = Cmd.Execute
    rsS.MoveFirst
    Do Until rsS.EOF
        rsA.AddNew
        rsA.Fields("SMADT") = strDate
        rsA.Fields("TANCD") = rsS(0)
        rsA.Fields("TANNM") = rsS(1)
        rsA.Update
        rsS.MoveNext
    Loop
    rsS.Close
    end_time = Timer
    Debug.Print "TANMTA " & (end_time - start_time)
    
    start_time = Timer
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields("TANCD")
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
        strSQL = strSQL & "                         'SELECT sum(SREKN)"
        strSQL = strSQL & "                                 FROM V_SDNTRA"
        strSQL = strSQL & "                                 WHERE SMADT = ''" & strDate & "''"
        strSQL = strSQL & "                                 AND TANCD = ''" & strCODE & "''"
        strSQL = strSQL & "                                 GROUP BY TANCD"
        strSQL = strSQL & "                        ')"
        Cmd.CommandText = strSQL
        Set rsS = Cmd.Execute
        If rsS.EOF = False Then
            rsS.MoveFirst
            rsA.Fields("SREKN") = rsS(0)
        End If
        
        rsS.Close
        rsA.Update
        rsA.MoveNext
    Loop
    
    end_time = Timer
    Debug.Print "SDNTRA " & (end_time - start_time)
    
Exit_DB:

    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub
