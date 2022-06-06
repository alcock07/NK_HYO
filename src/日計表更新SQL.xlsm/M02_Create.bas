Attribute VB_Name = "M02_Create"
Option Explicit

Sub CR_TBL_NKC()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_NKC]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_NKC]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_NKCテーブル作成
    strSQL = strSQL & "CREATE TABLE [dbo].[W_NKC]( "
    strSQL = strSQL & "    [SMADT]     [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [TANBMNCD]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [URIKNM]    [real]      NULL, "
    strSQL = strSQL & "    [ARAKNM]    [real]      NULL, "
    strSQL = strSQL & "    [URIKND]    [real]      NULL, "
    strSQL = strSQL & "    [ARAKND]    [real]      NULL, "
    strSQL = strSQL & "    [JUZAN]     [real]      NULL, "
    strSQL = strSQL & "    [JUZANB]    [real]      NULL, "
    strSQL = strSQL & "    [URINP]     [real]      NULL, "
    strSQL = strSQL & "    [ARANP]     [real]      NULL, "
    strSQL = strSQL & "    [URISP]     [real]      NULL, "
    strSQL = strSQL & "    [ARASP]     [real]      NULL, "
    strSQL = strSQL & "    [TANCLAID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLBID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLCID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLANM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [TANCLBNM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [TANCLCNM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [JUN]       [nchar](4)  NULL, "
    strSQL = strSQL & "    [WDT]       [nchar](8)  NULL, "
    strSQL = strSQL & "    [WTM]       [nchar](6)  NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_NKC] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NKC()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_NKC]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_NKC]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

'Sub CR_TBL_KBN()
'
'Const DBK  As String = "\\192.168.128.4\hb\sys\MASTER\部門区分.accdb"
'
'Dim cnG       As New ADODB.Connection
'Dim cnA       As New ADODB.Connection
'Dim rsG       As New ADODB.Recordset
'Dim rsA       As New ADODB.Recordset
'Dim strNT     As String
'Dim strSQL    As String
'Dim strKBN(5, 299) As String
'Dim lngC      As Long
'Dim lngR      As Long
'
'    strNT = "Initial Catalog=process_os;"
'    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
'    cnG.Open
'
'    'W_KBNテーブル削除
'    strSQL = ""
'    strSQL = strSQL & "if exists (select * from sysobjects where id = "
'    strSQL = strSQL & "object_id(N'[dbo].[W_KBN]') and "
'    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
'    strSQL = strSQL & "DROP TABLE [dbo].[W_KBN]"
'    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
'
'    'W_KBNテーブル作成
'    strSQL = ""
'    strSQL = strSQL & "CREATE TABLE [dbo].[W_KBN]( "
'    strSQL = strSQL & "    [TANCD]     [nchar](4)  NOT NULL,"
'    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
'    strSQL = strSQL & "    [TANC8]     [nchar](8)  NULL, "
'    strSQL = strSQL & "    [STN]       [nchar](8)  NULL, "
'    strSQL = strSQL & "    [JUN]       [nchar](4)  NULL, "
'    strSQL = strSQL & "    [KBN]       [nchar](1)  NULL, "
'    strSQL = strSQL & "CONSTRAINT [PK_KBN] PRIMARY KEY CLUSTERED "
'    strSQL = strSQL & "( "
'    strSQL = strSQL & "[TANCD] ASC "
'    strSQL = strSQL & ") WITH "
'    strSQL = strSQL & "(PAD_INDEX = OFF, "
'    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
'    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
'    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
'    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
'    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
'    strSQL = strSQL & ") ON [PRIMARY]"
'    strSQL = strSQL & ") ON [PRIMARY]"
'    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
'
'     '部門区分読込み----------------------------------
'    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBK
'    cnA.Open
'    strSQL = ""
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & "       担当者ｺｰﾄﾞ,"
'    strSQL = strSQL & "       担当者名,"
'    strSQL = strSQL & "       担当者ｺｰﾄﾞ8,"
'    strSQL = strSQL & "       支店,"
'    strSQL = strSQL & "       順ｺｰﾄﾞ,"
'    strSQL = strSQL & "       区分"
'    strSQL = strSQL & "       FROM 部門区分"
'    rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
'    rsA.MoveFirst
'    lngR = 0
'    Do Until rsA.EOF
'        For lngC = 0 To 5
'            strKBN(lngC, lngR) = rsA.Fields(lngC)
'        Next lngC
'        lngR = lngR + 1
'        rsA.MoveNext
'    Loop
'    rsA.Close
'
'
'    strSQL = ""
'    strSQL = strSQL & "    SELECT * FROM  W_KBN"
'    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
'    lngR = 0
'    Do
'        If strKBN(1, lngR) = "" Then Exit Do
'        rsG.AddNew
'        For lngC = 0 To 5
'            rsG.Fields(lngC) = strKBN(lngC, lngR)
'        Next lngC
'        rsG.Update
'        lngR = lngR + 1
'    Loop
'
'    If Not rsG Is Nothing Then
'        If rsG.State = adStateOpen Then rsG.Close
'        Set rsG = Nothing
'    End If
'    If Not cnG Is Nothing Then
'        If cnG.State = adStateOpen Then cnG.Close
'        Set cnG = Nothing
'    End If
'    If Not rsA Is Nothing Then
'        If rsA.State = adStateOpen Then rsA.Close
'        Set rsA = Nothing
'    End If
'    If Not cnA Is Nothing Then
'        If cnA.State = adStateOpen Then cnA.Close
'        Set cnA = Nothing
'    End If
'
'End Sub
'
'Sub DR_TBL_KBN()
'
'Dim cnG       As New ADODB.Connection
'Dim rsG       As New ADODB.Recordset
'Dim strNT     As String
'Dim strSQL    As String
'
'    strNT = "Initial Catalog=process_os;"
'    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
'    cnG.Open
'
'    'W_KBNテーブル削除
'    strSQL = ""
'    strSQL = strSQL & "if exists (select * from sysobjects where id = "
'    strSQL = strSQL & "object_id(N'[dbo].[W_KBN]') and "
'    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
'    strSQL = strSQL & "DROP TABLE [dbo].[W_KBN]"
'    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
'
'    If Not rsG Is Nothing Then
'        If rsG.State = adStateOpen Then rsG.Close
'        Set rsG = Nothing
'    End If
'    If Not cnG Is Nothing Then
'        If cnG.State = adStateOpen Then cnG.Close
'        Set cnG = Nothing
'    End If
'End Sub

Sub CR_TBL_NKS()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_NKS]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_NKS]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_NKCテーブル作成
    strSQL = strSQL & "CREATE TABLE [dbo].[W_NKS]( "
    strSQL = strSQL & "    [SMADT]     [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [SREKN]     [real]      NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_NKS] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NKS()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_NKS]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_NKS]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

'日計データ累積（売上）
Sub CR_TBL_NKCR()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_NKCテーブル作成
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_URI]( "
    strSQL = strSQL & "    [SMADT]     [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [TANBMNCD]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [URIKNM]    [real]      NULL, "
    strSQL = strSQL & "    [ARAKNM]    [real]      NULL, "
    strSQL = strSQL & "    [URIKND]    [real]      NULL, "
    strSQL = strSQL & "    [ARAKND]    [real]      NULL, "
    strSQL = strSQL & "    [JUZAN]     [real]      NULL, "
    strSQL = strSQL & "    [JUZANB]    [real]      NULL, "
    strSQL = strSQL & "    [URINP]     [real]      NULL, "
    strSQL = strSQL & "    [ARANP]     [real]      NULL, "
    strSQL = strSQL & "    [URISP]     [real]      NULL, "
    strSQL = strSQL & "    [ARASP]     [real]      NULL, "
    strSQL = strSQL & "    [TANCLAID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLBID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLCID]  [nchar](6)  NULL, "
    strSQL = strSQL & "    [TANCLANM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [TANCLBNM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [TANCLCNM]  [nchar](20) NULL, "
    strSQL = strSQL & "    [JUN]       [nchar](4)  NULL, "
    strSQL = strSQL & "    [WDT]       [nchar](8)  NULL, "
    strSQL = strSQL & "    [WTM]       [nchar](6)  NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_NKURI] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

'日計データ累積（仕入）
Sub CR_TBL_NKSR()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_NKCテーブル削除
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_NKCテーブル作成
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_SRE]( "
    strSQL = strSQL & "    [SMADT]     [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [SREKN]     [real]      NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_NKSRE] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

