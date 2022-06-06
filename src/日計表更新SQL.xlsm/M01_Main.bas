Attribute VB_Name = "M01_Main"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public Const MAX_COMPUTERNAME_LENGTH = 15

Public strSQL    As String

'=== 日計データ集計処理 ===
Sub Proc_TZ()

    Dim strDateC  As String  '当月末
    Dim strDateZ  As String  '前月末
    Dim lngMM     As Long    '日付算出作業用
    Dim lngYY     As Long    '日付算出作業用
    Dim DateA     As Date    '日付算出作業用

    '当月末＆前月末算出_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    lngMM = CLng(Format(Now(), "m"))
    lngYY = CLng(Format(Now(), "yyyy"))
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngYY = lngYY + 1
        lngMM = 1
    End If
    DateA = CStr(lngYY) & "/" & CStr(lngMM) & "/01"
    strDateC = Format(DateA - 1, "yyyymmdd")
    Call Proc_NKC(strDateC, "C")
    
    DateA = Format(Now(), "yyyy/mm/") & "01"
    strDateZ = Format(DateA - 1, "yyyymmdd")
    Call Proc_NKC(strDateZ, "Z")
    
End Sub

Sub Proc_NKC(strDt As String, strCZ As String)

    '===================================================
    '�@作業テーブル作成
    '�A売上データ作成
    '�B営業計画から計画取得して売上データに入れる
    '�C受注残ﾃﾞｰﾀ取得して売上データに入れる
    '�D売上ﾄﾗﾝから当日売り取得して売上データに入れる
    '�E仕入データ作成
    '�F作業用から累積用へデータを入れる
    
    Dim start_time As Double
    Dim end_time As Double
    
    Sheets("Wait").Range("D15") = "準備中・・・"
    Sheets("Wait").Range("D16") = ""
    DoEvents
    
    '�@作業テーブル作成_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    start_time = Timer
    Call CR_TBL_NKC '作業テーブル作成
    end_time = Timer
    Debug.Print "CR_TBL_NKC " & (end_time - start_time)

'2022.3.8 SQL移行のため部門区分不要
'    start_time = Timer
'    Call CR_TBL_KBN '部門区分作成
'    end_time = Timer
'    Debug.Print "CR_TBL_KBN " & (end_time - start_time)
    
'    �A売上データ作成_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "売上データ取得中・・・"
    Sheets("Wait").Range("D16") = ""
    DoEvents
    
    start_time = Timer
    Call Get_TAN_Data(strDt)
    end_time = Timer
    Debug.Print "Get_TAN_Data " & (end_time - start_time)

    '�B営業計画から計画取得_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "計画データ取得中・・・"
    DoEvents
    start_time = Timer
    Call Get_Plan(strDt)
    end_time = Timer
    Debug.Print "Get_Plan " & (end_time - start_time)

    '�C受注残ﾃﾞｰﾀ取得_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "受注データ取得中・・・"
    DoEvents
    start_time = Timer
    Call Get_JZAN(strDt)
    end_time = Timer
    Debug.Print "Get_JZAN " & (end_time - start_time)
    
    '�D売上ﾄﾗﾝから当日売り取得_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "当日データ取得中・・・"
    DoEvents
    start_time = Timer
    If strCZ = "C" Then
        Call Get_URI(Format(Now(), "yyyymmdd"), Format(Now(), "hhnnss"))
    Else
        Call Get_URI(strDt, "000000")
    End If
    end_time = Timer
    Debug.Print "Get_URI " & (end_time - start_time)
    
    '売上、受注、計画が全て0のﾚｺｰﾄﾞを削除
    Call DEL_Nothing
    
    '�E仕入ﾃﾞｰﾀ取得_/_/_/_/_/_/_/_/_/_/_/_/_/_/_//_/_/_/_/_/_/
    
    Sheets("Wait").Range("D15") = "仕入データ取得中・・・"
    DoEvents
    start_time = Timer
    Call CR_TBL_NKS '作業テーブル作成
    Call Get_SIRE(strDt)
    end_time = Timer
    Debug.Print "Get_SIRE " & (end_time - start_time)

    '�Fデータを配信用DBへ_/_/_/_/_/_/_/_/_/_/_/_/_/_/_//_/_/_/
    
    Sheets("Wait").Range("D15") = "終了処理中・・・"
    DoEvents
    start_time = Timer
    Call Set_R(strDt)
    end_time = Timer
    Debug.Print "Set_R " & (end_time - start_time)
    
    Sheets("Wait").Range("D15") = "更新完了"
    DoEvents
    start_time = Timer
    Call DR_TBL_NKC '作業テーブル削除
    Call DR_TBL_NKS '作業テーブル削除
    Call DR_TBL_KBN '部門区分削除
    end_time = Timer
    Debug.Print "DR_TBL " & (end_time - start_time)
    
End Sub

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' コンピューター名の長さを設定
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' コンピューター名を取得
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' コンピューター名を表示
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)

End Function

Sub AP_END()
   
    Dim myBook As Workbook
    Dim strFN  As String
    Dim boolB  As Boolean
    
    'Excell内にこのブック以外のブックが有ればExcellを終了しない
    ThisWorkbook.Save

    strFN = ThisWorkbook.Name 'このブックの名前
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  'ファイルを閉じる
    Else
        Application.Quit  'Excellを終了
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
