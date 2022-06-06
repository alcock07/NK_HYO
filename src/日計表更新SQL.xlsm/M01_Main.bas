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

'=== ���v�f�[�^�W�v���� ===
Sub Proc_TZ()

    Dim strDateC  As String  '������
    Dim strDateZ  As String  '�O����
    Dim lngMM     As Long    '���t�Z�o��Ɨp
    Dim lngYY     As Long    '���t�Z�o��Ɨp
    Dim DateA     As Date    '���t�Z�o��Ɨp

    '���������O�����Z�o_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
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
    '�@��ƃe�[�u���쐬
    '�A����f�[�^�쐬
    '�B�c�ƌv�悩��v��擾���Ĕ���f�[�^�ɓ����
    '�C�󒍎c�ް��擾���Ĕ���f�[�^�ɓ����
    '�D������݂��瓖������擾���Ĕ���f�[�^�ɓ����
    '�E�d���f�[�^�쐬
    '�F��Ɨp����ݐϗp�փf�[�^������
    
    Dim start_time As Double
    Dim end_time As Double
    
    Sheets("Wait").Range("D15") = "�������E�E�E"
    Sheets("Wait").Range("D16") = ""
    DoEvents
    
    '�@��ƃe�[�u���쐬_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    start_time = Timer
    Call CR_TBL_NKC '��ƃe�[�u���쐬
    end_time = Timer
    Debug.Print "CR_TBL_NKC " & (end_time - start_time)

'2022.3.8 SQL�ڍs�̂��ߕ���敪�s�v
'    start_time = Timer
'    Call CR_TBL_KBN '����敪�쐬
'    end_time = Timer
'    Debug.Print "CR_TBL_KBN " & (end_time - start_time)
    
'    �A����f�[�^�쐬_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "����f�[�^�擾���E�E�E"
    Sheets("Wait").Range("D16") = ""
    DoEvents
    
    start_time = Timer
    Call Get_TAN_Data(strDt)
    end_time = Timer
    Debug.Print "Get_TAN_Data " & (end_time - start_time)

    '�B�c�ƌv�悩��v��擾_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "�v��f�[�^�擾���E�E�E"
    DoEvents
    start_time = Timer
    Call Get_Plan(strDt)
    end_time = Timer
    Debug.Print "Get_Plan " & (end_time - start_time)

    '�C�󒍎c�ް��擾_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "�󒍃f�[�^�擾���E�E�E"
    DoEvents
    start_time = Timer
    Call Get_JZAN(strDt)
    end_time = Timer
    Debug.Print "Get_JZAN " & (end_time - start_time)
    
    '�D������݂��瓖������擾_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

    Sheets("Wait").Range("D15") = "�����f�[�^�擾���E�E�E"
    DoEvents
    start_time = Timer
    If strCZ = "C" Then
        Call Get_URI(Format(Now(), "yyyymmdd"), Format(Now(), "hhnnss"))
    Else
        Call Get_URI(strDt, "000000")
    End If
    end_time = Timer
    Debug.Print "Get_URI " & (end_time - start_time)
    
    '����A�󒍁A�v�悪�S��0��ں��ނ��폜
    Call DEL_Nothing
    
    '�E�d���ް��擾_/_/_/_/_/_/_/_/_/_/_/_/_/_/_//_/_/_/_/_/_/
    
    Sheets("Wait").Range("D15") = "�d���f�[�^�擾���E�E�E"
    DoEvents
    start_time = Timer
    Call CR_TBL_NKS '��ƃe�[�u���쐬
    Call Get_SIRE(strDt)
    end_time = Timer
    Debug.Print "Get_SIRE " & (end_time - start_time)

    '�F�f�[�^��z�M�pDB��_/_/_/_/_/_/_/_/_/_/_/_/_/_/_//_/_/_/
    
    Sheets("Wait").Range("D15") = "�I���������E�E�E"
    DoEvents
    start_time = Timer
    Call Set_R(strDt)
    end_time = Timer
    Debug.Print "Set_R " & (end_time - start_time)
    
    Sheets("Wait").Range("D15") = "�X�V����"
    DoEvents
    start_time = Timer
    Call DR_TBL_NKC '��ƃe�[�u���폜
    Call DR_TBL_NKS '��ƃe�[�u���폜
    Call DR_TBL_KBN '����敪�폜
    end_time = Timer
    Debug.Print "DR_TBL " & (end_time - start_time)
    
End Sub

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' �R���s���[�^�[���̒�����ݒ�
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' �R���s���[�^�[�����擾
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' �R���s���[�^�[����\��
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)

End Function

Sub AP_END()
   
    Dim myBook As Workbook
    Dim strFN  As String
    Dim boolB  As Boolean
    
    'Excell���ɂ��̃u�b�N�ȊO�̃u�b�N���L���Excell���I�����Ȃ�
    ThisWorkbook.Save

    strFN = ThisWorkbook.Name '���̃u�b�N�̖��O
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  '�t�@�C�������
    Else
        Application.Quit  'Excell���I��
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
