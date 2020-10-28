Attribute VB_Name = "kccWsFuncRegExp"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccWsFuncRegExp
Rem
Rem  @description   Excel���[�N�V�[�g�p ���K�\�� ����UDF�W
Rem
Rem  @update        2020/10/28
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft VBScript Regular Expressions 5.5
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2020/09/19 ����ō쐬
Rem    2020/10/28 ���C�Z���X�E�h�L�������g�ǉ�
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem    |-----------------------------------------------------------------------------------------------------|
Rem    | �֐���            ����                                   �߂�l                                     |
Rem    |-----------------------------------------------------------------------------------------------------|
Rem    | RegexIsMatch      �}�b�`���邩���m�F                     True/False                                 |
Rem    | RegexReplace      �}�b�`�����������u��                 �u���㕶����                               |
Rem    | RegexMatches      �}�b�`�����C�ӂ̃v���p�e�B���         �v���p�e�B�ɂ���ĈقȂ�                   |
Rem    | RegexMatchCount   �}�b�`�����ӏ��̌�                   Variant/Long                               |
Rem    | RegexMatchIndexs  �}�b�`�����ӏ��̊J�n�C���f�b�N�X�z��   Variant/Long()                             |
Rem    | RegexMatchLengths �}�b�`�����ӏ��̕����񒷔z��           Variant/Long()                             |
Rem    | RegexMatchValues  �}�b�`�����ӏ��̒l�z��                 Variant/Variant()                          |
Rem    | RegexSubMatches   �}�b�`�����ӏ��̔z��̃T�u�}�b�`�l�z�� Variant/Variant(1 to N)(Variant(1 to M))   |
Rem    |-----------------------------------------------------------------------------------------------------|
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Sub Test_ALLTEST()
    Call Test_RegexIsMatch
    Call Test_RegexReplace
    Call Test_RegexMatches
    Call Test_RegexSubMatches
End Sub

Rem �}�b�`���邩���m�F
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem
Rem  @return As Boolean     True:�}�b�`�����BFalse:�}�b�`���Ȃ�����
Rem
Function RegexIsMatch(strSource As String, strPattern As String) As Boolean
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''�����p�^�[����ݒ�
        .IgnoreCase = False         ''�啶���Ə���������ʂ���
        .Global = True              ''������S�̂�����
        RegexIsMatch = re.Test(strSource)
    End With
End Function

Sub Test_RegexIsMatch()
    Debug.Print "-----Test_RegexIsMatch-----"
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexIsMatch(src, "abc")
    Debug.Print RegexIsMatch(src, "dgh")
    Debug.Print
End Sub

Rem �}�b�`�����������u��
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem  @param strReplace      �u��������
Rem
Rem  @return As String      �u����̕�����
Rem
Function RegexReplace(strSource As String, strPattern As String, strReplace As String) As String
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern
        .IgnoreCase = False
        .Global = True
        RegexReplace = re.Replace(strSource, strReplace)
    End With
End Function

Sub Test_RegexReplace()
    Debug.Print "-----Test_RegexReplace-----"
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexReplace(src, "abc", "XXX")
    Debug.Print RegexReplace(src, "xyz", "XXX")
    Debug.Print
End Sub

Rem �}�b�`�����C�ӂ̃v���p�e�B���
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem  @param strProperty     �擾�������v���p�e�B
Rem                         ���w��, Count, FirstIndex, Length, Value, SubMatches
Rem
Rem  @return As Variant     �v���p�e�B�ɂ���ĈقȂ�
Rem                         ���w��   VBScript_RegExp_55.MatchCollection
Rem                         Count    �}�b�`��������
Rem                         ����ȊO (1 To N)�̔z��
Rem
Function RegexMatches(strSource As String, strPattern As String, strProperty As String) As Variant
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern
        .IgnoreCase = False
        .Global = True
        
        Dim mc As VBScript_RegExp_55.MatchCollection
        Set mc = re.Execute(strSource)
        If strProperty = "" Then Set RegexMatches = mc: Exit Function
        If strProperty = "Count" Then RegexMatches = mc.Count: Exit Function
        If mc.Count = 0 Then: RegexMatches = Array(): Exit Function
        
        Dim arr()
        ReDim arr(1 To mc.Count)
        Dim i As Long
        For i = 1 To mc.Count
            If strProperty = "SubMatches" Then
                Dim sm As VBScript_RegExp_55.SubMatches
                Set sm = mc.Item(i - 1).SubMatches
                Dim subarr()
                ReDim subarr(1 To sm.Count)
                Dim j As Long
                For j = 1 To sm.Count
                    subarr(j) = sm.Item(j - 1)
                Next
                arr(i) = subarr
            Else
                arr(i) = CallByName(mc.Item(i - 1), strProperty, VbGet)
            End If
        Next
        RegexMatches = arr
    End With
End Function

Rem �}�b�`�����ӏ��̌�
Function RegexMatchCount(strSource As String, strPattern As String)
    RegexMatchCount = RegexMatches(strSource, strPattern, "Count")
End Function

Rem �}�b�`�����ӏ��̊J�n�C���f�b�N�X�z��
Function RegexMatchIndexs(strSource As String, strPattern As String)
    RegexMatchIndexs = RegexMatches(strSource, strPattern, "FirstIndex")
End Function

Rem �}�b�`�����ӏ��̕����񒷔z��
Function RegexMatchLengths(strSource As String, strPattern As String)
    RegexMatchLengths = RegexMatches(strSource, strPattern, "Length")
End Function

Rem �}�b�`�����ӏ��̒l�z��
Function RegexMatchValues(strSource As String, strPattern As String)
    RegexMatchValues = RegexMatches(strSource, strPattern, "Value")
End Function

Sub Test_RegexMatches()
    Debug.Print "-----Test_RegexMatches-----"
    Const src = "aabbcc axxyyzzc ghi jkl abbaac ghi"
    Const ptn = "a.+?c" '�ua�v�Ŏn�܂�uc�v�ŏI��镶����i�ŒZ�j�Ɉ�v
    Debug.Print RegexMatchCount(src, ptn)
    Debug.Print Join(RegexMatchIndexs(src, ptn), ",")
    Debug.Print Join(RegexMatchLengths(src, ptn), ",")
    Debug.Print Join(RegexMatchValues(src, ptn), ",")
    Debug.Print
End Sub

Rem �}�b�`�����ӏ��̔z��̃T�u�}�b�`�l�z��
Function RegexSubMatches(strSource As String, strPattern As String)
    RegexSubMatches = RegexMatches(strSource, strPattern, "SubMatches")
End Function

Sub Test_RegexSubMatches()
    Debug.Print "-----Test_RegexSubMatches-----"
    Const src = "AAAAA BB001 AA202 jk345 abcde i030k X12345"
    Const ptn = "([A-Z]+)([0-9]+)" '�u�A���t�@�x�b�g�啶���̃O���[�v�v�u���l�̃O���[�v�v�Ɉ�v
    
    Dim jagArr
    jagArr = RegexSubMatches(src, ptn)
    Stop
End Sub
