Attribute VB_Name = "kccWsFuncRegExp"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccWsFuncRegExp
Rem
Rem  @description   Excelワークシート用 正規表現 検証UDF集
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
Rem    2020/09/19 初回版作成
Rem    2020/10/28 ライセンス・ドキュメント追加
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem    |-----------------------------------------------------------------------------------------------------|
Rem    | 関数名            役割                                   戻り値                                     |
Rem    |-----------------------------------------------------------------------------------------------------|
Rem    | RegexIsMatch      マッチするかを確認                     True/False                                 |
Rem    | RegexReplace      マッチした文字列を置換                 置換後文字列                               |
Rem    | RegexMatches      マッチした任意のプロパティ情報         プロパティによって異なる                   |
Rem    | RegexMatchCount   マッチした箇所の個数                   Variant/Long                               |
Rem    | RegexMatchIndexs  マッチした箇所の開始インデックス配列   Variant/Long()                             |
Rem    | RegexMatchLengths マッチした箇所の文字列長配列           Variant/Long()                             |
Rem    | RegexMatchValues  マッチした箇所の値配列                 Variant/Variant()                          |
Rem    | RegexSubMatches   マッチした箇所の配列のサブマッチ値配列 Variant/Variant(1 to N)(Variant(1 to M))   |
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

Rem マッチするかを確認
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem
Rem  @return As Boolean     True:マッチした。False:マッチしなかった
Rem
Function RegexIsMatch(strSource As String, strPattern As String) As Boolean
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = False         ''大文字と小文字を区別する
        .Global = True              ''文字列全体を検索
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

Rem マッチした文字列を置換
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem  @param strReplace      置換文字列
Rem
Rem  @return As String      置換後の文字列
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

Rem マッチした任意のプロパティ情報
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem  @param strProperty     取得したいプロパティ
Rem                         未指定, Count, FirstIndex, Length, Value, SubMatches
Rem
Rem  @return As Variant     プロパティによって異なる
Rem                         未指定   VBScript_RegExp_55.MatchCollection
Rem                         Count    マッチした件数
Rem                         それ以外 (1 To N)の配列
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

Rem マッチした箇所の個数
Function RegexMatchCount(strSource As String, strPattern As String)
    RegexMatchCount = RegexMatches(strSource, strPattern, "Count")
End Function

Rem マッチした箇所の開始インデックス配列
Function RegexMatchIndexs(strSource As String, strPattern As String)
    RegexMatchIndexs = RegexMatches(strSource, strPattern, "FirstIndex")
End Function

Rem マッチした箇所の文字列長配列
Function RegexMatchLengths(strSource As String, strPattern As String)
    RegexMatchLengths = RegexMatches(strSource, strPattern, "Length")
End Function

Rem マッチした箇所の値配列
Function RegexMatchValues(strSource As String, strPattern As String)
    RegexMatchValues = RegexMatches(strSource, strPattern, "Value")
End Function

Sub Test_RegexMatches()
    Debug.Print "-----Test_RegexMatches-----"
    Const src = "aabbcc axxyyzzc ghi jkl abbaac ghi"
    Const ptn = "a.+?c" '「a」で始まり「c」で終わる文字列（最短）に一致
    Debug.Print RegexMatchCount(src, ptn)
    Debug.Print Join(RegexMatchIndexs(src, ptn), ",")
    Debug.Print Join(RegexMatchLengths(src, ptn), ",")
    Debug.Print Join(RegexMatchValues(src, ptn), ",")
    Debug.Print
End Sub

Rem マッチした箇所の配列のサブマッチ値配列
Function RegexSubMatches(strSource As String, strPattern As String)
    RegexSubMatches = RegexMatches(strSource, strPattern, "SubMatches")
End Function

Sub Test_RegexSubMatches()
    Debug.Print "-----Test_RegexSubMatches-----"
    Const src = "AAAAA BB001 AA202 jk345 abcde i030k X12345"
    Const ptn = "([A-Z]+)([0-9]+)" '「アルファベット大文字のグループ」「数値のグループ」に一致
    
    Dim jagArr
    jagArr = RegexSubMatches(src, ptn)
    Stop
End Sub
