Attribute VB_Name = "CsvParser"
Option Explicit


' 概要
'   CSVを読み込むモジュール。
'
'   以下の値は読み取りをサポートしている。
'
'   ・整数、負数、小数（Double型に変換可能な値）。
'   ・改行を含む文字列。
'   ・エスケープ済みのダブルクォーテーションを含む文字列。
'
'   列数が揃っていない場合、空の要素で補完する。
'   区切り文字のキャリッジリターン（"\r"）は無視される。
'   値に含まれない空白文字は無視される。
'   ファイル末尾の改行は無視される。
'   単項演算子は"+"および"-"を受け入れる。
'   数値内の空白文字を許容しない。
'   数値はDouble型にパースされるため、正確な数値が必要な場合は
'   AsStringオプションをTrueにすることで文字列として取得できる。


Private FileNumber As Long


Private Enum TokenKind
    StringToken
    DigitsToken
    EmptyToken
    CommaToken
    NewLineToken
End Enum


Private Type Token
    Kind As TokenKind
    Value As String
End Type


Private Type Tokens
    Kinds As Collection
    Values As Collection
End Type


Public Function ReadCsv( _
    ByVal FilePath As String, _
    Optional ByVal AsString As Boolean = False, _
    Optional ByVal SkipLines As Long = 0) As Collection
    ' CSVファイルを読み込み、二次元配列に変換する。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   AsString: すべての要素を文字列として読み込む。
    '   SkipLines: 先頭から読み飛ばす行数。

    Dim Tokens As Tokens
    Tokens = Tokenize(FilePath, SkipLines)

    Set ReadCsv = Tokens.Values
End Function


Private Sub OpenFileAsReadOnly(ByVal PathName As String)
    ' 読み取り専用でテキストファイルを開く。
    ' 一度に全てを読み込まないことで省メモリ化を図る。

    FileNumber = FreeFile()
    Open PathName _
        For Input _
        Access Read _
        Lock Write _
        As #FileNumber
End Sub


Private Sub CloseFile()
    Close #FileNumber
End Sub


Private Function NextChar(ByRef Char As String) As Boolean
    ' 次の文字を読み込む。
    '
    ' 引数
    '   Char:
    '       読み取った文字を代入する。
    '       読み取れない場合、空文字列を代入する。
    '
    ' 戻り値
    '   次の文字が読み取れればTrueを返す。

    Char = ""
    If Not EOF(FileNumber) Then
        Char = Input(1, #FileNumber)
        NextChar = True
    End If
End Function


Private Sub RewindChar()
    ' 現在の読み取り位置を1文字分戻す。
    Seek #FileNumber, Seek(FileNumber) - 1
End Sub


Private Function Tokenize( _
    ByVal FilePath As String, _
    ByVal SkipLines As Long) As Tokens

    Dim Tokens As Tokens
    Set Tokens.Kinds = New Collection
    Set Tokens.Values = New Collection

    Call OpenFileAsReadOnly(FilePath)
    On Error GoTo ErrorHandler

    Dim Char As String
    Do While NextChar(Char)
        Select Case Char
            Case ","
                Call AddCommaToken(Tokens)

            Case vbLf
                Call AddNewLineToken(Tokens)

            Case vbCr
            Case vbTab
            Case " "
                GoTo Continue

            Case """"
                Call AddStringToken(Tokens)

            Case Else
                If Char Like "[-+.0-9]" Then
                    Call RewindChar
                    Call AddDigitsToken(Tokens)
                Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " 文字目に不正な文字 """ _
                        & Char & """ を検出しました。")
                End If
        End Select
Continue:
    Loop

ErrorHandler:
    Call CloseFile
    If Err Then Call Err.Raise(Err)
    Tokenize = Tokens
End Function


Private Sub AddEmptyToken(ByRef Tokens As Tokens)
    Call Tokens.Kinds.Add(TokenKind.EmptyToken)
    Call Tokens.Values.Add(Empty)
End Sub


Private Sub AddCommaToken(ByRef Tokens As Tokens)
    Call AddEmptyTokenIfLastTokenIsDelimiter(Tokens)
    Call Tokens.Kinds.Add(TokenKind.CommaToken)
    Call Tokens.Values.Add(Empty)
End Sub


Private Sub AddNewLineToken(ByRef Tokens As Tokens)
    Call AddEmptyTokenIfLastTokenIsDelimiter(Tokens)
    Call Tokens.Kinds.Add(TokenKind.NewLineToken)
    Call Tokens.Values.Add(Empty)
End Sub


Private Sub AddEmptyTokenIfLastTokenIsDelimiter(ByRef Tokens As Tokens)
    If 0 < Tokens.Kinds.Count Then
        Dim LastToken As TokenKind
        LastToken = Tokens.Kinds(Tokens.Kinds.Count)
        If LastToken = TokenKind.CommaToken _
        Or LastToken = TokenKind.NewLineToken Then
            Call AddEmptyToken(Tokens)
        End If
    End If
End Sub


Private Function AddStringToken(ByRef Tokens As Tokens)
    Dim Value As String
    Dim Char As String

    Do
        If Not NextChar(Char) Then
            Call Err.Raise(513, , _
                "文字列はダブルクォーテーションで閉じる必要があります。" _
                & vbNewLine & "検出した文字列: " & Value)
        End If

        ' `"`を検出した場合、次の文字も`"`であれば
        ' 文字列の閉じ記号とみなし、トークンを確定する。
        If Char = """" Then
            If Not NextChar(Char) Then Exit Do
            If Char <> """" Then
                Call RewindChar
                Exit Do
            End If
        End If

        Value = Value + Char
    Loop

    Call Tokens.Kinds.Add(TokenKind.StringToken)
    Call Tokens.Values.Add(Value)
End Function


Private Function AddDigitsToken(ByRef Tokens As Tokens)
    Dim Value As String

    Dim Char As String
    Call NextChar(Char)
    Value = Char

    Do While NextChar(Char)
        If Char Like "[.0-9]" Then
            Value = Value + Char
        Else
            Call RewindChar
            Exit Do
        End If
    Loop

    If (Value Like "[-+.]") Or (Value Like "[-+].") Then
        Call Err.Raise(513, , _
            "不正な数値を検出しました。" _
            & vbNewLine & "最低でも1つの数字が必要です。" _
            & vbNewLine & "検出した数値: " & Value)
    ElseIf Value Like "*.*.*" Then
        Call Err.Raise(513, , _
            "不正な数値を検出しました。" _
            & vbNewLine & "小数点が複数存在します。" _
            & vbNewLine & "検出した数値: " & Value)
    End If

    Call Tokens.Kinds.Add(TokenKind.DigitsToken)
    Call Tokens.Values.Add(Value)
End Function
