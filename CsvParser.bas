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
'   値に含まれない空白文字は無視される。
'   ファイル末尾の改行は無視される。
'   単項演算子`+`と`-`を受け入れる。
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


Private Function OpenFileAsReadOnly(ByVal PathName As String)
    ' 読み取り専用でテキストファイルを開く。
    ' 一度に全てを読み込まないことで省メモリ化を図る。

    FileNumber = FreeFile()
    Open PathName _
        For Input _
        Access Read _
        Lock Write _
        As #FileNumber
End Function


Private Function CloseFile()
    Close #FileNumber
End Function


Private Function NextChar(ByRef Into As String) As Boolean
    ' 次の文字を読み込む。
    '
    ' 引数
    '   Into:
    '       読み取った文字を代入する。
    '       読み取れない場合、空文字列を代入する。
    '
    ' 戻り値
    '   次の文字が読み取れればTrueを返す。

    Into = ""
    If Not EOF(FileNumber) Then
        Into = Input(1, #FileNumber)
        NextChar = True
    End If
End Function


Private Function RewindChar()
    ' 現在の読み取り位置を1文字分戻す。
    Seek #FileNumber, Seek(FileNumber) - 1
End Function


Private Function Tokenize(ByVal FilePath As String, ByVal SkipLines As Long) As Tokens
    ' CSVファイルを1文字ずつ読み取り、トークン化する。
    ' トークンには数値、文字列、カンマ、改行、空要素が存在する。

    Dim Tokens As Tokens
    Set Tokens.Kinds = New Collection
    Set Tokens.Values = New Collection

    Call OpenFileAsReadOnly(FilePath)
    On Error GoTo ErrorHandler

    Dim Char As String
    Do While NextChar(Into:=Char)
        Select Case Char
            Case ","
                Call TakeCommaToken(Into:=Tokens)

            Case vbCr
                Call TakeNewLineToken(Into:=Tokens, Char:=Char)

            Case vbLf
                Call TakeNewLineToken(Into:=Tokens, Char:=Char)

            Case vbTab
                ' 読み飛ばす。

            Case " "
                ' 読み飛ばす。

            Case """"
                Call TakeStringToken(Into:=Tokens)

            Case Else
                If Char Like "[-+.0-9]" Then
                    Call RewindChar
                    Call TakeDigitsToken(Into:=Tokens)
                Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " 文字目に不正な文字 `" _
                        & Char & "` を検出しました。")
                End If
        End Select
    Loop

ErrorHandler:
    Call CloseFile
    If Err Then Call Err.Raise(Err)
    Tokenize = Tokens
End Function


Private Function SkipLines(ByVal Times As Long)
    ' 開いているファイルから指定された行数分を読み飛ばす。

    Dim T As Long
    For T = 0 To Times
        If Not EOF(FileNumber) Then
            Dim Temp As String
            Line Input #FileNumber, Temp
        End If
    Next
End Function


Private Function TakeCommaToken(ByRef Into As Tokens)
    ' 必要であれば空トークンを追加してからカンマトークンを追加する。

    Call AddEmptyTokenIfLastTokenIsDelimiter(Into)
    Call Into.Kinds.Add(TokenKind.CommaToken)
    Call Into.Values.Add(Empty)
End Function


Private Function TakeNewLineToken(ByRef Into As Tokens, ByVal Char As String)
    ' 必要であれば空トークンを追加してから改行トークンを追加する。
    ' `\r`を検出した場合、次の文字が`\n`であれば`\r\n`とみなし、`\n`を読み飛ばす。

    Call AddEmptyTokenIfLastTokenIsDelimiter(Into)
    Call Into.Kinds.Add(TokenKind.NewLineToken)
    Call Into.Values.Add(Empty)

    If Char = vbCr Then
        If NextChar(Into:=Char) Then
            If Char <> vbLf Then Call RewindChar
        End If
    End If
End Function


Private Function AddEmptyTokenIfLastTokenIsDelimiter(ByRef Into As Tokens)
    ' 最後に切り出したトークンが改行、カンマである、または
    ' 切り出したトークンがまだなければ空トークンを追加する。

    If Into.Kinds.Count = 0 Then
        Call AddEmptyToken(Into)
    Else
        Dim LastToken As TokenKind
        LastToken = Into.Kinds(Into.Kinds.Count)
        If LastToken = TokenKind.CommaToken _
        Or LastToken = TokenKind.NewLineToken Then
            Call AddEmptyToken(Into)
        End If
    End If
End Function


Private Function AddEmptyToken(ByRef Into As Tokens)
    Call Into.Kinds.Add(TokenKind.EmptyToken)
    Call Into.Values.Add(Empty)
End Function


Private Function TakeStringToken(ByRef Into As Tokens)
    ' 現在の読み取り位置から文字列を取り出し、Tokensに格納する。

    Dim Value As String
    Dim Char As String

    Do
        If Not NextChar(Into:=Char) Then
            Call Err.Raise(513, , _
                "文字列はダブルクォーテーションで閉じる必要があります。" _
                & vbNewLine & "検出した文字列: " & Value)
        End If

        ' `"`を検出した場合、次の文字も`"`であれば
        ' 文字列の閉じ記号とみなし、トークンを確定する。
        If Char = """" Then
            If Not NextChar(Into:=Char) Then Exit Do
            If Char <> """" Then
                Call RewindChar
                Exit Do
            End If
        End If

        Value = Value + Char
    Loop

    Call Into.Kinds.Add(TokenKind.StringToken)
    Call Into.Values.Add(Value)
End Function


Private Function TakeDigitsToken(ByRef Into As Tokens)
    ' 現在の読み取り位置から数値を取り出し、Tokensに格納する。

    Dim Value As String

    Dim Char As String
    Call NextChar(Into:=Char)
    Value = Char

    ' 数字、小数点以外を検出したら切り出しを終了する。
    Do While NextChar(Into:=Char)
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

    Call Into.Kinds.Add(TokenKind.DigitsToken)
    Call Into.Values.Add(Value)
End Function
