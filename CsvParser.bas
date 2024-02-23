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
    Str
    Num
    Emp
End Enum


Public Function ReadCsv( _
    ByVal FilePath As String, _
    Optional ByVal NumberOfSkipLines As Long) As Variant()
    ' CSVファイルを読み込み、二次元配列に変換する。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   NumberOfSkipLines:
    '       先頭から読み飛ばす行数。
    '       読み飛ばした先が文字列中の場合、例外が発生する。

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, NumberOfSkipLines)

    ReadCsv = To2dArray(Tokens)
End Function


Private Function To2dArray(ByRef Tokens As Collection) As Variant()
    ' トークン集合を二次元配列に変換する。

    Call RemoveEmptyLinesAtEnd(Tokens)

    Dim RowLength As Long, ColumnLength As Long
    RowLength = Tokens.Count
    ColumnLength = CountColumnLength(Tokens)

    Dim TokensArray() As Variant

    If ColumnLength = 0 Then
        ReDim Arr(0, 0)
        To2dArray = TokensArray
        Exit Function
    End If

    ReDim TokensArray(RowLength - 1, ColumnLength - 1)

    Dim RowIndex As Long, ColumnIndex As Long
    For RowIndex = 1 To RowLength
        For ColumnIndex = 1 To ColumnLength
            If 0 < Tokens(RowIndex).Count Then
                If Tokens(RowIndex).Count < ColumnIndex Then Exit For

                Dim Token As Variant
                Token = Tokens(RowIndex)(ColumnIndex)

                Select Case Left(Token, 1)
                    Case CStr(TokenKind.Str)
                        Token = Right(Token, Len(Token) - 1)
                    Case CStr(TokenKind.Num)
                        Token = CDbl(Right(Token, Len(Token) - 1))
                    Case CStr(TokenKind.Emp)
                        Token = Empty
                    End Select

                TokensArray(RowIndex - 1, ColumnIndex - 1) = Token
            End If
        Next
    Next

    To2dArray = TokensArray
End Function


Private Function RemoveEmptyLinesAtEnd(ByRef Tokens As Collection)
    ' トークン集合末尾の不要な改行を削除する。

    Do
        Dim Last As Collection
        Set Last = Tokens(Tokens.Count)

        If 2 < Last.Count Then
            Exit Do
        ElseIf Last.Count = 1 Then
            If Last(1) = CStr(TokenKind.Emp) Then
                Call Tokens.Remove(Tokens.Count)
            End If
        Else
            Call Tokens.Remove(Tokens.Count)
        End If
    Loop
End Function


Private Function CountColumnLength(ByRef Tokens As Collection) As Long
    ' トークン集合の列数を求める。

    Dim Column As Collection
    For Each Column In Tokens
        If CountColumnLength < Column.Count Then
            CountColumnLength = Column.Count
        End If
    Next
End Function


Private Function Tokenize( _
    ByVal FilePath As String, _
    ByVal NumberOfSkipLines As Long) As Collection
    ' CSVファイルを1文字ずつ読み取り、トークン化する。
    ' トークンには数値、文字列、カンマ、改行、空要素が存在する。

    Dim Tokens As New Collection
    Call Tokens.Add(New Collection)

    Call OpenFileAsReadOnly(FilePath)
    Call SkipLines(NumberOfSkipLines)

    ' トークンの読み取り後、区切り文字を
    ' 待機している状態であればTrueとなる。
    Dim WaitingForDelimiter As Boolean

    On Error GoTo ErrorHandler

    Dim Char As String
    Do While NextChar(Into:=Char)
        If (Char = " ") Or (Char = vbTab) Then GoTo Continue

        If WaitingForDelimiter Then
            Select Case Char
                Case ","
                    ' 何もしない。
                Case vbLf
                    Call Tokens.Add(New Collection)
                Case vbCr
                    Call Tokens.Add(New Collection)
                    Call SkipIfNextCharIsLf
                Case Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " 文字目に不正な文字 `" _
                        & Char & "` を検出しました。")
            End Select

            WaitingForDelimiter = False
            GoTo Continue
        End If

        Select Case Char
            Case ","
                Call Tokens(Tokens.Count).Add(EmptyToken())

            Case vbCr
                Call Tokens(Tokens.Count).Add(EmptyToken())
                Call Tokens.Add(New Collection)
                Call SkipIfNextCharIsLf

            Case vbLf
                Call Tokens(Tokens.Count).Add(EmptyToken())
                Call Tokens.Add(New Collection)

            Case """"
                Call Tokens(Tokens.Count).Add(StringToken())
                WaitingForDelimiter = True

            Case Else
                If Char Like "[-+.0-9]" Then
                    Call RewindChar
                    Call Tokens(Tokens.Count).Add(NumberToken())
                    WaitingForDelimiter = True
                Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " 文字目に不正な文字 `" _
                        & Char & "` を検出しました。")
                End If
        End Select
Continue:
    Loop

    Set Tokenize = Tokens

ErrorHandler:
    Call CloseFile
    If Err Then Call Err.Raise(Err)
End Function


Private Function StringToken() As String
    ' 現在の読み取り位置から文字列を取り出す。

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

    StringToken = CStr(TokenKind.Str) & Value
End Function


Private Function NumberToken() As String
    ' 現在の読み取り位置から数値を取り出す。

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

    NumberToken = CStr(TokenKind.Num) & Value
End Function


Private Function EmptyToken() As String
    EmptyToken = CStr(TokenKind.Emp)
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


Private Function SkipLines(ByVal Times As Long)
    ' 開いているファイルから指定された行数分を読み飛ばす。

    Dim T As Long
    For T = 1 To Times
        If Not EOF(FileNumber) Then
            Dim Temp As String
            Line Input #FileNumber, Temp
        End If
    Next
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


Private Function SkipIfNextCharIsLf()
    ' 次の文字が`\n`なら読み飛ばす。
    Dim Char As String
    If NextChar(Char) Then
        If Char <> vbLf Then Call RewindChar
    End If
End Function
