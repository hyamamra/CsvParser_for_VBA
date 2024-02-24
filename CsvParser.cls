VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' CSVファイルを2次元配列にパースするモジュール｡
' ---------------------------------------------
'
' 数値について
'
'   数値はDouble型に変換される。
'   間に半角スペースが挿入された数字はString型となる。ただし、数字の外側にしか
'   半角スペースが存在しなければDouble型となる。
'   またDouble型として解釈可能な場合、符号と数字の間の半角スペースは無視される。
'
'   以下の要素はDouble型に解釈される。
'   1,02,4.,.9,+.7,   -  6  ,"2,147,483,648"," 1,2,3 "
'
' 文字列について
'
'   空文字列はEmpty値となる。
'   タブが含まれる要素はString型となる。
'   要素の1文字目に`"`が存在する場合、文字列の開始記号とみなし、要素には含めない。
'   また、閉じ記号までの値を文字列とみなす。それ以外の箇所で`"`が使用された場合は
'   文字列の一部とみなす。ただし、`"`で囲まれた文字列内で使用する場合はエスケープ
'   しなければならない。
'   `"`の外側に存在する値は文字列の一部とみなす。その際は`"`自体も文字列とみなす。
'   `,`および改行を文字列に含める場合、`"`で囲まなければならない。
'
' 日時について
'
'   日時はDate型に変換される。
'
'
'   以下の要素はDate型に変換される。
'   2000/01/01,2000-01-01
'
'   以下の要素はDate型に変換されない。
'   2001/02/29,20001231
'
' 通貨について
'
'   通貨はCurrency型に変換される。
'   通貨記号が単体で存在し、後方に数値が存在する場合、その値は通貨とみなされる。
'   また、要素が`"`で囲まれており、適切に`,`で区切られていれば通貨と解釈される。
'   ただし、Currency型で表現できない場合はString型に変換される。
'
'   以下の要素はCurrency型に解釈される。
'   "1,234","$20","  \  100","6,000."  ,$.0,"$-0",+$000
'
'   以下の要素はCurrency型に解釈されない。
'   "1.234",  "$20","$$50","1,000.00001","$-0,000"
'
'   Currency型が表現できる範囲。小数部は4桁まで。
'   -922,337,203,685,477.5808 ～ 922,337,203,685,477.5807
'
' その他の仕様
'
'   半角スペースは要素の一部とみなす。
'   値が存在しない要素はEmpty値を挿入する。
'   列数が揃っていない場合、足りない要素をEmpty値で補完する。
'   改行コードは`\r\n`, `\r`, `\n`を使用できる。ただし、混在する場合は不正な動作を
'   する場合がある。また、ファイル末尾の空の改行は無視される。
'
'   文字セットパラメーターについては、ADODBの仕様に準ずる。
'   Charset プロパティ (ADO)
'   https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/charset-property-ado


' CSVファイルから読み込んだ文字の配列を含む構造体。
' 字句解析の際に先頭から1文字ずつ取得していく。
Private FileContent As String

' 字句解析中のFileContentの読み取り位置を示す。
Private FileCursor As Long


Private Enum TokenKind
    StringToken
    NumberToken
    DateToken
    CurrencyToken
    EmptyToken
End Enum


Public Function ParseCsv(ByVal FilePath As String, _
    Optional ByVal Charset As String = "UTF-8", _
    Optional ByVal SkipLinesTimes As Long) As Variant()
    ' CSVファイルを読み込み、2次元配列に変換する。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   SkipLinesTimes: 先頭から読み飛ばす行数。

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, Charset, SkipLinesTimes)

    ParseCsv = To2dArray(Tokens)
End Function


Public Function ParseCsvAsString(ByVal FilePath As String, _
    Optional ByVal Charset As String = "UTF-8", _
    Optional ByVal SkipLinesTimes As Long) As String()
    ' CSVファイルを読み込み、文字列型の2次元配列に変換する。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   SkipLinesTimes: 先頭から読み飛ばす行数。

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, Charset, SkipLinesTimes)

    ParseCsvAsString = To2dStringArray(Tokens)
End Function


' ________ CSVファイルからトークン集合を作成するための関数群 ________

Private Function Tokenize(ByVal FilePath As String, _
    ByVal Charset As String, ByVal SkipLinesTimes As Long) As Collection
    ' FileContentを1文字ずつ読み取り、トークン化する。

    Call LoadTextFile(FilePath, Charset)
    Call SkipLines(SkipLinesTimes)

    Dim Tokens As New Collection
    Call Tokens.Add(New Collection)

    Dim Char As String
    Do While NextCharInto(Char)
        Select Case Char
            Case ","
                Call Tokens(Tokens.Count).Add(Empty)

            Case vbCr
                Call Tokens(Tokens.Count).Add(Empty)
                Call Tokens.Add(New Collection)
                Call SkipIfNextCharIs(vbLf)

            Case vbLf
                Call Tokens(Tokens.Count).Add(Empty)
                Call Tokens.Add(New Collection)

            Case """"
                Call QuotedTokenInto(Tokens)

            Case Else
                Call RewindCursor
                Call TokenInto(Tokens)
        End Select
    Loop

    Call RemoveEmptyLinesAtEnd(Tokens)
    Set Tokenize = Tokens
End Function


Private Function LoadTextFile( _
    ByVal FilePath As String, ByVal Charset As String)
    ' テキストファイルを開き、メンバー変数FileContentにロードする。

    With New ADODB.Stream
        .Charset = Charset
        Call .Open
        Call .LoadFromFile(FilePath)

        FileContent = .ReadText()
        Call .Close
    End With
End Function


Private Function SkipLines(ByVal Times As Long)
    ' メンバー変数FileContentから指定された行数分を読み飛ばす。

    Dim LineNumber As Long
    Do Until LineNumber = Times
        Dim Char As String
        If NextCharInto(Char) Then
            Select Case Char
                Case vbCr
                    LineNumber = LineNumber + 1
                    Call SkipIfNextCharIs(vbLf)
                Case vbLf
                    LineNumber = LineNumber + 1
            End Select
        Else
            Exit Do
        End If
    Loop
End Function


Private Function EndOfFile() As Boolean
    ' FileContentの読み取り位置が終端に達していればTrueを返す。
    EndOfFile = Len(FileContent) = FileCursor
End Function


Private Function NextCharInto(ByRef Char As String) As Boolean
    ' FileContentから次の文字を読み込む。
    '
    ' 引数
    '   Char:
    '       読み取った文字を代入する。
    '       読み取れなかった場合は空文字列を代入する。
    '
    ' 戻り値
    '   次の文字が読み取れればTrueを返す。

    If EndOfFile() Then
        Char = ""
    Else
        FileCursor = FileCursor + 1
        Char = Mid(FileContent, FileCursor, 1)
        NextCharInto = True
    End If
End Function


Private Function RewindCursor(Optional Times As Long = 1)
    ' FileContentの読み取り位置を指定された数だけ戻す。
    FileCursor = FileCursor - Times
End Function


Private Function SkipIfNextCharIs(ByVal Char As String)
    ' FileContentの次の文字が指定された文字なら読み飛ばす。

    Dim Temp As String
    If NextCharInto(Temp) Then
        If Temp <> Char Then Call RewindCursor
    End If
End Function


Private Function TokenInto(ByRef Tokens As Collection)
    ' 現在の読み取り位置から要素を取り出す。

    Dim Token As String
    Dim Char As String

    ' 区切り文字が出現するまでトークンに文字を追加する。
    Do While NextCharInto(Char)
        ' 区切り文字に該当すればトークンを確定する。
        Select Case Char
            Case ","
                Call Tokens(Tokens.Count).Add(Token)
                Exit Function
            Case vbCr
                Call Tokens(Tokens.Count).Add(Token)
                Call Tokens.Add(New Collection)
                Call SkipIfNextCharIs(vbLf)
                Exit Function
            Case vbLf
                Call Tokens(Tokens.Count).Add(Token)
                Call Tokens.Add(New Collection)
                Exit Function
        End Select

        Token = Token + Char
    Loop

    Call Tokens(Tokens.Count).Add(Token)
End Function


Private Function QuotedTokenInto(ByRef Tokens As Collection)
    ' 現在の読み取り位置から要素を取り出す。

    ' FileContentの終端に達していれば`"`が単体で使用されたと
    ' 判断し、`"`そのものをトークン化して処理を終了する。
    If EndOfFile() Then
        Call Tokens(Tokens.Count).Add("""")
        Exit Function
    End If

    Dim Token As String
    Dim Char As String

    Do While NextCharInto(Char)
        If Char = """" Then
            ' ファイルの終端のため、トークンを確定する。
            If Not NextCharInto(Char) Then
                Exit Do
            End If

            Select Case Char
                Case ","
                    ' 区切り文字のため、トークンを確定する。
                    Exit Do
                Case vbCr
                    ' 区切り文字のため、トークンを確定する。
                    Call Tokens(Tokens.Count).Add(Token)
                    Call Tokens.Add(New Collection)
                    Call SkipIfNextCharIs(vbLf)
                    Exit Function
                Case vbLf
                    ' 区切り文字のため、トークンを確定する。
                    Call Tokens(Tokens.Count).Add(Token)
                    Call Tokens.Add(New Collection)
                    Exit Function
                Case """"
                    ' エスケープ文字とみなし、無視する。
                Case Else
                    ' エスケープ忘れとみなし、要素に追加する。
                    Char = """" & Char
            End Select
        End If

        Token = Token + Char
    Loop

    Call Tokens(Tokens.Count).Add(Token)
End Function


Private Function RemoveEmptyLinesAtEnd(ByRef Tokens As Collection)
    ' トークン集合末尾の不要な改行を削除する。

    Do
        Dim Line As Collection
        Set Line = Tokens(Tokens.Count)

        If 2 < Line.Count Then
            Exit Do
        ElseIf Line.Count = 1 Then
            If Line(1) = Empty Then
                Call Tokens.Remove(Tokens.Count)
            Else
                Exit Do
            End If
        Else
            Call Tokens.Remove(Tokens.Count)
        End If
    Loop
End Function


' ________ トークン集合から2次元配列を作成するための関数群 ________

Private Function To2dArray(ByRef Tokens As Collection) As Variant()
    ' トークン集合を2次元配列に変換する。

    Dim RowLength As Long, ColumnLength As Long
    RowLength = Tokens.Count
    ColumnLength = CountColumnLength(Tokens)

    Dim TokensArray() As Variant

    If ColumnLength = 0 Then
        ReDim TokensArray(0, 0)
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

                Select Case KindOf(Token)
                    Case StringToken
                        ' do nothing
                    Case Else
                        Stop ' 未実装。
                End Select

                TokensArray(RowIndex - 1, ColumnIndex - 1) = Token
            End If
        Next
    Next

    To2dArray = TokensArray
End Function


Private Function To2dStringArray(ByRef Tokens As Collection) As String()
    ' トークン集合を文字列型の2次元配列に変換する。

    Dim RowLength As Long, ColumnLength As Long
    RowLength = Tokens.Count
    ColumnLength = CountColumnLength(Tokens)

    Dim TokensArray() As String

    If ColumnLength = 0 Then
        ReDim TokensArray(0, 0)
        TokensArray(0, 0) = ""
        To2dStringArray = TokensArray
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

                TokensArray(RowIndex - 1, ColumnIndex - 1) = Token
            End If
        Next
    Next

    To2dStringArray = TokensArray
End Function


Private Function KindOf(ByVal Token As Variant) As TokenKind
    ' 文字列からトークンの種類を判定する。

    ' 未実装。テスト用にすべて文字列と判定する。
    KindOf = StringToken
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
