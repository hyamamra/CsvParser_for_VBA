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


' CSVを2次元配列にパースするモジュール｡
' -------------------------------------
'
' 数値であればDouble型へ、文字列であればString型へパースする。
' 日付および通貨についてはサポートしていない。ただし、文字列型へ解釈された値を
' セルに貼り付ける際、Excel側が日付及び通貨として解釈可能な文字列である場合は、
' 動的に表示形式が変更される。
'
' 数値について
'
'   Double型へ変換可能な文字列であればDouble型へ変換される。
'   セルに貼り付けた時に表示形式を通貨にしたい場合、先頭に通貨記号を付与した上で
'   `"`で囲むことでDouble型へ変換されるのを防ぐことができる。
'
'   以下の要素はDouble型になる。
'   0,01,0.,.0,+.0,  -  0  ,"0,000","0,0,0","  000  ",$0
'
'   以下の要素はDobule型にならない。
'   +-1,"1  0",-.,"$1",1.1.1
'
' 文字列について
'
'   空文字列はEmpty値となる。
'   要素の1文字目に`"`が存在する場合、文字列の開始記号とみなし、要素には含めない。
'   また、閉じ記号までの値を文字列とみなす。それ以外の箇所で`"`が使用された場合は
'   文字列の一部とみなす。ただし、`"`で囲まれた文字列内で使用する場合はエスケープ
'   しなければならない。
'   `"`の外側に存在する値は文字列の一部とみなす。その際は`"`自体も文字列とみなす。
'   `,`および改行を文字列に含める場合、`"`で囲まなければならない。
'
' その他の仕様
'
'   列数が揃っていない場合、足りない要素をEmpty値で補完する。値が存在しない要素は
'   Empty値を挿入する。半角スペースおよびタブは要素の一部とみなす。ファイル末尾の
'   空の改行は無視される。改行コードは`\r\n`, `\r`, `\n`を混在させることができる。
'   ただし、`\r`の直後に`\n`が存在する場合は`\r\n`と解釈され、改行1回分となる。
'
'   文字セットパラメーターについてはADOの仕様に準拠する。
'   Charset プロパティ (ADO)
'   https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/charset-property-ado


' 解析対象。
Private Target As Target


' 字句解析の対象となるテキストの読み取り用データ。
Private Type Target
    ' 先頭から1文字ずつ取得していくための文字列。
    Text As String
    ' 字句解析中のTextの読み取り位置を示す。
    Cursor As Long
End Type


Private Function NewTarget(ByVal CsvText As String) As Target
    NewTarget.Text = CsvText
End Function


Public Function ParseCsv(ByVal FilePath As String, _
    Optional ByVal Charset As String = "UTF-8", _
    Optional ByVal SkipLinesTimes As Long) As Variant()
    ' CSVファイルを読み込み、2次元配列にパースする。
    ' セルに貼り付ける場合はこちらのメソッドの使用を推奨する。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   Charset: ファイルの文字セット。詳細はADOの仕様を参照。
    '   SkipLinesTimes: 先頭から読み飛ばす行数。

    Call LoadTextFile(FilePath, Charset)
    Call SkipLines(SkipLinesTimes)

    ParseCsv = To2dArray(Tokenize())
End Function


Public Function ParseCsvAsString(ByVal FilePath As String, _
    Optional ByVal Charset As String = "UTF-8", _
    Optional ByVal SkipLinesTimes As Long) As String()
    ' CSVファイルを読み込み、文字列型の2次元配列にパースする。
    '
    ' 引数
    '   FilePath: 読み込むファイルの絶対パス。
    '   Charset: ファイルの文字セット。詳細はADOの仕様を参照。
    '   SkipLinesTimes: 先頭から読み飛ばす行数。

    Call LoadTextFile(FilePath, Charset)
    Call SkipLines(SkipLinesTimes)

    ParseCsvAsString = To2dStringArray(Tokenize())
End Function


Public Function ParseCsvFromString(ByVal CsvText As String) As Variant()
    ' CSV形式の文字列を2次元配列にパースする。

    Target = NewTarget(CsvText)
    ParseCsvFromString = To2dArray(Tokenize())
End Function


Public Function ParseCsvFromStringAsString(ByVal CsvText As String) As String()
    ' CSV形式の文字列を文字列型の2次元配列にパースする。

    Target = NewTarget(CsvText)
    ParseCsvFromStringAsString = To2dStringArray(Tokenize())
End Function


' ________ CSVファイルからトークン集合を作成するためのメソッド群 ________

Private Function Tokenize() As Collection
    ' メンバー変数Targetを1文字ずつ読み取り、トークン化する。

    Dim Tokens As New Collection
    Call Tokens.Add(New Collection)

    Dim Char As String

    ' 要素の先頭1文字を読み取り、デリミタかトークンかを判断する。
    ' デリミタであればEmpty値をトークン集合に追加する。デリミタが
    ' 改行であればトークン集合に新しい行を追加する。
    ' トークンであればデリミタが出現するまでトークンとして読み取る。
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
    ' テキストファイルを開き、メンバー変数Targetにロードする。

    With CreateObject("ADODB.Stream")
        .Charset = Charset
        Call .Open
        Call .LoadFromFile(FilePath)
        Target = NewTarget(.ReadText())
        Call .Close
    End With
End Function


Private Function SkipLines(ByVal Times As Long)
    ' メンバー変数Targetから指定された行数分を読み飛ばす。

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
    ' メンバー変数Targetの読み取り位置が終端に達していればTrueを返す。
    EndOfFile = Len(Target.Text) = Target.Cursor
End Function


Private Function NextCharInto(ByRef Char As String) As Boolean
    ' メンバー変数Targetから次の文字を読み込む。
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
        Target.Cursor = Target.Cursor + 1
        Char = Mid(Target.Text, Target.Cursor, 1)
        NextCharInto = True
    End If
End Function


Private Function RewindCursor(Optional Times As Long = 1)
    ' メンバー変数Targetの読み取り位置を指定された数だけ戻す。
    Target.Cursor = Target.Cursor - Times
End Function


Private Function SkipIfNextCharIs(ByVal Char As String)
    ' メンバー変数Targetの次の文字が指定された文字なら読み飛ばす。

    Dim Temp As String
    If NextCharInto(Temp) Then
        If Temp <> Char Then Call RewindCursor
    End If
End Function


Private Function TokenInto(ByRef Tokens As Collection)
    ' 現在の読み取り位置から要素を取り出す。

    Dim Token As String
    Dim Char As String

    ' デリミタが出現するまでトークンに文字を追加する。
    Do While NextCharInto(Char)
        ' デリミタに該当すればトークンを確定する。
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

    ' メンバー変数Targetの終端に達していれば`"`が単体で使用されたと
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
                    ' デリミタのため、トークンを確定する。
                    Exit Do
                Case vbCr
                    ' デリミタのため、トークンを確定する。
                    Call Tokens(Tokens.Count).Add(Token)
                    Call Tokens.Add(New Collection)
                    Call SkipIfNextCharIs(vbLf)
                    Exit Function
                Case vbLf
                    ' デリミタのため、トークンを確定する。
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

    If Tokens.Count = 1 Then Exit Function

    Do
        Dim Line As Collection
        Set Line = Tokens(Tokens.Count)

        If 2 < Line.Count Then Exit Do

        If Line.Count = 1 Then
            If Line(1) <> Empty Then Exit Do
        End If

        Call Tokens.Remove(Tokens.Count)
    Loop
End Function


' ________ トークン集合から2次元配列を作成するためのメソッド群 ________

Private Function To2dArray(ByRef Tokens As Collection) As Variant()
    ' トークン集合を2次元配列に変換する。

    Dim RowLength As Long, ColumnLength As Long
    RowLength = Tokens.Count
    ColumnLength = CountColumns(Tokens)

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

                Call TryCDbl(Token)
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
    ColumnLength = CountColumns(Tokens)

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


Private Function TryCDbl(ByRef Token As Variant)
    ' Double型に変換可能であれば変換する。

    On Error Resume Next
    Token = CDbl(Token)
End Function


Private Function CountColumns(ByRef Tokens As Collection) As Long
    ' トークン集合の列数を求める。

    Dim Column As Collection
    For Each Column In Tokens
        If CountColumns < Column.Count Then
            CountColumns = Column.Count
        End If
    Next
End Function
