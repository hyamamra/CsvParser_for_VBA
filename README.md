# CsvParser for VBA

## 使い方
```VBA
Sub Sample()
    ' 戻り値はVariant型の動的配列となる。
    Dim Csv() As Variant
    Csv = CsvParser.ParseCsv(FILE_PATH)

    ' 配列の大きさを取得する。
    Dim RowLength As Long, ColumnLength As Long
    RowLength = UBound(Csv, 1)
    ColumnLength = UBound(Csv, 2)

    ' ワークシートの先頭を基準にCSVの内容を貼り付ける。
    ActiveSheet.Range(Cells(1, 1), _
        Cells(RowLength, ColumnLength)).Value = Csv
End Sub
```

ヘッダーがあると正しく読み込めません。  
その場合は、先頭行を指定行数無視することができます。

```VBA
' 第2引数に無視する行数を指定する。
Dim Csv() As Variant
Csv = CsvParser.ParseCsv(FILE_PATH, NumberOfSkipLines:=20)
```

すべての値を文字列として読み込むことも可能です。

```VBA
' 戻り値はString型の動的配列となる。
Dim Csv() As String
Csv = CsvParser.ParseCsvAsString(FILE_PATH)
```

もちろん先頭行を無視することもできます。

```VBA
Dim Csv() As String
Csv = CsvParser.ParseCsvAsString(FILE_PATH, NumberOfSkipLines:=20)
```
