# CsvParser for VBA

CSVファイルおよびCSV形式の文字列を2次元配列にパースするためのクラスです。  
Double型とString型をサポートしています。

### ParseCsvメソッド

```vba
Sub Sample()
    Dim parser As New CsvParser

    ' Variant型の動的2次元配列に解析結果を格納します。
    Dim csv() As Variant
    csv = parser.ParseCsv(FILE_PATH)

    ' 配列の行数と列数を取得します。
    Dim rowLength As Long, rolumnLength As Long
    rowLength = UBound(csv, 1) + 1
    columnLength = UBound(csv, 2) + 1

    ' A1セルを基準に配列の内容を貼り付けます。
    With ActiveSheet
        .Range(.Range("A1"), .Cells(rowLength, columnLength)).Value = csv
    End With
End Sub
```

UTF-8以外のファイルを読み込む場合は文字セットを指定してください。  
使用可能な文字セットについてはADOに準拠します。
[[Charset プロパティ (ADO)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/charset-property-ado)]

```vba
' 第2引数に文字セットを指定します。
Dim csv() As Variant
csv = parser.ParseCsv(FILE_PATH, "Shift-JIS")
```

先頭から指定行数無視することができます。

```vba
' 第3引数に無視する行数を指定します。
Dim csv() As Variant
csv = parser.ParseCsv(FILE_PATH, , SkipLinesTimes:=20)
```

### ParseCsvAsStringメソッド

すべての値を文字列として読み込みます。

```vba
Dim csv() As String
csv = parser.ParseCsvAsString(FILE_PATH)
```

### ParseCsvFromStringメソッド

文字列をCSV形式として読み込みます。

```vba
Dim str as String
str = "text,10,""2000/01/01"",""$2,000.00"""

Dim csv() As Variant
csv = parser.ParseCsvFromString(str)
```

### ParseCsvFromStringAsStringメソッド

文字列をCSV形式として読み込み、文字列型の配列に変換します。

```vba
Dim str as String
str = "text,10,""2000/01/01"",""$2,000.00"""

Dim csv() As String
csv = parser.ParseCsvFromStringAsString(str)
```

## 詳細な仕様

数値であればDouble型へ、文字列であればString型へパースします。  
日付および通貨についてはサポートしておりません。ただし、文字列型へ解釈された値をセルに貼り付ける際、Excel側が日付及び通貨として解釈可能な文字列である場合は、動的に表示形式が変更されます。

### 数値について

Double型へ変換可能な文字列であればDouble型へ変換されます。  
セルに貼り付けた時に表示形式を通貨にしたい場合、先頭に通貨記号を付与した上で`"`で囲むことでDouble型へ変換されるのを防ぐことが可能です。

以下の要素はDouble型になります。  
```csv
0,01,0.,.0,+.0,  -  0  ,"0,000","0,0,0","  000  ",$0
```

以下の要素はDobule型になりません。  
```csv
+-1,"1  0",-.,"$1",1.1.1
```

### 文字列について

空文字列はEmpty値となります。  
要素の1文字目に`"`が存在する場合、文字列の開始記号とみなし、要素には含めません。また、閉じ記号までの値を文字列とみなされます。それ以外の箇所で`"`が使用された場合は文字列の一部とみなされます。ただし、`"`で囲まれた文字列内で使用する場合はエスケープが必要です。  
`"`の外側に存在する値は文字列の一部とみなされます。その際は`"`自体も文字列と解釈されます。`,`および改行を文字列に含める場合、`"`で囲む必要があります。

### その他の仕様

列数が揃っていない場合、足りない要素をEmpty値で補完します。値が存在しない要素にはEmpty値が挿入されます。半角スペースおよびタブは要素の一部とみなされます。また、ファイル末尾の空の改行は無視されます。  
改行コードは`\r\n`, `\r`, `\n`を混在させることができます。ただし、`\r`の直後に`\n`が存在する場合は`\r\n`と解釈され、改行1回分となります。
