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


' CSV�t�@�C����2�����z��Ƀp�[�X���郂�W���[���
' ---------------------------------------------
'
' ���l�ɂ���
'
'   ���l��Double�^�ɕϊ������B
'   �Ԃɔ��p�X�y�[�X���}�����ꂽ������String�^�ƂȂ�B�������A�����̊O���ɂ���
'   ���p�X�y�[�X�����݂��Ȃ����Double�^�ƂȂ�B
'   �܂�Double�^�Ƃ��ĉ��߉\�ȏꍇ�A�����Ɛ����̊Ԃ̔��p�X�y�[�X�͖��������B
'
'   �ȉ��̗v�f��Double�^�ɉ��߂����B
'   1,02,4.,.9,+.7,   -  6  ,"2,147,483,648"," 1,2,3 "
'
' ������ɂ���
'
'   �󕶎����Empty�l�ƂȂ�B
'   �^�u���܂܂��v�f��String�^�ƂȂ�B
'   �v�f��1�����ڂ�`"`�����݂���ꍇ�A������̊J�n�L���Ƃ݂Ȃ��A�v�f�ɂ͊܂߂Ȃ��B
'   �܂��A���L���܂ł̒l�𕶎���Ƃ݂Ȃ��B����ȊO�̉ӏ���`"`���g�p���ꂽ�ꍇ��
'   ������̈ꕔ�Ƃ݂Ȃ��B�������A`"`�ň͂܂ꂽ��������Ŏg�p����ꍇ�̓G�X�P�[�v
'   ���Ȃ���΂Ȃ�Ȃ��B
'   `"`�̊O���ɑ��݂���l�͕�����̈ꕔ�Ƃ݂Ȃ��B���̍ۂ�`"`���̂�������Ƃ݂Ȃ��B
'   `,`����щ��s�𕶎���Ɋ܂߂�ꍇ�A`"`�ň͂܂Ȃ���΂Ȃ�Ȃ��B
'
' �����ɂ���
'
'   ������Date�^�ɕϊ������B
'
'
'   �ȉ��̗v�f��Date�^�ɕϊ������B
'   2000/01/01,2000-01-01
'
'   �ȉ��̗v�f��Date�^�ɕϊ�����Ȃ��B
'   2001/02/29,20001231
'
' �ʉ݂ɂ���
'
'   �ʉ݂�Currency�^�ɕϊ������B
'   �ʉ݋L�����P�̂ő��݂��A����ɐ��l�����݂���ꍇ�A���̒l�͒ʉ݂Ƃ݂Ȃ����B
'   �܂��A�v�f��`"`�ň͂܂�Ă���A�K�؂�`,`�ŋ�؂��Ă���Βʉ݂Ɖ��߂����B
'   �������ACurrency�^�ŕ\���ł��Ȃ��ꍇ��String�^�ɕϊ������B
'
'   �ȉ��̗v�f��Currency�^�ɉ��߂����B
'   "1,234","$20","  \  100","6,000."  ,$.0,"$-0",+$000
'
'   �ȉ��̗v�f��Currency�^�ɉ��߂���Ȃ��B
'   "1.234",  "$20","$$50","1,000.00001","$-0,000"
'
'   Currency�^���\���ł���͈́B��������4���܂ŁB
'   -922,337,203,685,477.5808 �` 922,337,203,685,477.5807
'
' ���̑��̎d�l
'
'   ���p�X�y�[�X�͗v�f�̈ꕔ�Ƃ݂Ȃ��B
'   �l�����݂��Ȃ��v�f��Empty�l��}������B
'   �񐔂������Ă��Ȃ��ꍇ�A����Ȃ��v�f��Empty�l�ŕ⊮����B
'   ���s�R�[�h��`\r\n`, `\r`, `\n`���g�p�ł���B�������A���݂���ꍇ�͕s���ȓ����
'   ����ꍇ������B�܂��A�t�@�C�������̋�̉��s�͖��������B
'
'   �����Z�b�g�p�����[�^�[�ɂ��ẮAADODB�̎d�l�ɏ�����B
'   Charset �v���p�e�B (ADO)
'   https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/charset-property-ado


' CSV�t�@�C������ǂݍ��񂾕����̔z����܂ލ\���́B
' �����͂̍ۂɐ擪����1�������擾���Ă����B
Private FileContent As String

' �����͒���FileContent�̓ǂݎ��ʒu�������B
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
    ' CSV�t�@�C����ǂݍ��݁A2�����z��ɕϊ�����B
    '
    ' ����
    '   FilePath: �ǂݍ��ރt�@�C���̐�΃p�X�B
    '   SkipLinesTimes: �擪����ǂݔ�΂��s���B

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, Charset, SkipLinesTimes)

    ParseCsv = To2dArray(Tokens)
End Function


Public Function ParseCsvAsString(ByVal FilePath As String, _
    Optional ByVal Charset As String = "UTF-8", _
    Optional ByVal SkipLinesTimes As Long) As String()
    ' CSV�t�@�C����ǂݍ��݁A������^��2�����z��ɕϊ�����B
    '
    ' ����
    '   FilePath: �ǂݍ��ރt�@�C���̐�΃p�X�B
    '   SkipLinesTimes: �擪����ǂݔ�΂��s���B

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, Charset, SkipLinesTimes)

    ParseCsvAsString = To2dStringArray(Tokens)
End Function


' ________ CSV�t�@�C������g�[�N���W�����쐬���邽�߂̊֐��Q ________

Private Function Tokenize(ByVal FilePath As String, _
    ByVal Charset As String, ByVal SkipLinesTimes As Long) As Collection
    ' FileContent��1�������ǂݎ��A�g�[�N��������B

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
    ' �e�L�X�g�t�@�C�����J���A�����o�[�ϐ�FileContent�Ƀ��[�h����B

    With New ADODB.Stream
        .Charset = Charset
        Call .Open
        Call .LoadFromFile(FilePath)

        FileContent = .ReadText()
        Call .Close
    End With
End Function


Private Function SkipLines(ByVal Times As Long)
    ' �����o�[�ϐ�FileContent����w�肳�ꂽ�s������ǂݔ�΂��B

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
    ' FileContent�̓ǂݎ��ʒu���I�[�ɒB���Ă����True��Ԃ��B
    EndOfFile = Len(FileContent) = FileCursor
End Function


Private Function NextCharInto(ByRef Char As String) As Boolean
    ' FileContent���玟�̕�����ǂݍ��ށB
    '
    ' ����
    '   Char:
    '       �ǂݎ����������������B
    '       �ǂݎ��Ȃ������ꍇ�͋󕶎����������B
    '
    ' �߂�l
    '   ���̕������ǂݎ����True��Ԃ��B

    If EndOfFile() Then
        Char = ""
    Else
        FileCursor = FileCursor + 1
        Char = Mid(FileContent, FileCursor, 1)
        NextCharInto = True
    End If
End Function


Private Function RewindCursor(Optional Times As Long = 1)
    ' FileContent�̓ǂݎ��ʒu���w�肳�ꂽ�������߂��B
    FileCursor = FileCursor - Times
End Function


Private Function SkipIfNextCharIs(ByVal Char As String)
    ' FileContent�̎��̕������w�肳�ꂽ�����Ȃ�ǂݔ�΂��B

    Dim Temp As String
    If NextCharInto(Temp) Then
        If Temp <> Char Then Call RewindCursor
    End If
End Function


Private Function TokenInto(ByRef Tokens As Collection)
    ' ���݂̓ǂݎ��ʒu����v�f�����o���B

    Dim Token As String
    Dim Char As String

    ' ��؂蕶�����o������܂Ńg�[�N���ɕ�����ǉ�����B
    Do While NextCharInto(Char)
        ' ��؂蕶���ɊY������΃g�[�N�����m�肷��B
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
    ' ���݂̓ǂݎ��ʒu����v�f�����o���B

    ' FileContent�̏I�[�ɒB���Ă����`"`���P�̂Ŏg�p���ꂽ��
    ' ���f���A`"`���̂��̂��g�[�N�������ď������I������B
    If EndOfFile() Then
        Call Tokens(Tokens.Count).Add("""")
        Exit Function
    End If

    Dim Token As String
    Dim Char As String

    Do While NextCharInto(Char)
        If Char = """" Then
            ' �t�@�C���̏I�[�̂��߁A�g�[�N�����m�肷��B
            If Not NextCharInto(Char) Then
                Exit Do
            End If

            Select Case Char
                Case ","
                    ' ��؂蕶���̂��߁A�g�[�N�����m�肷��B
                    Exit Do
                Case vbCr
                    ' ��؂蕶���̂��߁A�g�[�N�����m�肷��B
                    Call Tokens(Tokens.Count).Add(Token)
                    Call Tokens.Add(New Collection)
                    Call SkipIfNextCharIs(vbLf)
                    Exit Function
                Case vbLf
                    ' ��؂蕶���̂��߁A�g�[�N�����m�肷��B
                    Call Tokens(Tokens.Count).Add(Token)
                    Call Tokens.Add(New Collection)
                    Exit Function
                Case """"
                    ' �G�X�P�[�v�����Ƃ݂Ȃ��A��������B
                Case Else
                    ' �G�X�P�[�v�Y��Ƃ݂Ȃ��A�v�f�ɒǉ�����B
                    Char = """" & Char
            End Select
        End If

        Token = Token + Char
    Loop

    Call Tokens(Tokens.Count).Add(Token)
End Function


Private Function RemoveEmptyLinesAtEnd(ByRef Tokens As Collection)
    ' �g�[�N���W�������̕s�v�ȉ��s���폜����B

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


' ________ �g�[�N���W������2�����z����쐬���邽�߂̊֐��Q ________

Private Function To2dArray(ByRef Tokens As Collection) As Variant()
    ' �g�[�N���W����2�����z��ɕϊ�����B

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
                        Stop ' �������B
                End Select

                TokensArray(RowIndex - 1, ColumnIndex - 1) = Token
            End If
        Next
    Next

    To2dArray = TokensArray
End Function


Private Function To2dStringArray(ByRef Tokens As Collection) As String()
    ' �g�[�N���W���𕶎���^��2�����z��ɕϊ�����B

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
    ' �����񂩂�g�[�N���̎�ނ𔻒肷��B

    ' �������B�e�X�g�p�ɂ��ׂĕ�����Ɣ��肷��B
    KindOf = StringToken
End Function


Private Function CountColumnLength(ByRef Tokens As Collection) As Long
    ' �g�[�N���W���̗񐔂����߂�B

    Dim Column As Collection
    For Each Column In Tokens
        If CountColumnLength < Column.Count Then
            CountColumnLength = Column.Count
        End If
    Next
End Function