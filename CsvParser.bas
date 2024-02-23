Attribute VB_Name = "CsvParser"
Option Explicit


' �T�v
'   CSV��ǂݍ��ރ��W���[���B
'
'   �ȉ��̒l�͓ǂݎ����T�|�[�g���Ă���B
'
'   �E�����A�����A�����iDouble�^�ɕϊ��\�Ȓl�j�B
'   �E���s���܂ޕ�����B
'   �E�G�X�P�[�v�ς݂̃_�u���N�H�[�e�[�V�������܂ޕ�����B
'
'   �񐔂������Ă��Ȃ��ꍇ�A��̗v�f�ŕ⊮����B
'   �l�Ɋ܂܂�Ȃ��󔒕����͖��������B
'   �t�@�C�������̉��s�͖��������B
'   �P�����Z�q`+`��`-`���󂯓����B
'   ���l���̋󔒕��������e���Ȃ��B
'   ���l��Double�^�Ƀp�[�X����邽�߁A���m�Ȑ��l���K�v�ȏꍇ��
'   AsString�I�v�V������True�ɂ��邱�Ƃŕ�����Ƃ��Ď擾�ł���B


Private FileNumber As Long


Private Enum TokenKind
    Str
    Num
    Emp
End Enum


Public Function ReadCsv( _
    ByVal FilePath As String, _
    Optional ByVal NumberOfSkipLines As Long) As Variant()
    ' CSV�t�@�C����ǂݍ��݁A�񎟌��z��ɕϊ�����B
    '
    ' ����
    '   FilePath: �ǂݍ��ރt�@�C���̐�΃p�X�B
    '   NumberOfSkipLines:
    '       �擪����ǂݔ�΂��s���B
    '       �ǂݔ�΂����悪�����񒆂̏ꍇ�A��O����������B

    Dim Tokens As Collection
    Set Tokens = Tokenize(FilePath, NumberOfSkipLines)

    ReadCsv = To2dArray(Tokens)
End Function


Private Function To2dArray(ByRef Tokens As Collection) As Variant()
    ' �g�[�N���W����񎟌��z��ɕϊ�����B

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
    ' �g�[�N���W�������̕s�v�ȉ��s���폜����B

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
    ' �g�[�N���W���̗񐔂����߂�B

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
    ' CSV�t�@�C����1�������ǂݎ��A�g�[�N��������B
    ' �g�[�N���ɂ͐��l�A������A�J���}�A���s�A��v�f�����݂���B

    Dim Tokens As New Collection
    Call Tokens.Add(New Collection)

    Call OpenFileAsReadOnly(FilePath)
    Call SkipLines(NumberOfSkipLines)

    ' �g�[�N���̓ǂݎ���A��؂蕶����
    ' �ҋ@���Ă����Ԃł����True�ƂȂ�B
    Dim WaitingForDelimiter As Boolean

    On Error GoTo ErrorHandler

    Dim Char As String
    Do While NextChar(Into:=Char)
        If (Char = " ") Or (Char = vbTab) Then GoTo Continue

        If WaitingForDelimiter Then
            Select Case Char
                Case ","
                    ' �������Ȃ��B
                Case vbLf
                    Call Tokens.Add(New Collection)
                Case vbCr
                    Call Tokens.Add(New Collection)
                    Call SkipIfNextCharIsLf
                Case Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " �����ڂɕs���ȕ��� `" _
                        & Char & "` �����o���܂����B")
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
                        Seek(FileNumber) & " �����ڂɕs���ȕ��� `" _
                        & Char & "` �����o���܂����B")
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
    ' ���݂̓ǂݎ��ʒu���當��������o���B

    Dim Value As String
    Dim Char As String

    Do
        If Not NextChar(Into:=Char) Then
            Call Err.Raise(513, , _
                "������̓_�u���N�H�[�e�[�V�����ŕ���K�v������܂��B" _
                & vbNewLine & "���o����������: " & Value)
        End If

        ' `"`�����o�����ꍇ�A���̕�����`"`�ł����
        ' ������̕��L���Ƃ݂Ȃ��A�g�[�N�����m�肷��B
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
    ' ���݂̓ǂݎ��ʒu���琔�l�����o���B

    Dim Value As String

    Dim Char As String
    Call NextChar(Into:=Char)
    Value = Char

    ' �����A�����_�ȊO�����o������؂�o�����I������B
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
            "�s���Ȑ��l�����o���܂����B" _
            & vbNewLine & "�Œ�ł�1�̐������K�v�ł��B" _
            & vbNewLine & "���o�������l: " & Value)
    ElseIf Value Like "*.*.*" Then
        Call Err.Raise(513, , _
            "�s���Ȑ��l�����o���܂����B" _
            & vbNewLine & "�����_���������݂��܂��B" _
            & vbNewLine & "���o�������l: " & Value)
    End If

    NumberToken = CStr(TokenKind.Num) & Value
End Function


Private Function EmptyToken() As String
    EmptyToken = CStr(TokenKind.Emp)
End Function


Private Function OpenFileAsReadOnly(ByVal PathName As String)
    ' �ǂݎ���p�Ńe�L�X�g�t�@�C�����J���B
    ' ��x�ɑS�Ă�ǂݍ��܂Ȃ����Ƃŏȃ���������}��B

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
    ' �J���Ă���t�@�C������w�肳�ꂽ�s������ǂݔ�΂��B

    Dim T As Long
    For T = 1 To Times
        If Not EOF(FileNumber) Then
            Dim Temp As String
            Line Input #FileNumber, Temp
        End If
    Next
End Function


Private Function NextChar(ByRef Into As String) As Boolean
    ' ���̕�����ǂݍ��ށB
    '
    ' ����
    '   Into:
    '       �ǂݎ����������������B
    '       �ǂݎ��Ȃ��ꍇ�A�󕶎����������B
    '
    ' �߂�l
    '   ���̕������ǂݎ����True��Ԃ��B

    Into = ""
    If Not EOF(FileNumber) Then
        Into = Input(1, #FileNumber)
        NextChar = True
    End If
End Function


Private Function RewindChar()
    ' ���݂̓ǂݎ��ʒu��1�������߂��B
    Seek #FileNumber, Seek(FileNumber) - 1
End Function


Private Function SkipIfNextCharIsLf()
    ' ���̕�����`\n`�Ȃ�ǂݔ�΂��B
    Dim Char As String
    If NextChar(Char) Then
        If Char <> vbLf Then Call RewindChar
    End If
End Function
