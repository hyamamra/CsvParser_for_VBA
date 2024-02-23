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
    ' CSV�t�@�C����ǂݍ��݁A�񎟌��z��ɕϊ�����B
    '
    ' ����
    '   FilePath: �ǂݍ��ރt�@�C���̐�΃p�X�B
    '   AsString: ���ׂĂ̗v�f�𕶎���Ƃ��ēǂݍ��ށB
    '   SkipLines: �擪����ǂݔ�΂��s���B

    Dim Tokens As Tokens
    Tokens = Tokenize(FilePath, SkipLines)

    Set ReadCsv = Tokens.Values
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


Private Function Tokenize(ByVal FilePath As String, ByVal SkipLines As Long) As Tokens
    ' CSV�t�@�C����1�������ǂݎ��A�g�[�N��������B
    ' �g�[�N���ɂ͐��l�A������A�J���}�A���s�A��v�f�����݂���B

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
                ' �ǂݔ�΂��B

            Case " "
                ' �ǂݔ�΂��B

            Case """"
                Call TakeStringToken(Into:=Tokens)

            Case Else
                If Char Like "[-+.0-9]" Then
                    Call RewindChar
                    Call TakeDigitsToken(Into:=Tokens)
                Else
                    Call Err.Raise(513, , _
                        Seek(FileNumber) & " �����ڂɕs���ȕ��� `" _
                        & Char & "` �����o���܂����B")
                End If
        End Select
    Loop

ErrorHandler:
    Call CloseFile
    If Err Then Call Err.Raise(Err)
    Tokenize = Tokens
End Function


Private Function SkipLines(ByVal Times As Long)
    ' �J���Ă���t�@�C������w�肳�ꂽ�s������ǂݔ�΂��B

    Dim T As Long
    For T = 0 To Times
        If Not EOF(FileNumber) Then
            Dim Temp As String
            Line Input #FileNumber, Temp
        End If
    Next
End Function


Private Function TakeCommaToken(ByRef Into As Tokens)
    ' �K�v�ł���΋�g�[�N����ǉ����Ă���J���}�g�[�N����ǉ�����B

    Call AddEmptyTokenIfLastTokenIsDelimiter(Into)
    Call Into.Kinds.Add(TokenKind.CommaToken)
    Call Into.Values.Add(Empty)
End Function


Private Function TakeNewLineToken(ByRef Into As Tokens, ByVal Char As String)
    ' �K�v�ł���΋�g�[�N����ǉ����Ă�����s�g�[�N����ǉ�����B
    ' `\r`�����o�����ꍇ�A���̕�����`\n`�ł����`\r\n`�Ƃ݂Ȃ��A`\n`��ǂݔ�΂��B

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
    ' �Ō�ɐ؂�o�����g�[�N�������s�A�J���}�ł���A�܂���
    ' �؂�o�����g�[�N�����܂��Ȃ���΋�g�[�N����ǉ�����B

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
    ' ���݂̓ǂݎ��ʒu���當��������o���ATokens�Ɋi�[����B

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

    Call Into.Kinds.Add(TokenKind.StringToken)
    Call Into.Values.Add(Value)
End Function


Private Function TakeDigitsToken(ByRef Into As Tokens)
    ' ���݂̓ǂݎ��ʒu���琔�l�����o���ATokens�Ɋi�[����B

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

    Call Into.Kinds.Add(TokenKind.DigitsToken)
    Call Into.Values.Add(Value)
End Function
