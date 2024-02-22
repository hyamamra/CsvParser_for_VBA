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
'   ��؂蕶���̃L�����b�W���^�[���i"\r"�j�͖��������B
'   �l�Ɋ܂܂�Ȃ��󔒕����͖��������B
'   �t�@�C�������̉��s�͖��������B
'   �P�����Z�q��"+"�����"-"���󂯓����B
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


Private Sub OpenFileAsReadOnly(ByVal PathName As String)
    ' �ǂݎ���p�Ńe�L�X�g�t�@�C�����J���B
    ' ��x�ɑS�Ă�ǂݍ��܂Ȃ����Ƃŏȃ���������}��B

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
    ' ���̕�����ǂݍ��ށB
    '
    ' ����
    '   Char:
    '       �ǂݎ����������������B
    '       �ǂݎ��Ȃ��ꍇ�A�󕶎����������B
    '
    ' �߂�l
    '   ���̕������ǂݎ����True��Ԃ��B

    Char = ""
    If Not EOF(FileNumber) Then
        Char = Input(1, #FileNumber)
        NextChar = True
    End If
End Function


Private Sub RewindChar()
    ' ���݂̓ǂݎ��ʒu��1�������߂��B
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
                        Seek(FileNumber) & " �����ڂɕs���ȕ��� """ _
                        & Char & """ �����o���܂����B")
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
                "������̓_�u���N�H�[�e�[�V�����ŕ���K�v������܂��B" _
                & vbNewLine & "���o����������: " & Value)
        End If

        ' `"`�����o�����ꍇ�A���̕�����`"`�ł����
        ' ������̕��L���Ƃ݂Ȃ��A�g�[�N�����m�肷��B
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
            "�s���Ȑ��l�����o���܂����B" _
            & vbNewLine & "�Œ�ł�1�̐������K�v�ł��B" _
            & vbNewLine & "���o�������l: " & Value)
    ElseIf Value Like "*.*.*" Then
        Call Err.Raise(513, , _
            "�s���Ȑ��l�����o���܂����B" _
            & vbNewLine & "�����_���������݂��܂��B" _
            & vbNewLine & "���o�������l: " & Value)
    End If

    Call Tokens.Kinds.Add(TokenKind.DigitsToken)
    Call Tokens.Values.Add(Value)
End Function
