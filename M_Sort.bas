Attribute VB_Name = "M_Sort"
Option Explicit

Declare Function StrCmpLogicalW Lib "SHLWAPI.DLL" (ByVal lpStr1 As String, ByVal lpStr2 As String) As Long

Public Function SortByIntuitiveFileName(ByRef PathAry As Variant)
'+ Sort As Explorer

    Dim i           As Long
    Dim j           As Long
    Dim TempPath    As String
    Dim ConvAry()   As String
    
    If IsArray(PathAry) = False Then Exit Function
    
    ReDim ConvAry(LBound(PathAry, 1) To UBound(PathAry, 1))
    
    For i = LBound(PathAry, 1) To UBound(PathAry, 1)
        ConvAry(i) = StrConv(CStr(PathAry(i)), vbUnicode)
    Next
    
    For i = LBound(PathAry) To UBound(PathAry)
        
        For j = i To UBound(PathAry)
            
            If StrCmpLogicalW(ConvAry(i), ConvAry(j)) > 0 Then
                
                TempPath = PathAry(i)
                PathAry(i) = PathAry(j)
                PathAry(j) = TempPath
                
            End If
            
        Next
        
    Next

End Function

Public Function Sort_Quick(SortAry As Variant) As Variant
'- QuickSort
    
    Dim At_Min          As Long
    Dim At_Max          As Long
    Dim i               As Long
    Dim Flg_IsNumeric   As Boolean
    
    At_Min = LBound(SortAry, 1)
    At_Max = UBound(SortAry, 1)
    
    '- �Ώۂ̔z�񂪐��l������
    Flg_IsNumeric = True
    For i = At_Min To At_Max
        If IsNumeric(SortAry(i)) = False Then
            Flg_IsNumeric = False
            Exit For
        End If
    Next
    
    '- �N�C�b�N�\�[�g���s���ʂ�z��Ɋi�[
    If Flg_IsNumeric = True Then
        '+ ���l��p
        Call QuickSort_Num(SortAry, At_Min, At_Max)
    Else
        '+ �S�l
        Call QuickSort_Variant(SortAry, At_Min, At_Max)
    End If
    
End Function

Private Function QuickSort_Num(SortAry As Variant, ByVal At_Min As Long, ByVal At_Max As Long)
'+ �z��͈ꎟ��
'+  ��A����

    Dim At_Mid          As Long
    Dim Val_Mid         As Double
    Dim At_Next         As Long
    Dim At_Temp         As Long
    Dim Val_Temp        As Double
    
    '- ���̒l���E�̒l�𒴂����ꍇ�I��
    If At_Min >= At_Max Then Exit Function
    
    '- �����̈ʒu���Z�o���ăs�{�b�g��ݒ�
    At_Mid = (At_Max + At_Min) \ 2
    
    Val_Mid = SortAry(At_Mid)
    
    '- �J�n�ʒu�v�f�𒆉��ɃZ�b�g
    SortAry(At_Mid) = SortAry(At_Min)
    
    At_Temp = At_Min
    
    At_Next = At_Min + 1
    Do While At_Next <= At_Max
        
        If SortAry(At_Next) < Val_Mid Then
            
            '- �l�� ��菬�����ꍇ�A��l�����ւ���
            At_Temp = At_Temp + 1
            Val_Temp = SortAry(At_Temp)
            SortAry(At_Temp) = SortAry(At_Next)
            SortAry(At_Next) = Val_Temp
            
        End If
        At_Next = At_Next + 1
    Loop
    
    '- ����ւ����l���J�n�ʒu�ɃZ�b�g���A
    '- �ʂ̕ϐ��Ɋm�ۂ��Ă����l���Ō�ɓ���ւ����ʒu�ɃZ�b�g
    SortAry(At_Min) = SortAry(At_Temp)
    SortAry(At_Temp) = Val_Mid
    
    '---------------------------------------------------------------------------
    
    ' �����O�����ċA�Ăяo����SORT
    Call QuickSort_Num(SortAry, At_Min, At_Temp - 1)
    '---------------------------------------------------------------------------
    ' �����㔼���ċA�Ăяo����SORT
    Call QuickSort_Num(SortAry, At_Temp + 1, At_Max)
    
End Function

Private Function QuickSort_Variant(SortAry As Variant, ByVal At_Min As Long, ByVal At_Max As Long)
'+ �z��͈ꎟ��
'+  ��A����

    Dim At_Mid          As Long
    Dim Val_Mid         As Variant
    Dim At_Next         As Long
    Dim At_Temp         As Long
    Dim Val_Temp        As Variant
    
    '- ���̒l���E�̒l�𒴂����ꍇ�I��
    If At_Min >= At_Max Then Exit Function
    
    '- �����̈ʒu���Z�o���ăs�{�b�g��ݒ�
    At_Mid = (At_Max + At_Min) \ 2
    
    Val_Mid = SortAry(At_Mid)
    
    '- �J�n�ʒu�v�f�𒆉��ɃZ�b�g
    SortAry(At_Mid) = SortAry(At_Min)
    
    At_Temp = At_Min
    
    At_Next = At_Min + 1
    Do While At_Next <= At_Max
        
        If CStr(SortAry(At_Next)) < CStr(Val_Mid) Then
            
            '- �l����菬�����ꍇ�A��l�����ւ���
            At_Temp = At_Temp + 1
            Val_Temp = SortAry(At_Temp)
            SortAry(At_Temp) = SortAry(At_Next)
            SortAry(At_Next) = Val_Temp
            
        End If
        At_Next = At_Next + 1
    Loop
    
    '- ����ւ����l���J�n�ʒu�ɃZ�b�g���A
    '- �ʂ̕ϐ��Ɋm�ۂ��Ă����l���Ō�ɓ���ւ����ʒu�ɃZ�b�g
    SortAry(At_Min) = SortAry(At_Temp)
    SortAry(At_Temp) = Val_Mid
    
    '---------------------------------------------------------------------------
    
    ' �����O�����ċA�Ăяo����SORT
    Call QuickSort_Variant(SortAry, At_Min, At_Temp - 1)
    '---------------------------------------------------------------------------
    ' �����㔼���ċA�Ăяo����SORT
    Call QuickSort_Variant(SortAry, At_Temp + 1, At_Max)
    
End Function

Public Function Sort_Bubble(ArrayValue As Variant, _
                            Optional CaseNo As Long = 1, _
                            Optional Target_Col As Long = 0) As Variant
    '�o�u���\�[�g�����s
    
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim N                   As Long
    Dim S_Row               As Long
    Dim E_Row               As Long
    Dim TempValue           As Variant
    
    S_Row = LBound(ArrayValue, 1)
    E_Row = UBound(ArrayValue, 1)
    
    '�s�̒���
    If Target_Col = 0 Then
        Target_Col = LBound(ArrayValue, 2)
    End If
    
    For k = E_Row - 1 To S_Row Step -1
        For i = S_Row To k
            j = i + 1
                        
            '����(CaseNo=1)
            If CaseNo = 1 Then
                If ArrayValue(i, Target_Col) > ArrayValue(j, Target_Col) Then
                    
                    For N = LBound(ArrayValue, 2) To UBound(ArrayValue, 2)
                        TempValue = ArrayValue(i, N)
                        ArrayValue(i, N) = ArrayValue(j, N)
                        ArrayValue(j, N) = TempValue
                    Next
                    
                End If
            '�~��
            Else
                If ArrayValue(i, Target_Col) < ArrayValue(j, Target_Col) Then
                
                    For N = LBound(ArrayValue, 2) To UBound(ArrayValue, 2)
                        TempValue = ArrayValue(i, N)
                        ArrayValue(i, N) = ArrayValue(j, N)
                        ArrayValue(j, N) = TempValue
                    Next
                    
                End If
            End If
        Next
    Next
    
    Sort_Bubble = ArrayValue
    
End Function

Private Function QuicktSort_Str_Down(StrAry() As String, At_Min As Long, At_Max As Long)
    
    Dim TempStr     As String
    Dim i           As Long
    Dim j           As Long
    
    '���[�ƉE�[����v���Ă����Ȃ�v���V�[�W���𔲂���
    If At_Min >= At_Max Then Exit Function
    
    i = At_Min + 1
    j = At_Max
    
    Do While i <= j
    
        Do While i <= j
            If StrComp(StrAry(i), StrAry(At_Min), 1) = -1 Then
                Exit Do
            End If
            i = i + 1
        Loop
        
        Do While i <= j
            If StrComp(StrAry(i), StrAry(At_Min), 1) = 1 Then
                Exit Do
            End If
            j = j - 1
        Loop
        
        If i >= j Then Exit Do
            
        TempStr = StrAry(j)
        StrAry(j) = StrAry(i)
        StrAry(i) = TempStr
            
        i = i + 1
        j = j - 1
        
    Loop
    
    TempStr = StrAry(j)
    StrAry(j) = StrAry(At_Min)
    StrAry(At_Min) = TempStr
    
    Call QuicktSort_Str_Down(StrAry, At_Min, j - 1)
    Call QuicktSort_Str_Down(StrAry, j + 1, At_Max)
    
End Function

Private Function QuickSort_Str_Up(StrAry() As String, At_Min As Long, At_Max As Long)
    
    Dim TempStr     As String
    Dim i           As Long
    Dim j           As Long
    
    '���[�ƉE�[����v���Ă����Ȃ�v���V�[�W���𔲂���
    If At_Min >= At_Max Then Exit Function
    
    i = At_Min + 1
    j = At_Max
    
    Do While i <= j
        
        Do While i <= j
            If StrComp(StrAry(i), StrAry(At_Min), 1) = 1 Then
                Exit Do
            End If
            i = i + 1
        Loop
        
        Do While i <= j
            If StrComp(StrAry(j), StrAry(At_Min), 1) = -1 Then
                Exit Do
            End If
            j = j - 1
        Loop
        
        If i >= j Then Exit Do
                
        TempStr = StrAry(j)
        StrAry(j) = StrAry(i)
        StrAry(i) = TempStr
            
        i = i + 1
        j = j - 1
        
    Loop
    
    TempStr = StrAry(j)
    StrAry(j) = StrAry(At_Min)
    StrAry(At_Min) = TempStr
    
    Call QuickSort_Str_Up(StrAry, At_Min, j - 1)
    Call QuickSort_Str_Up(StrAry, j + 1, At_Max)
    
End Function


