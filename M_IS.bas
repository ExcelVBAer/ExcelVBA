Attribute VB_Name = "M_IS"
Option Explicit

Public Function IsNothing(Expression As Object) As Boolean
    
    IsNothing = (Expression Is Nothing)
    
End Function

Public Function IsAry(Expression As Variant) As Boolean
    
    Dim Len_Ary     As Long
    
    If IsArray(Expression) = True Then
        
        On Error Resume Next
        Len_Ary = UBound(Expression, 1) - LBound(Expression, 1) + 1
        On Error GoTo 0
        
        If 0 < Len_Ary Then
            IsAry = True
        End If
        
    End If
    
End Function

Public Function IsCollection(Expression As Variant) As Boolean
    
    If IsObject(Expression) = True Then
        
        If TypeName(Expression) = "Collection" Then
            
            IsCollection = True
            
        End If
        
    End If
        
End Function

Public Function IsDictionary(Expression As Variant) As Boolean
    
    If IsObject(Expression) = True Then
        
        If TypeName(Expression) = "Dictionary" Then
            
            IsDictionary = True
            
        End If
        
    End If
    
End Function

Public Function IsNumber(Expression As Variant) As Boolean
'+ �Ώۂ�����������
    
    Dim i           As Long
    Dim Len_Val     As Long
    Dim Flg_Num     As Boolean
    
    '- ���l�ƌ��Ȃ���ꍇ
    If IsNumeric(Expression) = True Then
        
        '- �t���O������
        Flg_Num = True
        
        '- �������擾
        Len_Val = Len(CStr(Expression))
        
        '- �e�������ɁA�������ǂ������肵�A
        For i = 1 To Len_Val
            
            If InStr(1, "0123456789", Mid$(Expression, i, 1), vbTextCompare) = 0 Then
                
                '- �����ȊO���܂܂�Ă����ꍇ�A�t���O�������Ĕ�����
                Flg_Num = False
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    IsNumber = Flg_Num
    
End Function

Public Function IsDecimal(Expression As Variant) As Boolean
'+ �Ώۂ�10�i���̐��l������
    
    Dim i           As Long
    Dim Len_Val     As Long
    Dim Flg_Num     As Boolean
    
    '- ���l�ƌ��Ȃ���ꍇ
    If IsNumeric(Expression) = True Then
        
        '- �t���O������
        Flg_Num = True
        
        '- �������擾
        Len_Val = Len(CStr(Expression))
        
        '- �e�������ɁA10�i���̕����񂩂ǂ�������
        For i = 1 To Len_Val
            
            If InStr(1, "-.0123456789", Mid$(Expression, i, 1), vbTextCompare) = 0 Then
                
                '- �����ȊO���܂܂�Ă����ꍇ�A�t���O�������Ĕ�����
                Flg_Num = False
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    IsDecimal = Flg_Num
    
End Function
