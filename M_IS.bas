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
'+ 対象が数字か判定
    
    Dim i           As Long
    Dim Len_Val     As Long
    Dim Flg_Num     As Boolean
    
    '- 数値と見なせる場合
    If IsNumeric(Expression) = True Then
        
        '- フラグ初期化
        Flg_Num = True
        
        '- 長さを取得
        Len_Val = Len(CStr(Expression))
        
        '- 各文字毎に、数字かどうか判定し、
        For i = 1 To Len_Val
            
            If InStr(1, "0123456789", Mid$(Expression, i, 1), vbTextCompare) = 0 Then
                
                '- 数字以外が含まれていた場合、フラグを下げて抜ける
                Flg_Num = False
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    IsNumber = Flg_Num
    
End Function

Public Function IsDecimal(Expression As Variant) As Boolean
'+ 対象が10進数の数値か判定
    
    Dim i           As Long
    Dim Len_Val     As Long
    Dim Flg_Num     As Boolean
    
    '- 数値と見なせる場合
    If IsNumeric(Expression) = True Then
        
        '- フラグ初期化
        Flg_Num = True
        
        '- 長さを取得
        Len_Val = Len(CStr(Expression))
        
        '- 各文字毎に、10進数の文字列かどうか判定
        For i = 1 To Len_Val
            
            If InStr(1, "-.0123456789", Mid$(Expression, i, 1), vbTextCompare) = 0 Then
                
                '- 数字以外が含まれていた場合、フラグを下げて抜ける
                Flg_Num = False
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    IsDecimal = Flg_Num
    
End Function
