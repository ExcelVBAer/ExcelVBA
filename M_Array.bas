Attribute VB_Name = "M_Array"
Option Explicit

Public Function Array_DimCount(DataAry As Variant) As Long
'+ 配列の次元数を取得

    Dim T_Base      As Long
    Dim T_Dim       As Long
    
    On Error GoTo Terminate
    For T_Dim = 1 To 100
        T_Base = LBound(DataAry, T_Dim)
    Next

Terminate:
    On Error GoTo 0
    
    Array_DimCount = T_Dim - 1
    
End Function

Public Function Array_Lbound_Ubound(DataAry As Variant, Optional Row_L As Long, Optional Row_U As Long, Optional Col_L As Long, Optional Col_U As Long)
    
    Dim T_Dim       As Long
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        Row_L = LBound(DataAry, 1)
        Row_U = UBound(DataAry, 1)
        
    Case 2
        Row_L = LBound(DataAry, 1)
        Row_U = UBound(DataAry, 1)
        Col_L = LBound(DataAry, 2)
        Col_U = UBound(DataAry, 2)
        
    End Select
    
End Function

Public Function Array_Paste(T_Range As Range, DataAry As Variant)
'+ T_Rangeを基点に配列を貼り付ける
    
    Dim Row_Paste       As Long
    Dim Col_Paste       As Long
    Dim T_Dim     As Long
    
    '- 配列以外は終了
    If IsArray(DataAry) = False Then Exit Function
    
    '- 配列の次元数を取得
    T_Dim = Array_DimCount(DataAry)
    
    '- 配列の行・列数を取得
    Select Case T_Dim
        
        Case 1
            
            Row_Paste = 1
            Col_Paste = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
        
        Case 2
            
            Row_Paste = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
            Col_Paste = UBound(DataAry, 2) - LBound(DataAry, 2) + 1
        
    End Select
    
    '配列の貼付け
    T_Range.Resize(Row_Paste, Col_Paste).Value = DataAry
    
End Function

Public Function Array_RandomArray(Dimention As Long, _
                                  Row_End As Long, _
                                  Optional Col_End As Long = 0, _
                                  Optional NotRandom As Boolean = False, _
                                  Optional Base0 As Boolean = True, _
                                  Optional RandomSize As Long = 1000) As Variant
    
    Dim T_Ary           As Variant
    Dim i               As Long
    Dim j               As Long
    Dim Base_Num        As Long
    Dim T_Size          As Long
    
    '配列の始まる場所を決定
    If Base0 = False Then
        Base_Num = 1
    Else
        Base_Num = 0
    End If
    
    'ランダム数の桁数を決定
    T_Size = RandomSize
    If T_Size <> 1000 Then
        T_Size = 10 ^ CLng((Len(CStr(T_Size)) - 1))
    End If
    
    Select Case Dimention
    
    Case 1
        
        If Row_End < Base_Num Then Exit Function
        
        'ランダム数を格納する配列を設定
        ReDim T_Ary(Base_Num To Row_End)
        
        'ランダム数を配列に格納
        If NotRandom = False Then
            Call Randomize
            For i = Base_Num To Row_End
                T_Ary(i) = Int(T_Size * Rnd + 1)
            Next
        Else
            For i = Base_Num To Row_End
                T_Ary(i) = i
            Next
        End If
        
    Case 2
        
        If Row_End < Base_Num Then Exit Function
        If Col_End < Base_Num Then Exit Function
        
        'ランダム数を格納する配列を設定
        ReDim T_Ary(Base_Num To Row_End, Base_Num To Col_End)
        
        If NotRandom = False Then
            
            'ランダム数を配列に格納
            Call Randomize
            For i = Base_Num To Row_End
                For j = Base_Num To Col_End
                    T_Ary(i, j) = Int(T_Size * Rnd + 1)
                Next
            Next
        
        Else
            
            For i = Base_Num To Row_End
                For j = Base_Num To Col_End
                    T_Ary(i, j) = CStr(CStr(i) & "," & CStr(j))
                Next
            Next
            
        End If
        
    End Select
    
    Array_RandomArray = T_Ary
    
End Function

Public Function Array_Redim(RedimAry As Variant, BaseAry As Variant, Optional T_Dim As Long = 0)
    'ByRef
    
    If T_Dim = 0 Then
        T_Dim = Array_DimCount(RedimAry)
    End If
    
    Select Case T_Dim
    
    Case 1
        ReDim RedimAry(LBound(BaseAry, 1) To UBound(BaseAry, 1))
        
    Case 2
        ReDim RedimAry(LBound(BaseAry, 1) To UBound(BaseAry, 1), LBound(BaseAry, 2) To UBound(BaseAry, 2))
        
    End Select
    
End Function

Public Function Array_ReDimPreserve(RedimAry As Variant, Optional AddSpace As Long = 1, Optional T_Dim As Long = 0)
    'ByRef
    
    '- 次元数の指定が無ければ、取得する
    If T_Dim = 0 Then
        T_Dim = Array_DimCount(RedimAry)
    End If
    
    '- １次元・２次元で、枠を拡張する
    Select Case T_Dim
    
    Case 1
        ReDim Preserve RedimAry(LBound(RedimAry, 1) To UBound(RedimAry, 1) + AddSpace)
        
    Case 2
        ReDim Preserve RedimAry(LBound(RedimAry, 1) To UBound(RedimAry, 1), LBound(RedimAry, 2) To UBound(RedimAry, 2) + AddSpace)
        
    End Select
    
End Function

Private Function Array_Trans(DataAry As Variant, Optional ByRefReturn As Boolean = False) As Variant
'+ 配列を逆転する(1次配列(横)の場合は、2次配列(縦)にする)
'- AppのTransposeでは,行or列方向の要素数が65536を超えた場合「型が一致しない」とエラーになる。
'+ AppのTransposeとの処理速度は、ほとんど変わらない微差
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        ReDim RetAry(Row_L To Row_U, Row_L To Row_L)
        
        For i = Row_L To Row_U
            RetAry(i, Row_L) = DataAry(i)
        Next
        
    Case 2
        ReDim RetAry(Col_L To Col_U, Row_L To Row_U)
        
        For i = Row_L To Row_U
            For j = Col_L To Col_U
                RetAry(j, i) = DataAry(i, j)
            Next
        Next
        
    End Select
    
    If ByRefReturn = False Then
        Array_Trans = RetAry
    Else
        '+ データ配列が大き過ぎる場合は、引数を一旦解放して格納し直す
        DataAry = Empty
        DataAry = RetAry
    End If
    
End Function

Public Function Array_Replace(DataAry As Variant, FindVal As Variant, ReplaceVal As Variant)
    
    Dim i       As Long
    Dim j       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim T_Dim   As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            If DataAry(i) = FindVal Then
                DataAry(i) = ReplaceVal
            End If
        Next
        
    Case 2
        
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                
                If DataAry(i, j) = FindVal Then
                    DataAry(i, j) = ReplaceVal
                End If
                
            Next
        Next
        
    End Select
    
End Function

Public Function Array_BaseN(DataAry As Variant, BaseN As Long) As Variant
    
    Dim T_Dim       As Long
    Dim i           As Long
    Dim j           As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim RetAry      As Variant
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        ReDim RetAry(BaseN To BaseN + Row_U - Row_L)
        
        For i = Row_L To Row_U
            RetAry(BaseN + i - Row_L) = DataAry(i)
        Next
        
    Case 2
        
        ReDim RetAry(BaseN To BaseN + Row_U - Row_L, BaseN To BaseN + Col_U - Col_L)
        
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                RetAry(BaseN + i - Row_L, BaseN + j - Col_L) = DataAry(i, j)
            Next
        Next
        
    End Select
    
    Array_BaseN = RetAry
    
End Function

Public Function Array_Merge(Array1 As Variant, Array2 As Variant, Optional Direction As XlRowCol = XlRowCol.xlRows) As Variant
    
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim h           As Long
    Dim Row_L1      As Long
    Dim Row_U1      As Long
    Dim Col_L1      As Long
    Dim Col_U1      As Long
    Dim Row_L2      As Long
    Dim Row_U2      As Long
    Dim Col_L2      As Long
    Dim Col_U2      As Long
    Dim Row_S       As Long
    Dim Row_E       As Long
    Dim Col_S       As Long
    Dim Col_E       As Long
    Dim Row_Diff    As Long
    Dim Dim_Ary1    As Long
    Dim Dim_Ary2    As Long
    Dim MergeAry    As Variant
    
    If IsAry(Array1) = False Then
        MergeAry = Array2
        
    ElseIf IsAry(Array2) = False Then
        MergeAry = Array1
    
    Else
        
        Call Array_Lbound_Ubound(Array1, Row_L1, Row_U1, Col_L1, Col_U1)
        Call Array_Lbound_Ubound(Array2, Row_L2, Row_U2, Col_L2, Col_U2)
        
        Dim_Ary1 = Array_DimCount(Array1)
        Dim_Ary2 = Array_DimCount(Array2)
        
        If Dim_Ary1 <> Dim_Ary2 Then
            Call MsgBox("結合する配列の次元が一致していません", vbCritical)
            Exit Function
        End If
        
        Select Case Dim_Ary1
        
        Case 1
            
            If Direction = xlRows Then
                
                MergeAry = Array1
                
                Row_S = Row_L1
                Row_E = Row_U1 + (Row_U2 - Row_L2 + 1)
                
                ReDim Preserve MergeAry(Row_S To Row_E)
                
                Row_S = Row_U1 + 1
                k = Row_L2
                For i = Row_S To Row_E
                    MergeAry(i) = Array2(k)
                    k = k + 1
                Next
                
            Else
                
                Row_S = Row_L1
                
                If (Row_U1 - Row_L1) < (Row_U2 - Row_L2) Then
                    Row_E = Row_U1 + ((Row_U2 - Row_L2) - (Row_U1 - Row_L1))
                Else
                    Row_E = Row_U1
                End If
                
                ReDim MergeAry(Row_S To Row_E, Row_S To Row_S + 1)
                
                For i = Row_L1 To Row_U1
                    MergeAry(i, Row_S) = Array1(i)
                Next
                
                Row_Diff = Row_L2 - Row_L1
                For i = Row_L2 To Row_U2
                    MergeAry(i - Row_Diff, Row_S + 1) = Array2(i)
                Next
                
            End If
            
        Case 2
            
            
            If Direction = xlRows Then
                    
                Row_S = Row_L1
                Row_E = Row_U1 + (Row_U2 - Row_L2 + 1)
                
                Col_S = Col_L1
                If (Col_U1 - Col_L1) < (Col_U2 - Col_L2) Then
                    Col_E = Col_U1 + ((Col_U2 - Col_L2) - (Col_U1 - Col_L1))
                Else
                    Col_E = Col_U1
                End If
                
                ReDim MergeAry(Row_S To Row_E, Col_S To Col_E)
                
                '+ 高速化の為,Colから回す
                For j = Col_L1 To Col_U1
                    For i = Row_L1 To Row_U1
                        MergeAry(i, j) = Array1(i, j)
                    Next
                Next
                
                '+ 高速化の為,Colから回す
                Row_S = Row_U1 + 1
                Col_S = Col_L1
                Col_E = Col_U2 - (Col_L2 - Col_L1)
                h = Col_L2
                For j = Col_S To Col_E
                    k = Row_L2
                    For i = Row_S To Row_E
                        MergeAry(i, j) = Array2(k, h)
                        k = k + 1
                    Next
                    h = h + 1
                Next
                
            Else
                
                Row_S = Row_L1
                If (Row_U1 - Row_L1) < (Row_U2 - Row_L2) Then
                    Row_E = Row_U1 + ((Row_U2 - Row_L2) - (Row_U1 - Row_L1))
                Else
                    Row_E = Row_U1
                End If
                
                Col_S = Col_L1
                Col_E = Col_U1 + (Col_U2 - Col_L2 + 1)
                
                ReDim MergeAry(Row_S To Row_E, Col_S To Col_E)
                
                '+ 高速化の為,Colから回す
                For j = Col_L1 To Col_U1
                    For i = Row_L1 To Row_U1
                        MergeAry(i, j) = Array1(i, j)
                    Next
                Next
                
                Row_S = Row_L1
                Row_E = Row_U2 - (Row_L2 - Row_L1)
                Col_S = Col_U1 + 1
                
                '+ 高速化の為,Colから回す
                h = Col_L2
                For j = Col_S To Col_E
                    k = Row_L2
                    For i = Row_S To Row_E
                        MergeAry(i, j) = Array2(k, h)
                        k = k + 1
                    Next
                    h = h + 1
                Next
                
            End If
            
        End Select
        
    End If
    
    Array_Merge = MergeAry
    
End Function

Public Function Array_Convert_IsNull_to_NullStr(DataAry As Variant)
    
    Dim i           As Long
    Dim j           As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim T_Dim       As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            
            If IsNull(DataAry(i)) = True Then
                DataAry(i) = ""
            End If
            
        Next
        
    Case 2
            
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                
                If IsNull(DataAry(i, j)) = True Then
                    DataAry(i, j) = ""
                End If
                
            Next
        Next
        
    End Select
    
End Function

Public Function Array_Convert_NullStr_to_Null(DataAry As Variant)
    
    Dim i           As Long
    Dim j           As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim T_Dim       As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            
            If UCase(Trim$(CStr(DataAry(i)))) = "NULL" Then
                DataAry(i) = Null
            End If
            
        Next
        
    Case 2
        
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                
                If UCase(Trim$(CStr(DataAry(i, j)))) = "NULL" Then
                    DataAry(i, j) = Null
                End If
                
            Next
        Next
        
    End Select
    
End Function

Public Function Array_Exist(DataAry As Variant, FindVals As Variant) As Boolean
    
    Dim T_Dim       As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim FindAry     As Variant
    Dim Flg_Exit    As Boolean
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim Fin_L       As Long
    Dim Fin_U       As Long
    
    T_Dim = Array_DimCount(DataAry)
    
    If IsArray(FindVals) = True Then
        FindAry = FindVals
    Else
        FindAry = Array(FindVals)
    End If
    
    Select Case T_Dim
    
    Case 1
        
        Call Array_Lbound_Ubound(DataAry, Row_L, Row_U)
        Call Array_Lbound_Ubound(FindAry, Fin_L, Fin_U)
        
        For i = Row_L To Row_U
            For j = Fin_L To Fin_U
                If DataAry(i) = FindAry(j) Then
                    Flg_Exit = True
                    GoTo Terminate
                End If
            Next
        Next
        
    Case 2
        
        Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
        Call Array_Lbound_Ubound(FindAry, Fin_L, Fin_U)
        
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            
            For i = Row_L To Row_U
                
                For k = Fin_L To Fin_U
                    
                    If DataAry(i, j) = FindAry(k) Then
                        
                        Flg_Exit = True
                        GoTo Terminate
                        
                    End If
                    
                Next
                
            Next
            
        Next
        
    End Select

Terminate:

    Array_Exist = Flg_Exit
    
End Function

Public Function Array_Push(DataAry As Variant, Val As Variant)
    
    If IsArray(DataAry) = True Then
        
        If Array_DimCount(DataAry) <> 1 Then Exit Function
        
        ReDim Preserve DataAry(LBound(DataAry, 1) To UBound(DataAry, 1) + 1)
        DataAry(UBound(DataAry, 1)) = Val
        
    Else
        
        ReDim DataAry(0)
        DataAry(0) = Val
        
    End If
    
End Function

Public Function Array_Pop(DataAry As Variant) As Variant
    
    Dim Pop     As Variant
    Dim Len_Ary As Long
    
    If Array_DimCount(DataAry) <> 1 Then Exit Function
    
    Pop = DataAry(UBound(DataAry, 1))
    
    Len_Ary = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
    
    If Len_Ary = 1 Then
            
        DataAry = Empty
            
    Else
        
        ReDim Preserve DataAry(LBound(DataAry, 1) To UBound(DataAry, 1) - 1)
        
    End If
    
    Array_Pop = Pop
    
End Function

Public Function Array_Length(DataAry As Variant, Optional Direction As XlRowCol = XlRowCol.xlRows) As Long
    
    Dim T_Dim       As Long
    Dim Len_Ary     As Long
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        Len_Ary = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
        
    Case 2
        If Direction = xlRows Then
            Len_Ary = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
        Else
            Len_Ary = UBound(DataAry, 2) - LBound(DataAry, 2) + 1
        End If
        
    End Select
    
    Array_Length = Len_Ary
    
End Function

Public Function Array_Fill(DataAry As Variant, FillVal As Variant)
    
    Dim i       As Long
    Dim j       As Long
    Dim T_Dim   As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        For i = Row_L To Row_U
            DataAry(i) = FillVal
        Next
        
    Case 2
        
        '+ 高速化の為,Colから回す
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                DataAry(i, j) = FillVal
            Next
        Next
        
    End Select
    
End Function

Public Function Array_First(DataAry As Variant) As Variant
    
    Dim T_Dim   As Long
    Dim T_Val   As Variant
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        T_Val = DataAry(LBound(DataAry, 1))
        
    Case 2
        T_Val = DataAry(LBound(DataAry, 1), LBound(DataAry, 2))
        
    End Select
    
    Array_First = T_Val
    
End Function

Public Function Array_Last(DataAry As Variant) As Variant
    
    Dim T_Dim   As Long
    Dim T_Val   As Variant
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        T_Val = DataAry(UBound(DataAry, 1))
        
    Case 2
        T_Val = DataAry(UBound(DataAry, 1), UBound(DataAry, 2))
        
    End Select
    
    Array_Last = T_Val
    
End Function

Public Function Array_Join(DataAry As Variant, Delimiter As String) As String
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim JoinAry     As Variant
    Dim LineAry     As Variant
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        JoinAry = Join(DataAry, Delimiter)
        
    Case 2
        ReDim JoinAry(LBound(DataAry, 1) To UBound(DataAry, 1))
        ReDim LineAry(LBound(DataAry, 2) To UBound(DataAry, 2))
        
        For i = Row_L To Row_U
            For j = Col_L To Col_U
                LineAry(j) = DataAry(i, j)
                JoinAry(i) = Join(LineAry, Delimiter)
            Next
        Next
        
    End Select
    
    Array_Join = JoinAry
    
End Function

Public Function Array_CVar(DataAry As Variant) As Variant
    
    Call PRV_Convert_DataType(DataAry, Array_CVar)
    
End Function

Public Function Array_CStr(DataAry As Variant) As String()
    
    Call PRV_Convert_DataType(DataAry, Array_CStr)
    
End Function

Public Function Array_CInt(DataAry As Variant) As Integer()
    
    Call PRV_Convert_DataType(DataAry, Array_CInt)
    
End Function

Public Function Array_CLng(DataAry As Variant) As Long()
    
    Call PRV_Convert_DataType(DataAry, Array_CLng)
    
End Function

Public Function Array_CSng(DataAry As Variant) As Single()
    
    Call PRV_Convert_DataType(DataAry, Array_CSng)
    
End Function

Public Function Array_CDbl(DataAry As Variant) As Double()
    
    Call PRV_Convert_DataType(DataAry, Array_CDbl)
    
End Function

Public Function Array_CCur(DataAry As Variant) As Currency()
    
    Call PRV_Convert_DataType(DataAry, Array_CCur)
    
End Function

Public Function Array_CDate(DataAry As Variant) As Date()
    
    Call PRV_Convert_DataType(DataAry, Array_CDate)
    
End Function

Public Function Array_CByte(DataAry As Variant) As Byte()
    
    Call PRV_Convert_DataType(DataAry, Array_CByte)
    
End Function

Public Function Array_CBool(DataAry As Variant) As Boolean()
    
    Call PRV_Convert_DataType(DataAry, Array_CBool)
    
End Function

Private Function PRV_Convert_DataType(DataAry As Variant, ConvAry As Variant) As Variant
    
    Dim i           As Long
    Dim j           As Long
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    
'+ 配列の値が型に合わなかった場合、Emptyで返す
On Error GoTo Terminate
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        ReDim ConvAry(Row_L To Row_U)
        For i = Row_L To Row_U
            ConvAry(i) = DataAry(i)
        Next
        
    Case 2
        '+ 高速化の為,Colから回す
        ReDim ConvAry(Row_L To Row_U, Col_L To Col_U)
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                ConvAry(i, j) = DataAry(i, j)
            Next
        Next
        
    End Select
    
On Error GoTo 0
    
    Exit Function
    
Terminate:
    
    ConvAry = Empty
    
End Function

Public Function Array_to_String(DataAry As Variant, Optional Delimiter As String = ",", Optional CrLf As String = vbCrLf) As String
'配列の値を文字列にする
    
    Dim i           As Long
    Dim j           As Long
    Dim LineAry()   As String
    Dim LineStr     As String
    Dim AryStr      As String
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        AryStr = Join(DataAry, vbCrLf)
        
    Case 2
        
        Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
        
        '- 行単位のデータ格納配列を準備
        ReDim LineAry(Col_L To Col_U)
        
        '- 各行で
        For i = Row_L To Row_U
            
            '- 行単位の各値を格納
            For j = Col_L To Col_U
                
                LineAry(j) = CStr(DataAry(i, j))
                
            Next
            
            '- 行単位の文字列を作成
            LineStr = Join(LineAry, Delimiter)
            
            '- 行単位を結合していく
            AryStr = AryStr & CrLf & LineStr
            
        Next
        
        '- 頭を除去
        If AryStr <> "" Then
            AryStr = Mid$(AryStr, Len(CrLf) + 1)
        End If
        
    End Select
                
    '戻り値
    Array_to_String = AryStr
        
End Function

Public Function Array_to_Text(DataAry As Variant, Path_Text As String, Optional Delimiter As String = ",", Optional Write_Line As Boolean = True) As Boolean
'配列の値をCSVデータに置き換える
    
    Dim i           As Long
    Dim j           As Long
    Dim Num_Text    As Long
    Dim LineAry()   As String
    Dim LineStr     As String
    Dim Flg_Write   As Boolean
    
    '配列に値がなかったら
    If IsArray(DataAry) = False Then
        
        Call MsgBox("配列にデータがありませんでした。", vbOKOnly + vbCritical)
        GoTo Terminate
        
    End If
    
    'ファイルを作成
    Num_Text = File_Open_Text(Path_Text, E_OpenMode.Writes, E_ExistError.e_Alart)
    
    'テキストファイル作成に失敗した場合終了
    If Num_Text < 0 Then
        
        Call MsgBox("ファイルの作成に失敗しました。", vbOKOnly + vbCritical)
        GoTo Terminate
        
    End If
    
    '- 行単位で書き込む場合
    If Write_Line = True Then
        
        '- 行単位のデータ格納配列を準備
        ReDim LineAry(LBound(DataAry, 2) To UBound(DataAry, 2))
        
        '- 各行で
        For i = LBound(DataAry, 1) To UBound(DataAry, 1)
            
            '- 行単位の各値を格納
            For j = LBound(DataAry, 2) To UBound(DataAry, 2)
                
                LineAry(j) = CStr(DataAry(i, j))
                
            Next
            
            '- 行単位の文字列を作成
            LineStr = Join(LineAry, Delimiter)
            
            '値の書き出し
            Print #Num_Text, LineStr
                    
        Next
        
    '- 単語毎に書き込む場合
    Else
        
        '- 各行で
        For i = LBound(DataAry, 1) To UBound(DataAry, 1)
            
            '- 行単位の各値を格納
            For j = LBound(DataAry, 2) To UBound(DataAry, 2)
                
                '値の書き出し(;は改行を除去）
                Print #Num_Text, CStr(DataAry(i, j));
                
                'データ区切り(;は改行を除去）
                Print #Num_Text, Delimiter;
                
            Next
            
            '改行(;を付けない空文字で改行）
            Print #Num_Text, ""
                
        Next
        
    End If
    
    'テキストを閉じる
    Close #Num_Text
    
    '書き込みフラグを立てる
    Flg_Write = True
    
Terminate:
    
    '戻り値
    Array_to_Text = Flg_Write
        
End Function

Public Function Array_to_Dictionary(DataAry As Variant) As Scripting.Dictionary
    
    Dim T_Dim       As Long
    Dim Dic_Ary     As Scripting.Dictionary
    Dim i           As Long
    Dim Col_K       As Long
    Dim Col_I       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    
    Set Dic_Ary = New Scripting.Dictionary
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        Row_L = LBound(DataAry, 1)
        Row_U = UBound(DataAry, 1)
        
        For i = Row_L To Row_U
            If Dic_Ary.Exists(DataAry(i)) = False Then
                Call Dic_Ary.Add(DataAry(i), Empty)
            End If
        Next
        
    Case 2
        
        Row_L = LBound(DataAry, 1)
        Row_U = UBound(DataAry, 1)
        Col_L = LBound(DataAry, 2)
        Col_U = UBound(DataAry, 2)
        
        '- 1列目をKey列、2列目をItem列に設定
        Col_K = Col_L
        Col_I = Col_K + 1
        
        '- 1列だけの場合
        If Col_L = Col_U Then
            
            For i = Row_L To Row_U
                If Dic_Ary.Exists(DataAry(i, Col_K)) = False Then
                    Call Dic_Ary.Add(DataAry(i, Col_K), Empty)
                End If
            Next
            
        '- 複数列の場合
        Else
            
            For i = Row_L To Row_U
                If Dic_Ary.Exists(DataAry(i, Col_K)) = False Then
                    Call Dic_Ary.Add(DataAry(i, Col_K), DataAry(i, Col_I))
                End If
            Next
            
        End If
    
    Case Else
        
    End Select
    
    Set Array_to_Dictionary = Dic_Ary
    
    Set Dic_Ary = Nothing
    
End Function

Public Function Array_to_Dictionary2(KeyAry As Variant, Optional ItemAry As Variant = Empty) As Scripting.Dictionary
    
    Dim Dim_Key     As Long
    Dim Dim_Item    As Long
    Dim Dic_Ary     As Scripting.Dictionary
    Dim i_Key       As Long
    Dim i_Item      As Long
    Dim Row_Key_L   As Long
    Dim Row_Key_U   As Long
    Dim Row_Item_L  As Long
    Dim Row_Item_U  As Long
    
    '- 次元数を取得
    Dim_Key = Array_DimCount(KeyAry)
    Dim_Item = Array_DimCount(ItemAry)
    
    Select Case Dim_Key
    
    Case 1
        
        '- KeyとItemから辞書を作成
        Set Dic_Ary = New Scripting.Dictionary
        
        '- 要素情報を取得
        Call Array_Lbound_Ubound(KeyAry, Row_Key_L, Row_Key_U)
        
        '- Item配列が1次元の場合、要素情報を取得しておく
        If Dim_Item = 1 Then
            Call Array_Lbound_Ubound(ItemAry, Row_Item_L, Row_Item_U)
        Else
            Row_Item_U = -1
        End If
        
        '- 開始位置を取得
        i_Item = Row_Item_L
        
        For i_Key = Row_Key_L To Row_Key_U
            
            If Dic_Ary.Exists(KeyAry(i_Key)) = False Then
                
                '- Itemがある場合
                If i_Item <= Row_Item_U Then
                    
                    Call Dic_Ary.Add(KeyAry(i_Key), ItemAry(i_Item))
                    
                Else
                    
                    Call Dic_Ary.Add(KeyAry(i_Key), Empty)
                    
                End If
                
            End If
            
            '- KeyとItemの位置は同期させておく
            i_Item = i_Item + 1
            
        Next
        
    End Select
    
    Set Array_to_Dictionary2 = Dic_Ary
    
    Set Dic_Ary = Nothing
    
End Function

Public Function Array_CountA(DataAry As Variant, Optional NoCountNullString As Boolean = True) As Long
'+ 配列内の値をカウントする
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim Val_Judge   As Variant
    Dim Cnt_Val     As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        For i = Row_L To Row_U
            
            Val_Judge = DataAry(i)
            
            If IsEmpty(Val_Judge) = True Then
                
            ElseIf (CStr(Val_Judge) = "") = NoCountNullString Then
            
            Else
                Cnt_Val = Cnt_Val + 1
            End If
            
        Next
        
    Case 2
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                
            Val_Judge = DataAry(i, j)
            
            If IsEmpty(Val_Judge) = True Then
                
            ElseIf (CStr(Val_Judge) = "") = NoCountNullString Then
            
            Else
                Cnt_Val = Cnt_Val + 1
            End If
                
            Next
        Next
        
    End Select
    
    Array_CountA = Cnt_Val
    
End Function

Public Function Array_Equals(Array1 As Variant, Array2 As Variant) As Boolean
    
    Dim T_Dim1      As Long
    Dim T_Dim2      As Long
    Dim Row_L1      As Long
    Dim Row_U1      As Long
    Dim Col_L1      As Long
    Dim Col_U1      As Long
    Dim Row_L2      As Long
    Dim Row_U2      As Long
    Dim Col_L2      As Long
    Dim Col_U2      As Long
    Dim Adjust_Row  As Long
    Dim Adjust_Col  As Long
    Dim i           As Long
    Dim j           As Long
    Dim Flg_Equal   As Boolean
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    T_Dim1 = Array_DimCount(Array1)
    T_Dim2 = Array_DimCount(Array2)
    
    Call Array_Lbound_Ubound(Array1, Row_L1, Row_U1, Col_L1, Col_U1)
    Call Array_Lbound_Ubound(Array2, Row_L2, Row_U2, Col_L2, Col_U2)
    
    Select Case T_Dim1
    
    Case 1
        
        '- 行要素の調整用
        Adjust_Row = Row_L2 - Row_L1
        
        '- 各値を比較し、一致しなかった場合は抜ける
        For i = Row_L1 To Row_U1
            
            If Array1(i) <> Array2(i + Adjust_Row) Then Exit Function
            
        Next
        
        '- 辿り着いたら、一致フラグを立てる
        Flg_Equal = True
        
    Case 2
        
        '- 要素の調整用
        Adjust_Row = Row_L2 - Row_L1
        Adjust_Col = Col_L2 - Col_L1
        
        '- 各値を比較し、一致しなかった場合は抜ける
        For j = Col_L1 To Col_U1
            
            For i = Row_L1 To Row_U1
                
                If Array1(i, j) <> Array2(i + Adjust_Row, j + Adjust_Col) Then Exit Function
                
            Next
            
        Next
        
        '- 辿り着いたら、一致フラグを立てる
        Flg_Equal = True
        
    End Select
    
    '- 戻り値
    Array_Equals = Flg_Equal
    
End Function

Public Function Array_EqualSize(Array1 As Variant, Array2 As Variant) As Boolean
    
    Dim T_Dim1      As Long
    Dim T_Dim2      As Long
    Dim Row_L1      As Long
    Dim Row_U1      As Long
    Dim Col_L1      As Long
    Dim Col_U1      As Long
    Dim Row_L2      As Long
    Dim Row_U2      As Long
    Dim Col_L2      As Long
    Dim Col_U2      As Long
    Dim Len_1       As Long
    Dim Len_2       As Long
    Dim Flg_Equal   As Boolean
    
    T_Dim1 = Array_DimCount(Array1)
    T_Dim2 = Array_DimCount(Array2)
    
    If T_Dim1 <> T_Dim2 Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L1, Row_U1, Col_L1, Col_U1)
    Call Array_Lbound_Ubound(Array2, Row_L2, Row_U2, Col_L2, Col_U2)
    
    Select Case T_Dim1
    
    Case 1
        
        '- 行数が異なった場合、終了
        Len_1 = Row_U1 - Row_L1 + 1
        Len_2 = Row_U2 - Row_L2 + 1
        If Len_1 <> Len_2 Then Exit Function
        
        '- 辿り着いたら、一致フラグを立てる
        Flg_Equal = True
        
    Case 2
        
        '- 行数が異なった場合、終了
        Len_1 = Row_U1 - Row_L1 + 1
        Len_2 = Row_U2 - Row_L2 + 1
        If Len_1 <> Len_2 Then Exit Function
        
        '- 列数が異なった場合、終了
        Len_1 = Col_U1 - Col_L1 + 1
        Len_2 = Col_U2 - Col_L2 + 1
        If Len_1 <> Len_2 Then Exit Function
        
        '- 辿り着いたら、一致フラグを立てる
        Flg_Equal = True
        
    End Select
    
    Array_EqualSize = Flg_Equal
    
End Function

Public Function Array_Sum(Array1 As Variant, Array2 As Variant) As Variant
'- 加算

    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(Array1)
    
    Call Array_Redim(RetAry, Array1)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            RetAry(i) = Array1(i) + Array2(i)
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                RetAry(i, j) = Array1(i, j) + Array2(i, j)
            Next
        Next
        
    End Select
    
    Array_Sum = RetAry
    
End Function

Public Function Array_Difference(Array1 As Variant, Array2 As Variant) As Variant
'- 減算
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(Array1)
    
    Call Array_Redim(RetAry, Array1)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            RetAry(i) = Array1(i) - Array2(i)
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                RetAry(i, j) = Array1(i, j) - Array2(i, j)
            Next
        Next
        
    End Select
    
    Array_Difference = RetAry
    
End Function

Public Function Array_Product(Array1 As Variant, Array2 As Variant) As Variant
'- 乗算
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(Array1)
    
    Call Array_Redim(RetAry, Array1)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            RetAry(i) = Array1(i) * Array2(i)
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                RetAry(i, j) = Array1(i, j) * Array2(i, j)
            Next
        Next
        
    End Select
    
    Array_Product = RetAry
    
End Function

Public Function Array_Quotient(Array1 As Variant, Array2 As Variant) As Variant
'- 除算
'+ 除数が0の場合,0を返す
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(Array1)
    
    Call Array_Redim(RetAry, Array1)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            If PRV_Is0(Array2(i)) = False Then
                RetAry(i) = Array1(i) / Array2(i)
            Else
                RetAry(i) = 0
            End If
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                If PRV_Is0(Array2(i)) = False Then
                    RetAry(i, j) = Array1(i, j) / Array2(i, j)
                Else
                    RetAry(i, j) = 0
                End If
            Next
        Next
        
    End Select
    
    Array_Quotient = RetAry
    
End Function

Public Function Array_Mod(Array1 As Variant, Array2 As Variant) As Variant
'- 除算の余り
'+ 除数が0の場合,そのままの値を返す
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    If Array_EqualSize(Array1, Array2) = False Then Exit Function
    
    Call Array_Lbound_Ubound(Array1, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(Array1)
    
    Call Array_Redim(RetAry, Array1)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            If PRV_Is0(Array2(i)) = False Then
                RetAry(i) = Array1(i) Mod Array2(i)
            Else
                RetAry(i) = Array1(i)
            End If
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                If PRV_Is0(Array2(i, j)) = False Then
                    RetAry(i, j) = Array1(i, j) Mod Array2(i, j)
                Else
                    RetAry(i, j) = Array1(i, j)
                End If
            Next
        Next
        
    End Select
    
    Array_Mod = RetAry
    
End Function

Private Function PRV_Is0(Val As Variant) As Boolean
    
    If IsNumeric(Val) = True Then
        
        If CLng(Val) = 0 Then
            
            PRV_Is0 = True
            
        End If
        
    End If
    
End Function

Public Function Array_SumIf(DataAry As Variant, Col_Key As Long, key As Variant, Col_Sum As Long) As Double
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim i           As Long
    Dim SumIf       As Double
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 2
        
        '- 各データ行で
        For i = Row_L To Row_U
            
            If DataAry(i, Col_Key) = key Then
                
                If IsNumeric(DataAry(i, Col_Sum)) = True Then
                    
                    SumIf = SumIf + DataAry(i, Col_Sum)
                    
                End If
                
            End If
            
        Next
        
    End Select
    
    Array_SumIf = SumIf
    
End Function

Public Function Array_CountIf(DataAry As Variant, FindVal As Variant) As Long
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim Cnt_Find    As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        For i = Row_L To Row_U
            If DataAry(i) = FindVal Then
                Cnt_Find = Cnt_Find + 1
            End If
        Next
        
    Case 2
        
        '- 各データ行で
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                If DataAry(i, j) = FindVal Then
                    Cnt_Find = Cnt_Find + 1
                End If
            Next
        Next
        
    End Select
    
    Array_CountIf = Cnt_Find
    
End Function

Public Function Array_Split(DataAry As Variant, Optional Delimiter As String = ",") As Variant
    
    Dim i               As Long
    Dim j               As Long
    Dim Cnt_Deli        As Long
    Dim Cnt_Max         As Long
    Dim SplitAry()      As String
    Dim RetAry          As Variant
    Dim T_Dim           As Long
    Dim Row_L           As Long
    Dim Row_U           As Long
    Dim Col_L           As Long
    Dim Col_U           As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        '- 区分文字列の最大数をカウント
        For i = Row_L To Row_U
            Cnt_Deli = String_Count(CStr(DataAry(i)), Delimiter)
            If Cnt_Max < Cnt_Deli Then
                Cnt_Max = Cnt_Deli
            End If
        Next
        
        '- 格納配列を調整
        ReDim RetAry(Row_L To Row_U, 0 To Cnt_Max)
        
        '- Splitして格納していく
        For i = Row_L To Row_U
            
            SplitAry = Split(CStr(DataAry(i)), Delimiter)
            
            For j = LBound(SplitAry, 1) To UBound(SplitAry, 1)
                
                RetAry(i, j) = SplitAry(j)
                
            Next
            
        Next
        
    End Select
    
    Array_Split = RetAry
    
End Function

Public Function Array_Reverse(DataAry As Variant, Optional Direction As XlRowCol = xlRows) As Variant
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim RetAry      As Variant
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        RetAry = DataAry
        
        For i = Row_L To Row_U
            RetAry(Row_U - i + Row_L) = DataAry(i)
        Next
        
    Case 2
        
        RetAry = DataAry
        
        Select Case Direction
        
        Case XlRowCol.xlRows
            
            For j = Col_L To Col_U
                For i = Row_L To Row_U
                    RetAry(Row_U - i + Row_L, j) = DataAry(i, j)
                Next
            Next
            
        Case XlRowCol.xlColumns
            
            For j = Col_L To Col_U
                For i = Row_L To Row_U
                    RetAry(i, Col_U - j + Col_L) = DataAry(i, j)
                Next
            Next
            
        End Select
        
    End Select
    
    Array_Reverse = RetAry
    
End Function

Public Function Array_IsNumeric(DataAry As Variant) As Boolean
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim Flg_Is      As Boolean
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    '- 初期値
    Flg_Is = True
    
    Select Case T_Dim
    
    Case 1
        For i = Row_L To Row_U
            If IsNumeric(DataAry(i)) = False Then
                Flg_Is = False
                Exit For
            End If
        Next
        
    Case 2
        
        For j = Col_L To Col_U
            For i = Row_L To Row_U
                If IsNumeric(DataAry(i, j)) = False Then
                    Flg_Is = False
                    Exit For
                End If
                
            Next
        Next
        
    End Select
    
    Array_IsNumeric = Flg_Is
    
End Function

Public Function Array_Dim1_to_Dim2(DataAry As Variant, Optional Len_Col As Long = 1)
'- 1次元配列を指定列数の2次元配列に変換する
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Cnt_Row     As Long
    Dim Cnt_Col     As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim RetAry      As Variant
    Dim Row_E       As Long
    Dim i1          As Long
    Dim i2          As Long
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        '- 元の行数を取得
        Cnt_Row = Row_U - Row_L + 1
        
        '- 列数を設定
        If 1 <= Len_Col Then
            Cnt_Col = Len_Col
        Else
            Cnt_Col = 1
        End If
        
        '- 2次元に変換後の必要行数を取得
        i1 = Fix(Cnt_Row / Cnt_Col)
        If Cnt_Row Mod Cnt_Col <> 0 Then
            i2 = 1
        End If
        Row_E = i1 + i2
        
        ReDim RetAry(0 To Row_E - 1, 0 To Cnt_Col - 1)
        
        '- 1次元配列の行を、格納配列の行数分スキップしながら回す
        k = Row_L
        For i = Row_L To Row_U Step Cnt_Col
            
            '- 列数分、値を配列に格納
            For j = i To i + Cnt_Col - 1
                
                '- データ最終行まで、データを格納
                If j <= Row_U Then
                
                    RetAry(k, j - i) = DataAry(j)
                
                End If
                
            Next
            k = k + 1
        Next
        
    End Select
    
    Array_Dim1_to_Dim2 = RetAry
    
End Function

Public Function Array_Dim11_to_Dim2(DataAry As Variant)
'+ 1次元x1次元配列を2次元配列に変換する
    
    Dim T_Dim       As Long
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim TempAry     As Variant
    Dim ConvAry     As Variant
    
    Call Array_Lbound_Ubound(DataAry, Row_L, Row_U, Col_L, Col_U)
    
    T_Dim = Array_DimCount(DataAry)
    
    Select Case T_Dim
    
    Case 1
        
        TempAry = DataAry(Row_L)
        Col_L = LBound(TempAry, 1)
        Col_U = UBound(TempAry, 1)
        
        ReDim ConvAry(Row_L To Row_U, Col_L To Col_U)
        
        For i = Row_L To Row_U
            
            TempAry = DataAry(i)
            
            For j = Col_L To Col_U
                
                ConvAry(i, j) = TempAry(j)
                
            Next
            
        Next
        
    End Select
    
    Array_Dim11_to_Dim2 = ConvAry
    
End Function
