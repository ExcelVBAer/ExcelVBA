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
    
    '- 対象の配列が数値か判定
    Flg_IsNumeric = True
    For i = At_Min To At_Max
        If IsNumeric(SortAry(i)) = False Then
            Flg_IsNumeric = False
            Exit For
        End If
    Next
    
    '- クイックソート実行結果を配列に格納
    If Flg_IsNumeric = True Then
        '+ 数値専用
        Call QuickSort_Num(SortAry, At_Min, At_Max)
    Else
        '+ 全値
        Call QuickSort_Variant(SortAry, At_Min, At_Max)
    End If
    
End Function

Private Function QuickSort_Num(SortAry As Variant, ByVal At_Min As Long, ByVal At_Max As Long)
'+ 配列は一次元
'+  回帰処理

    Dim At_Mid          As Long
    Dim Val_Mid         As Double
    Dim At_Next         As Long
    Dim At_Temp         As Long
    Dim Val_Temp        As Double
    
    '- 左の値が右の値を超えた場合終了
    If At_Min >= At_Max Then Exit Function
    
    '- 中央の位置を算出してピボットを設定
    At_Mid = (At_Max + At_Min) \ 2
    
    Val_Mid = SortAry(At_Mid)
    
    '- 開始位置要素を中央にセット
    SortAry(At_Mid) = SortAry(At_Min)
    
    At_Temp = At_Min
    
    At_Next = At_Min + 1
    Do While At_Next <= At_Max
        
        If SortAry(At_Next) < Val_Mid Then
            
            '- 値が より小さい場合、基準値を入れ替える
            At_Temp = At_Temp + 1
            Val_Temp = SortAry(At_Temp)
            SortAry(At_Temp) = SortAry(At_Next)
            SortAry(At_Next) = Val_Temp
            
        End If
        At_Next = At_Next + 1
    Loop
    
    '- 入れ替えた値を開始位置にセットし、
    '- 別の変数に確保していた値を最後に入れ替えた位置にセット
    SortAry(At_Min) = SortAry(At_Temp)
    SortAry(At_Temp) = Val_Mid
    
    '---------------------------------------------------------------------------
    
    ' 分割前半を再帰呼び出しでSORT
    Call QuickSort_Num(SortAry, At_Min, At_Temp - 1)
    '---------------------------------------------------------------------------
    ' 分割後半を再帰呼び出しでSORT
    Call QuickSort_Num(SortAry, At_Temp + 1, At_Max)
    
End Function

Private Function QuickSort_Variant(SortAry As Variant, ByVal At_Min As Long, ByVal At_Max As Long)
'+ 配列は一次元
'+  回帰処理

    Dim At_Mid          As Long
    Dim Val_Mid         As Variant
    Dim At_Next         As Long
    Dim At_Temp         As Long
    Dim Val_Temp        As Variant
    
    '- 左の値が右の値を超えた場合終了
    If At_Min >= At_Max Then Exit Function
    
    '- 中央の位置を算出してピボットを設定
    At_Mid = (At_Max + At_Min) \ 2
    
    Val_Mid = SortAry(At_Mid)
    
    '- 開始位置要素を中央にセット
    SortAry(At_Mid) = SortAry(At_Min)
    
    At_Temp = At_Min
    
    At_Next = At_Min + 1
    Do While At_Next <= At_Max
        
        If CStr(SortAry(At_Next)) < CStr(Val_Mid) Then
            
            '- 値がより小さい場合、基準値を入れ替える
            At_Temp = At_Temp + 1
            Val_Temp = SortAry(At_Temp)
            SortAry(At_Temp) = SortAry(At_Next)
            SortAry(At_Next) = Val_Temp
            
        End If
        At_Next = At_Next + 1
    Loop
    
    '- 入れ替えた値を開始位置にセットし、
    '- 別の変数に確保していた値を最後に入れ替えた位置にセット
    SortAry(At_Min) = SortAry(At_Temp)
    SortAry(At_Temp) = Val_Mid
    
    '---------------------------------------------------------------------------
    
    ' 分割前半を再帰呼び出しでSORT
    Call QuickSort_Variant(SortAry, At_Min, At_Temp - 1)
    '---------------------------------------------------------------------------
    ' 分割後半を再帰呼び出しでSORT
    Call QuickSort_Variant(SortAry, At_Temp + 1, At_Max)
    
End Function

Public Function Sort_Bubble(ArrayValue As Variant, _
                            Optional CaseNo As Long = 1, _
                            Optional Target_Col As Long = 0) As Variant
    'バブルソートを実行
    
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim N                   As Long
    Dim S_Row               As Long
    Dim E_Row               As Long
    Dim TempValue           As Variant
    
    S_Row = LBound(ArrayValue, 1)
    E_Row = UBound(ArrayValue, 1)
    
    '行の調整
    If Target_Col = 0 Then
        Target_Col = LBound(ArrayValue, 2)
    End If
    
    For k = E_Row - 1 To S_Row Step -1
        For i = S_Row To k
            j = i + 1
                        
            '昇順(CaseNo=1)
            If CaseNo = 1 Then
                If ArrayValue(i, Target_Col) > ArrayValue(j, Target_Col) Then
                    
                    For N = LBound(ArrayValue, 2) To UBound(ArrayValue, 2)
                        TempValue = ArrayValue(i, N)
                        ArrayValue(i, N) = ArrayValue(j, N)
                        ArrayValue(j, N) = TempValue
                    Next
                    
                End If
            '降順
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
    
    '左端と右端が一致していたならプロシージャを抜ける
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
    
    '左端と右端が一致していたならプロシージャを抜ける
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


