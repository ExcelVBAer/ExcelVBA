Attribute VB_Name = "M_Dictionary"
Option Explicit


Public Function Dictionary_Copy(Dic_Base As Scripting.Dictionary) As Scripting.Dictionary
'- 辞書をコピーする
'+ そのまま[Set]すると、同じメモリを使用するので、一方で値を変えると、もう一方の値も変わってしまう為
    
    Dim Dic_Paste       As Scripting.Dictionary
    Dim LoopKey         As Variant
    
    Set Dic_Paste = New Scripting.Dictionary
    
    With Dic_Base
        
        For Each LoopKey In .Keys
            
            Call Dic_Paste.Add(LoopKey, .Item(LoopKey))
            
        Next
        
    End With
    
    Set Dictionary_Copy = Dic_Paste
    
    Set Dic_Paste = Nothing
    
End Function

Public Function Dictionary_Merge(Dic_Base As Scripting.Dictionary, Dic_Add As Scripting.Dictionary) As Scripting.Dictionary
    
    Dim Dic_Merge   As Scripting.Dictionary
    Dim LoopKey     As Variant
    
    If Dic_Base Is Nothing Then
        
        If Dic_Add Is Nothing Then
            
            Set Dic_Merge = Nothing
            
        Else
            
            Set Dic_Merge = Dic_Add
            
        End If
        
    Else
        
        If Dic_Add Is Nothing Then
            
            Set Dic_Merge = Dic_Base
            
        '- 両方の辞書がある場合
        Else
            
            '- ベースの辞書に
            Set Dic_Merge = Dic_Base
            
            '- 追加辞書を結合させる
            With Dic_Merge
                
                For Each LoopKey In Dic_Add.Keys
                    
                    If .Exists(LoopKey) = False Then
                        
                        Call .Add(LoopKey, Dic_Add.Item(LoopKey))
                        
                    End If
                    
                Next
                
            End With
            
        End If
        
    End If
    
    Set Dictionary_Merge = Dic_Merge
    
    Set Dic_Merge = Nothing
    
End Function

Public Function Dictionary_Invert(T_Dic As Scripting.Dictionary) As Variant
'+ KeyとItemを入れ替える
    
    Dim KeyAry      As Variant
    Dim ItemAry     As Variant
    Dim i           As Long
    Dim Cnt_Dic     As Long
    Dim Dic_Invert  As Scripting.Dictionary
    
    Set Dic_Invert = New Scripting.Dictionary
    
    With T_Dic
        
        Cnt_Dic = .Count
        
        If Cnt_Dic <> 0 Then
            
            KeyAry = .Keys
            ItemAry = .Items
            
        End If
        
    End With
    
    With Dic_Invert
        
        For i = 0 To Cnt_Dic - 1
            
            If .Exists(ItemAry(i)) = False Then
                
                Call .Add(ItemAry(i), KeyAry(i))
                
            End If
            
        Next
        
    End With
    
    Set Dictionary_Invert = Dic_Invert
    
    Set Dic_Invert = Nothing
    
End Function

Public Function Dictionary_to_Array(T_Dic As Scripting.Dictionary) As Variant
'+ 辞書を配列にする

    Dim KeyItemAry  As Variant
    Dim KeyAry      As Variant
    Dim ItemAry     As Variant
    Dim Cnt_Dic     As Long
    Dim i           As Long
    
    With T_Dic
        
        Cnt_Dic = .Count
        
        If Cnt_Dic = 0 Then Exit Function
        
        KeyAry = .Keys
        ItemAry = .Items
        
        ReDim KeyItemAry(0 To Cnt_Dic - 1, 0 To 1)
        
        For i = 0 To Cnt_Dic - 1
            
            KeyItemAry(i, 0) = KeyAry(i)
            KeyItemAry(i, 1) = ItemAry(i)
            
        Next
        
    End With
    
    Dictionary_to_Array = KeyItemAry
    
End Function

Public Function Dictionary_Item_Num(DataAry As Variant) As Scripting.Dictionary
'- 配列の値のディクショナリを返す(Key:値,Item:要素数)
'+ 前提：１次元配列

    Dim i               As Long
    Dim T_Dim           As Long
    Dim Dic_Item_Num    As Scripting.Dictionary
    
    Set Dic_Item_Num = New Scripting.Dictionary
    
    T_Dim = Array_DimCount(DataAry)
    
    If T_Dim <> 1 Then Exit Function
    
    With Dic_Item_Num
        
        For i = LBound(DataAry, 1) To UBound(DataAry, 1)
            
            If .Exists(DataAry(i)) Then
                
                Call .Add(DataAry(i), i)
                
            End If
            
        Next
        
    End With
    
    '- 戻り値
    Set Dictionary_Item_Num = Dic_Item_Num
    
    Set Dic_Item_Num = Nothing
    
End Function

Public Function Dictionary_Item_Split(T_Dic As Scripting.Dictionary, Optional Delimiter As String = ",") As Scripting.Dictionary
'+ 辞書のItemをSplitする
    
    Dim i           As Long
    Dim Dic_Split   As Scripting.Dictionary
    Dim KeyAry      As Variant
    Dim ItemAry     As Variant
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim SplitAry()  As String
    
    Set Dic_Split = New Scripting.Dictionary
    
    If T_Dic.Count > 0 Then
        
        With T_Dic
            
            KeyAry = .Keys
            ItemAry = .Items
            
        End With
        
        With Dic_Split
            
            Row_L = LBound(ItemAry, 1)
            Row_U = UBound(ItemAry, 1)
            For i = Row_L To Row_U
                
                If IsArray(ItemAry(i)) = False And IsObject(ItemAry(i)) = False Then
                    
                    SplitAry = Split(CStr(ItemAry(i)), Delimiter)
                    Call .Add(KeyAry(i), SplitAry)
                    
                Else
                    
                    Call .Add(KeyAry(i), ItemAry(i))
                    
                End If
                
            Next
            
        End With
        
    End If
        
    
    Set Dictionary_Item_Split = Dic_Split
    
    Set Dic_Split = Nothing
    
End Function

Public Function Dictionary_Item_Join(T_Dic As Scripting.Dictionary, Optional Delimiter As String = ",") As Scripting.Dictionary
'+ 辞書のItemをSplitする
    
    Dim i           As Long
    Dim Dic_Join    As Scripting.Dictionary
    Dim KeyAry      As Variant
    Dim ItemAry     As Variant
    Dim Row_L       As Long
    Dim Row_U       As Long
    Dim JoinStr     As String
    
    Set Dic_Join = New Scripting.Dictionary
    
    If T_Dic.Count > 0 Then
        
        With T_Dic
            
            KeyAry = .Keys
            ItemAry = .Items
            
        End With
        
        With Dic_Join
            
            Row_L = LBound(ItemAry, 1)
            Row_U = UBound(ItemAry, 1)
            For i = Row_L To Row_U
                
                If IsArray(ItemAry(i)) = True Then
                    
                    JoinStr = Join(ItemAry(i), Delimiter)
                    Call .Add(KeyAry(i), JoinStr)
                    
                Else
                    
                    Call .Add(KeyAry(i), ItemAry(i))
                    
                End If
                
            Next
            
        End With
        
    End If
    
    Set Dictionary_Item_Join = Dic_Join
    
    Set Dic_Join = Nothing
    
End Function

Public Function Dictionary_Paste(Cell As Range, T_Dic As Scripting.Dictionary)
    
    Dim KeyAry      As Variant
    Dim ItemAry     As Variant
    Dim PasteAry    As Variant
    Dim i           As Long
    Dim Cnt_Row     As Long
    Dim Cnt_Col     As Long
    
    If Cell Is Nothing Then Exit Function
    If T_Dic Is Nothing Then Exit Function
    If T_Dic.Count = 0 Then Exit Function
    
    KeyAry = T_Dic.Keys
    ItemAry = T_Dic.Items
    
    ReDim PasteAry(LBound(KeyAry, 1) To UBound(KeyAry, 1), 0 To 1)
    
    For i = LBound(PasteAry, 1) To UBound(PasteAry, 1)
        PasteAry(i, 0) = KeyAry(i)
        PasteAry(i, 1) = ItemAry(i)
    Next
    
    Cnt_Row = UBound(PasteAry, 1) - LBound(PasteAry, 1) + 1
    Cnt_Col = UBound(PasteAry, 2) - LBound(PasteAry, 2) + 1
    Cell.Resize(Cnt_Row, Cnt_Col).Value = PasteAry
    
End Function

Public Function Dictionary_Fill(T_Dic As Scripting.Dictionary, FillItem As Variant)
    
    Dim LoopKey     As Variant
    
    For Each LoopKey In T_Dic.Keys
        
        T_Dic.Item(LoopKey) = FillItem
        
    Next
    
End Function

Public Function Dictionary_Fill_AutoNo(T_Dic As Scripting.Dictionary, Optional StartNo As Long = 1)
    
    Dim LoopKey     As Variant
    Dim T_No        As Long
    
    T_No = StartNo
    
    For Each LoopKey In T_Dic.Keys
        
        T_Dic.Item(LoopKey) = T_No
        
        T_No = T_No + 1
        
    Next
    
End Function
