Attribute VB_Name = "M_String"
Option Explicit

Public Function String_ASCII(T_Range As Range) As String
'【ASCIIを調べて表示】

    Dim i       As Long
    Dim T_Str   As String
    Dim RetStr  As String
    
    T_Str = CStr(T_Range.Value)
    
    For i = 1 To Len(T_Str)
        
        RetStr = RetStr & Asc(Mid(T_Str, i, 1)) & " "
        
    Next i
    
    '- 末尾の[" "]を削除する
    RetStr = RTrim$(RetStr)
    
    String_ASCII = RetStr
    
End Function

Public Function String_Merge(StrBase As String, StrAdd As String, Optional Delimiter As String = ",") As String
'- 文字列を区切り文字で結合させていく
    
    Dim RetStr      As String
    
    If (StrBase = "" And StrAdd = "") = False Then
    
        If StrBase = "" Then
            
            RetStr = StrAdd
            
        ElseIf StrAdd = "" Then
            
            RetStr = StrBase
            
        Else
        
            RetStr = StrBase & Delimiter & StrAdd
            
        End If
        
    End If
    
    '- 戻り値
    String_Merge = RetStr
    
End Function

Public Function String_to_StrAry(T_Str As String) As String()
    
    Dim i           As Long
    Dim Len_End     As Long
    Dim StrAry()    As String
    
    If T_Str = "" Then Exit Function
    
    '- 配列にする最終位置を格納
    Len_End = Len(T_Str)
    
    ReDim StrAry(1 To Len_End)
    
    '- 1文字づつ,配列に格納
    For i = 1 To Len_End
        
        StrAry(i) = Mid$(T_Str, i, 1)
        
    Next
    
    '- 戻り値
    String_to_StrAry = StrAry
    
End Function

Public Function String_Count(T_String As String, FindStr As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Long
'- 文字列内の指定文字をカウント
'+ 200万Hit:0.30s

    Dim At_Start    As Long
    Dim Cnt_Found   As Long
    Dim Len_Find    As Long
    
    At_Start = 1
    Len_Find = Len(FindStr)
    Do
        
        At_Start = InStr(At_Start, T_String, FindStr, CompareMode) + Len_Find
        
        If At_Start = Len_Find Then
            Exit Do
        Else
            Cnt_Found = Cnt_Found + 1
        End If
        
    Loop
    
    String_Count = Cnt_Found
    
End Function

Public Function String_Repeat(T_String As String, Count As Long) As String
    
    Dim T_Str       As String
    Dim Len_Str     As Long
    
    If T_String = "" Then Exit Function
    If Count = 0 Then Exit Function
    
    Select Case Count
    
    Case 1
        T_Str = T_String
        
    Case Else
        
        Len_Str = Len(T_String)
        
        If Len_Str = 1 Then
            T_Str = String$(Count, T_String)
        Else
            T_Str = Replace(String$(Count, " "), " ", T_String)
        End If
        
    End Select
    
    String_Repeat = T_Str
    
End Function
    
Public Function String_RepeatStr_to_OneStr(LineStr As String, OneStr As String) As String
    
    Dim RetStr  As String
    Dim StrRp   As String
    
    '- 引数を一旦格納
    RetStr = LineStr
    
    '- 指定の連続文字列を作成
    StrRp = OneStr & OneStr
    
    '- 指定の連続文字列がある場合
    Do
        '- ２個→１個に変換
        RetStr = Replace(RetStr, StrRp, OneStr)
        
    Loop Until InStr(1, RetStr, StrRp) = 0
    
    '- 戻り値
    String_RepeatStr_to_OneStr = RetStr
    
End Function

Public Function String_At(T_String As String, T_At As Long) As String
    
    Dim Len_Str     As Long
    
    If T_String = "" Then Exit Function
    
    Len_Str = Len(T_String)
    If Len_Str < T_At Then Exit Function
    
    String_At = Mid$(T_String, T_At, 1)
    
End Function

Public Function String_Find_First(T_String As String, FindStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    String_Find_First = InStr(1, T_String, FindVal, Compare)
    
End Function

Public Function String_Find_Last(T_String As String, FindStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim FindAt  As Long
    Dim FoundAt As Long
    
    Do
        
        FoundAt = FindAt
        
        FindAt = InStr(FindAt + 1, T_String, FindVal, Compare)
        
    Loop Until FindAt = 0
    
    String_Find_Last = FoundAt
    
End Function

Public Function String_Insert(T_String As String, At_Insert As Long, Str_Insert As String) As String
    
    Dim T_Str       As String
    Dim Len_Str     As String
    
    Len_Str = Len(T_String)
    
    '- 追加箇所が文字列の途中の場合
    If At_Insert < Len_Str Then
        
        '- 途中に追加
        T_Str = Left$(T_String, At_Insert) & At_Insert & Mid$(T_String, At_Insert + 1)
        
    Else
    
        '- 後ろに追加
        T_Str = T_String & Str_Insert
        
    End If
    
    String_Insert = T_Str
    
End Function

Public Function String_Chop(T_String As String)
    
    T_String = Left$(T_String, Len(T_String) - 1)
    
End Function

Public Function String_Chomp(T_String As String, Optional Sepalator As String = "")
    
    If Sepalator = "" Then
        
        Select Case String_Last(T_String)
        
        Case vbCrLf, vbCr, vbLf
            
            Call String_Chop(T_String)
        
        End Select
        
    Else
        
        If String_Last(T_String) = Sepalator Then
            
            Call String_Chop(T_String)
            
        End If
        
    End If
    
End Function

Public Function String_Spilt(T_String As String, Delimiter As String) As String()
    
    If T_String = "" Then Exit Function
    
    String_Spilt = Split(T_String, Delimiter)
    
End Function

Public Function String_First(T_String As String) As String
    
    If T_String = "" Then Exit Function
    String_First = Left$(T_String, 1)
    
End Function

Public Function String_Last(T_String As String) As String
    
    If T_String = "" Then Exit Function
    String_Last = Right$(T_String, 1)
    
End Function

Public Function String_Left_With(T_String As String, WithStr As String, Optional Start As Long = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    
    If T_String = "" Or WithStr = "" Then Exit Function
    
    If InStr(Start, T_String, WithStr, Compare) = Start Then
        String_Left_With = True
    End If
    
End Function

Public Function String_Left_Fix(T_String As String, FixStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    If T_String = "" Or FixStr = "" Then Exit Function
    
    If String_Left_With(T_String, FixStr, Compare) = False Then
        String_Left_Fix = FixStr & T_String
    Else
        String_Left_Fix = T_String
    End If
    
End Function

Public Function String_Left_Chop(T_String As String, ChopStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    If T_String = "" Or ChopStr = "" Then Exit Function
    
    If String_Left_With(T_String, ChopStr, Compare) = True Then
        String_Left_Chop = Mid$(T_String, Len(ChopStr) + 1)
    Else
        String_Left_Chop = T_String
    End If
    
End Function

Public Function String_Left_Count(T_String As String, CountStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim i       As Long
    Dim Cnt     As Long
    Dim At_Cnt  As Long
    Dim At_Fnd  As Long
    Dim Len_Cnt As Long
    
    If T_String = "" Or CountStr = "" Then Exit Function
    
    Len_Cnt = Len(CountStr)
    At_Cnt = 1
    Do
        At_Fnd = InStr(At_Cnt, T_String, CountStr, Compare)
        If At_Fnd <> At_Cnt Then Exit Do
        Cnt = Cnt + 1
        At_Cnt = At_Fnd + Len_Cnt
    Loop
    
    String_Left_Count = Cnt
    
End Function

Public Function String_Mid_Count(T_String As String, CountStr As String, Optional Start As Long = 1, Optional Finish As Long = 0, _
                                 Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim i       As Long
    Dim Cnt     As Long
    Dim At_End  As Long
    Dim At_Cnt  As Long
    Dim At_Fnd  As Long
    Dim Len_Cnt As Long
    
    If T_String = "" Or CountStr = "" Then Exit Function
    
    If Finish = 0 Then
        At_End = Len(T_String)
    Else
        At_End = Finish
    End If
    '- 開始位置の設定が逆の場合、終了
    If At_End < Start Then Exit Function
    
    Len_Cnt = Len(CountStr)
    
    '- 検索エリアが検索文字より小さい場合、終了
    If (At_End - Start + 1) < Len_Cnt Then Exit Function
    
    At_Cnt = Start
    Do
        At_Fnd = InStr(At_Cnt, T_String, CountStr, Compare)
        If At_Fnd <> At_Cnt Then Exit Do
        If At_End < At_Fnd Then Exit Do
        Cnt = Cnt + 1
        At_Cnt = At_Fnd + Len_Cnt
    Loop
    
    String_Mid_Count = Cnt
    
End Function

Public Function String_Mid_With(T_String As String, WithStr As String, Start As Long, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    
    Dim Cnt_Mid As Long
    
    Cnt_Mid = String_Mid_Count(T_String, WithStr, Start, Start + Len(WithStr), Compare)
    
    String_Mid_With = Not CBool(Cnt_Mid = 0)
    
End Function

Public Function String_Right_With(T_String As String, WithStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    
    Dim At_With As Long
    Dim Len_Str As Long
    Dim Len_At  As Long
    
    If T_String = "" Or WithStr = "" Then Exit Function
    
    At_With = InStrRev(T_String, WithStr, , Compare)
    
    If At_With + Len(WithStr) - 1 = Len(T_String) Then
         String_Right_With = True
    End If
    
End Function

Public Function String_Right_Fix(T_String As String, FixStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    If T_String = "" Or FixStr = "" Then Exit Function
    
    If String_Right_With(T_String, FixStr, Compare) = False Then
        String_Right_Fix = T_String & FixStr
    Else
        String_Right_Fix = T_String
    End If
    
End Function

Public Function String_Right_Chop(T_String As String, ChopStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    If T_String = "" Or ChopStr = "" Then Exit Function
    
    If String_Right_With(T_String, ChopStr) = True Then
        String_Right_Chop = Left$(T_String, Len(T_String) - Len(ChopStr))
    Else
        String_Right_Chop = T_String
    End If
    
End Function

Public Function String_Right_Count(T_String As String, CountStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim i       As Long
    Dim Cnt     As Long
    Dim Len_Cnt As Long
    Dim At_Cnt  As Long
    Dim At_Fnd  As Long
    
    If T_String = "" Or CountStr = "" Then Exit Function
    
    Len_Cnt = Len(CountStr)
    
    At_Cnt = Len(T_String)
    Do
        At_Fnd = InStrRev(T_String, CountStr, At_Cnt, Compare)
        If At_Fnd + (Len_Cnt - 1) <> At_Cnt Then Exit Do
        Cnt = Cnt + 1
        At_Cnt = At_Fnd - 1
    Loop
    
    String_Right_Count = Cnt
    
End Function

Public Function String_Random(Min As Long, Max As Long) As String
    
    Dim StrAry()    As String
    Dim i           As Long
    Dim T_Str       As String
    
    Max = Int(Max * Rnd)
    ReDim StrAry(0 To Min + Max)
    For i = 0 To Min + Max
        Randomize
        StrAry(i) = Chr(65 + Int(26 * Rnd))
    Next
    
    T_Str = Join(StrAry, "")
    T_Str = Left$(T_Str, Max)
    
    String_Random = T_Str
    
End Function

Public Function String_Exists(T_String As String, FindStr As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Boolean
    
    If InStr(1, T_String, FindStr, CompareMode) > 0 Then
        
        String_Exists = True
        
    End If
    
End Function

Public Function String_Exists_Any(T_String As String, FindStrs As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Boolean
    
    Dim Len_Find    As Long
    Dim i           As Long
    Dim Flg_Exist   As Boolean
    
    Len_Find = Len(FindStrs)
    
    For i = 1 To Len_Find
        
        If InStr(1, T_String, Mid$(FindStrs, i, 1), CompareMode) <> 0 Then
            
            Flg_Exist = True
            Exit For
            
        End If
        
    Next
    
    String_Exists_Any = Flg_Exist
    
End Function

Public Function String_Equals(Str1 As String, Str2 As String, _
                              Optional IgnoreCase As Boolean = False, _
                              Optional IgnoreWidth As Boolean = False, _
                              Optional IgnoreNonSpace As Boolean = False, _
                              Optional IgnoreKanaType As Boolean = False) As Boolean
    
    Dim T_Str1          As String
    Dim T_Str2          As String
    
    T_Str1 = Str1
    T_Str2 = Str2
    
    If IgnoreWidth = True Then
        T_Str1 = StrConv(T_Str1, vbNarrow)
        T_Str2 = StrConv(T_Str2, vbNarrow)
    End If
    
    If IgnoreCase = True Then
        T_Str1 = LCase(T_Str1)
        T_Str2 = LCase(T_Str2)
    End If
    
    If IgnoreNonSpace = True Then
        T_Str1 = Replace(T_Str1, " ", "")
        T_Str2 = Replace(T_Str2, " ", "")
        T_Str1 = Replace(T_Str1, "　", "")
        T_Str2 = Replace(T_Str2, "　", "")
    End If
    
    If IgnoreKanaType = True Then
        T_Str1 = StrConv(T_Str1, vbKatakana)
        T_Str2 = StrConv(T_Str2, vbKatakana)
    End If
    
    String_Equals = CBool(T_Str1 = T_Str2)
    
End Function

Public Function String_A_z(index As Long) As String
'+ 65-90(A-Z)→1-26
'+ 97-122(a-z)→27-52
    
    Dim T_Index     As Integer
    Dim Idx_Max     As Integer
    Dim Str_A_z     As String
    Dim Idx_Adjust  As Integer
    
    Idx_Max = 52
    
    '- アルファベットのインデックスを調整
    T_Index = index Mod Idx_Max
    If T_Index = 0 Then
        T_Index = Idx_Max
    End If
    
    '- 文字コードのインデックスを調整
    If T_Index <= 26 Then
        Idx_Adjust = 65
    Else
        Idx_Adjust = 71
    End If
    
    Str_A_z = Chr(Idx_Adjust + T_Index - 1)
    
    String_A_z = Str_A_z
    
End Function

Public Function String_A_z_All() As String
    
    Dim i       As Long
    
    '- アルファベット文字列を作成
    For i = 1 To 52
        String_A_z_All = String_A_z_All & String_A_z(i)
    Next
    
End Function

Public Function String_Marks(Optional Wide As Boolean = False) As String()
    
    Dim MarkAry()   As String
    Dim Mark        As String
    Dim i           As Long
    Dim Idx_Max     As Long
    Dim Idx_Mark    As Long
    Dim Dic_Idx     As Scripting.Dictionary
    
    Set Dic_Idx = New Scripting.Dictionary
    
    Idx_Max = 255
    
    '- 記号のインデックス辞書を作成
    For i = 0 To Idx_Max
        Select Case i
        Case 32 To 47, 58 To 64, 91 To 96, 123 To 126, 160 To 165, 222 To 223
            If Dic_Idx.Exists(i) = False Then
                Call Dic_Idx.Add(i, Empty)
            End If
        End Select
    Next
    
    '- 格納配列を調整
    '+ カンマが含まれる為,後でSplitができないので
    ReDim MarkAry(0 To Dic_Idx.Count - 1)
    
    '- 各記号を配列に格納
    Idx_Mark = 0
    For i = 0 To Idx_Max
        
        If Dic_Idx.Exists(i) = True Then
            
            Mark = Chr(i)
            
            If Wide = True Then
                Mark = StrConv(Mark, vbWide)
            End If
            
            MarkAry(Idx_Mark) = Mark
            
            Idx_Mark = Idx_Mark + 1
            
        End If
        
    Next
    
    '- 戻り値
    String_Marks = MarkAry
    
    Set Dic_Idx = Nothing
    
End Function

Public Function String_CrLf(Repeat As Long) As String
    
    If Repeat > 0 Then
        String_CrLf = String$(Repeat, vbCrLf)
    End If
    
End Function

Public Function String_Cr(Repeat As Long) As String
    
    If Repeat > 0 Then
        String_Cr = String$(Repeat, vbCr)
    End If
    
End Function

Public Function String_Lf(Repeat As Long) As String
    
    If Repeat > 0 Then
        String_Lf = String$(Repeat, vbLf)
    End If
    
End Function

Public Function String_to_Binary(T_Str As String) As String()
    
    Dim ByteAry()   As Byte
    Dim i           As Long
    Dim Binary()    As String
    
    '- 文字列を文字コード配列に変換
    ByteAry = String_to_Byte(T_Str)
    
    ReDim Binary(LBound(ByteAry, 1) To UBound(ByteAry, 1))
    
    '- 文字コードを16進数に変換
    '+ 2バイト文字の考慮は無し
    For i = LBound(ByteAry, 1) To UBound(ByteAry, 1)
        
        Binary(i) = Hex(ByteAry(i))
        
    Next
    
    String_to_Binary = Binary
    
End Function

Public Function String_from_Binary(Binary() As Byte) As String()
    
    Dim StrAry()    As String
    Dim T_Hex       As String
    Dim T_Hex1      As String
    Dim T_Hex2      As String
    Dim T_Str       As String
    Dim i           As Long
    Dim i1          As Long
    Dim i2          As Long
    Dim Flg_Conv    As Boolean
    Dim T_Byte      As Byte
    Dim T_Byte1     As Byte
    Dim T_Byte2     As Byte
    
    ReDim StrAry(LBound(Binary, 1) To UBound(Binary, 1))
    
    Flg_Conv = False
    For i = LBound(Binary, 1) To UBound(Binary, 1)
        
        T_Byte = Binary(i)
        
        If Flg_Conv = False Then
        
        Select Case T_Byte
            
            Case 0 To 127, 161 To 223 '★英数字とカナ
                
                '- 文字コードを文字に変換
                T_Str = Chr(T_Byte)
                StrAry(i) = T_Str
                
            Case Else
                   
                '- 変換フラグをスイッチ
                Flg_Conv = True
                
                '- インデックスを保持
                i1 = i
                
                '- バイナリ値を保持
                T_Byte1 = T_Byte
                T_Byte2 = 0
                
                '- 文字コードとして文字に試し変換
                T_Str = Chr(T_Byte)
                StrAry(i) = T_Str
                    
                
        End Select
        
        Else
        
                    '- 変換フラグをスイッチ
                    Flg_Conv = False
                    
                    '- インデックスを保持
                    i2 = i
                    
                    '- バイナリ値を保持
                    T_Byte2 = T_Byte
                    
                    '- 連続している場合
                    If i1 + 1 = i2 Then
                        
                        '- 合わせて変換
                        T_Str = Chr(CLng("&H" & (Hex(T_Byte1) & Hex(T_Byte2))))
                        
                        '- 差し替える
                        StrAry(i - 1) = T_Str
                    
                    Else
                           
                        '- 文字コードとして文字に試し変換
                        T_Str = Chr(T_Byte)
                        StrAry(i) = T_Str
                        
                    End If
        
        End If
        
    Next
    
    String_from_Binary = StrAry
    
End Function

Public Function String_to_Byte(T_Str As String) As Byte()
    
    Dim StrAry()    As String
    Dim i           As Long
    Dim ByteAry()   As Byte
    
    If T_Str = "" Then Exit Function
    
    StrAry = String_to_StrAry(T_Str)
    
    ReDim ByteAry(LBound(StrAry, 1) To UBound(StrAry, 1))
    
    For i = LBound(StrAry, 1) To UBound(StrAry, 1)
        
        ByteAry(i) = Asc(StrAry(i))
        
    Next
    
    String_to_Byte = ByteAry
    
End Function

Public Function String_Join(Delimter As String, ParamArray Vals()) As String
'+ 配列の要素の場合、[Cstr]で渡さないとおかしくなる
    
    String_Join = Join(Vals, Delimter)
    
End Function

Public Function String_Is_Hiragana(T_Str As String) As Boolean
    
    Dim i       As Long
    Dim Len_Str As Long
    
    Len_Str = Len(T_Str)
    
    For i = 1 To Len_Str
        
        Select Case Mid$(T_Str, i, 1)
        
        Case "あ" To "ん"
        
        Case Else
            Exit Function
            
        End Select
        
    Next
    
    String_Is_Hiragana = True
    
End Function

Public Function String_Is_Katakana(T_Str As String) As Boolean
    
    Dim i       As Long
    Dim Len_Str As Long
    
    Len_Str = Len(T_Str)
    
    For i = 1 To Len_Str
        
        Select Case Mid$(T_Str, i, 1)
        
        Case "ア" To "ン"
        
        Case "ｱ" To "ﾝ"
        
        Case Else
            Exit Function
            
        End Select
        
    Next
    
    String_Is_Katakana = True
    
End Function

Public Function String_Is_Alphabet(T_Str As String) As Boolean
    
    Dim i       As Long
    Dim Len_Str As Long
    
    Len_Str = Len(T_Str)
    
    For i = 1 To Len_Str
        
        Select Case Mid$(T_Str, i, 1)
        
        Case "a" To "z"
        
        Case "A" To "Z"
        
        Case "Ａ" To "Ｚ"
        
        Case "ａ" To "ｚ"
        
        Case Else
            Exit Function
            
        End Select
        
    Next
    
    String_Is_Alphabet = True
    
End Function

Public Function String_Is_Number(T_Str As String) As Boolean
    
    Dim i       As Long
    Dim Len_Str As Long
    
    Len_Str = Len(T_Str)
    
    For i = 1 To Len_Str
        
        Select Case Mid$(T_Str, i, 1)
        
        Case "0" To "9"
        
        Case "０" To "９"
        
        Case Else
            Exit Function
            
        End Select
        
    Next
    
    String_Is_Number = True
    
End Function

Public Function String_Is_Mark(T_Str As String) As Boolean
    
    Dim i           As Long
    Dim Len_Str     As Long
    Dim MarkAry()   As String
    Dim Dic_Mark    As Scripting.Dictionary
    
    '+ 記号辞書を作成(半角、全角)
    Set Dic_Mark = New Scripting.Dictionary
    MarkAry = String_Marks()
    With Dic_Mark
        For i = LBound(MarkAry, 1) To UBound(MarkAry, 1)
            If .Exists(MarkAry(i)) = False Then
                Call .Add(MarkAry(i), Empty)
            End If
            If .Exists(StrConv(MarkAry(i), vbWide)) = False Then
                Call .Add(StrConv(MarkAry(i), vbWide), Empty)
            End If
        Next
    End With
    
    Len_Str = Len(T_Str)
    
    For i = 1 To Len_Str
        
        If Dic_Mark.Exists(Mid$(T_Str, i, 1)) = False Then Exit For
        
    Next
    
    If i = Len_Str + 1 Then
        String_Is_Mark = True
    End If
    
    Set Dic_Mark = Nothing
    
End Function

Public Function String_Is_Word(T_Str As String) As Boolean
    
    '+ 数字、記号以外 ＝ 文字(アルファベット、ひらがな、カタカナ、漢字)
    If String_Is_Number(T_Str) = True Then Exit Function
    
    If String_Is_Mark(T_Str) = True Then Exit Function
    
    If String_Is_Alphabet(T_Str) = False Then Exit Function
    
    If String_Is_Hiragana(T_Str) = False Then Exit Function
    
    If String_Is_Katakana(T_Str) = False Then Exit Function
    
    If String_Is_Kanji(T_Str) = False Then Exit Function
    
    String_Is_Word = True
    
End Function

Public Function String_Is_Kanji(T_Str As String) As Boolean
    
    
    Dim i       As Long
    Dim Len_Str As Long
    Dim Str1    As String
    
    Len_Str = Len(T_Str)
    
    '- 文字コードが255以下の場合、漢字は有り得ない
    For i = 1 To Len_Str
        
        Str1 = Mid$(T_Str, i, 1)
        
        If Asc(Str1) < 256 Then Exit Function
        
    Next
    
    '+ 数字、アルファベット、記号、ひらがな、カタカナ以外 ＝ 漢字
    If String_Is_Number(T_Str) = True Then Exit Function
    
    If String_Is_Alphabet(T_Str) = True Then Exit Function
    
    If String_Is_Mark(T_Str) = True Then Exit Function
    
    If String_Is_Hiragana(T_Str) = True Then Exit Function
    
    If String_Is_Katakana(T_Str) = True Then Exit Function
    
    String_Is_Kanji = True
    
End Function

Public Function String_Is_Narrow(T_Str As String) As Boolean
    
    String_Is_Narrow = CBool(T_Str = StrConv(T_Str, vbNarrow))
    
End Function

Public Function String_Is_Wide(T_Str As String) As Boolean
    
    String_Is_Wide = CBool(T_Str = StrConv(T_Str, vbWide))
    
End Function

Public Function String_Is_Upper(T_Str As String) As Boolean
    
    String_Is_Upper = CBool(T_Str = StrConv(T_Str, vbUpperCase))
    
End Function

Public Function String_Is_Lower(T_Str As String) As Boolean
    
    String_Is_Lower = CBool(T_Str = StrConv(T_Str, vbLowerCase))
    
End Function

Public Function String_Replace1(T_String As String, Find As String, Replaces As String, _
                                Optional Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As String
    
'- １回だけReplaceする

    String_Replace1 = Replace(T_String, Find, Replaces, 1, 1, Compare)
    
End Function

Public Function String_Binarys_to_Texts(ByteAry() As Byte) As String()
    
    Dim TextAry()       As String
    Dim i               As Long
    Dim T_Byte          As Byte
    Dim T_Byte1         As Byte
    Dim T_Byte2         As Byte
    Dim T_Hex1          As String
    Dim T_Hex2          As String
    Dim T_Str           As String
    Dim Flg_Conv        As Boolean
    Dim i1              As Long
    Dim i2              As Long
    Dim Len_Ary         As Long
    
    '- 配列が無かった場合、終了
    On Error Resume Next
    Len_Ary = UBound(ByteAry, 1) - LBound(ByteAry, 1) + 1
    On Error GoTo 0
    If Len_Ary = 0 Then Exit Function
    
    ReDim TextAry(LBound(ByteAry, 1) To UBound(ByteAry, 1))
    
    '+ 改行コードを探して、改行し、配列で格納する
    
    Flg_Conv = False
    For i = LBound(ByteAry, 1) To UBound(ByteAry, 1)
        
        T_Byte = ByteAry(i)
        
        Select Case T_Byte
            Case 0 To 127, 161 To 223 '★英数字とカナ
                T_Str = Chr(T_Byte)
                TextAry(i) = T_Str
                
            Case Else
                If Flg_Conv = False Then
                    Flg_Conv = True
                    
                    i1 = i
                    i2 = 0
                    T_Byte1 = T_Byte
                    T_Byte2 = 0
                    
                    T_Str = Chr(T_Byte)
                    TextAry(i) = T_Str
                    
                Else
                    Flg_Conv = False
                    
                    i2 = i
                    T_Byte2 = T_Byte
                    
                    If i1 + 1 = i2 Then
                        
                        T_Hex1 = Number_10_to_16(CLng(T_Byte1))
                        T_Hex2 = Number_10_to_16(CLng(T_Byte2))
                        
                        T_Str = Chr(CLng("&H" & T_Hex1 & T_Hex2))
                        
                        TextAry(i - 1) = T_Str
                    
                    Else
                        
                        T_Str = Chr(T_Byte)
                        TextAry(i) = T_Str
                        
                    End If
                    
                End If
                
        End Select
        
    Next
    
    String_Binarys_to_Texts = TextAry
    
End Function

Public Function String_Unvisible_to_Visible(T_Str As String) As String
'+ 制御文字や、目視不能な文字列を、目視可能にする
    
    Dim i           As Long
    Dim Len_Str     As Long
    Dim Str1        As String
    Dim StrAry()    As String
    Dim Dic_Unvisible   As Scripting.Dictionary
    
    If T_Str = "" Then Exit Function
    
    Len_Str = Len(T_Str)
    
    ReDim StrAry(1 To Len_Str)
    
    Set Dic_Unvisible = PRV_Dic_Unvisible
    
    For i = 1 To Len_Str
        
        Str1 = Mid$(T_Str, i, 1)
                
        If Dic_Unvisible.Exists(Str1) = True Then
            StrAry(i) = Dic_Unvisible.Item(Str1)
        Else
            StrAry(i) = Str1
        End If
        
    Next
    
    String_Unvisible_to_Visible = Join(StrAry, "")
    
End Function

Private Function PRV_Dic_Unvisible() As Scripting.Dictionary
    
    Dim Dic_Unvisible   As Scripting.Dictionary
    Dim Str_View        As String
    Dim i               As Long
    
    Set Dic_Unvisible = New Scripting.Dictionary
    
    For i = 0 To 128
    
        Select Case Chr(i)
            
            Case Chr(0):    Str_View = "[NullChar]"
            Case Chr(1):    Str_View = "[Start Of Heading]"
            Case Chr(2):    Str_View = "[Start Of Text]"
            Case Chr(3):    Str_View = "[End Of Text]"
            Case Chr(4):    Str_View = "[End Of Transmission]"
            Case Chr(5):    Str_View = "[Enquery]"
            Case Chr(6):    Str_View = "[Acknowledgement]"
            Case Chr(7):    Str_View = "[Bell]"
            Case Chr(8):    Str_View = "[Back Space]"
            Case Chr(9):    Str_View = "[Tab]"
            Case Chr(10):   Str_View = "[Lf]"
            Case Chr(11):   Str_View = "[VerticalTab]"
            Case Chr(12):   Str_View = "[FormFeed]"
            Case Chr(13):   Str_View = "[Cr]"
            Case Chr(14):   Str_View = "[Shift Out]"
            Case Chr(15):   Str_View = "[Shift In]"
            Case Chr(16):   Str_View = "[Data Link Escape]"
            Case Chr(17):   Str_View = "[Device Control 1]"
            Case Chr(18):   Str_View = "[Device Control 2]"
            Case Chr(19):   Str_View = "[Device Control 3]"
            Case Chr(20):   Str_View = "[Device Control 4]"
            Case Chr(21):   Str_View = "[Negative Acknowledgement]"
            Case Chr(22):   Str_View = "[Synchronous idle]"
            Case Chr(23):   Str_View = "[End of Transmission Block]"
            Case Chr(24):   Str_View = "[Cancel]"
            Case Chr(25):   Str_View = "[End of Medium]"
            Case Chr(26):   Str_View = "[End Of File]"
            Case Chr(27):   Str_View = "[Escape]"
            Case Chr(28):   Str_View = "[File Sepalator]"
            Case Chr(29):   Str_View = "[Group Sepalator]"
            Case Chr(30):   Str_View = "[Record Sepalator]"
            Case Chr(31):   Str_View = "[Unit Sepalator]"
            Case Chr(32):   Str_View = "[Space]"
            Case "　":      Str_View = "[SpaceW]"
            Case Chr(127):  Str_View = "[Delete]"
            Case Chr(128):  Str_View = "[Chr(128)"
            
            Case Else:      Str_View = ""
            
        End Select
        
        If Str_View <> "" Then
              
            Dic_Unvisible.Item(Chr(i)) = Str_View
            
        End If
                
    Next
    
    Set PRV_Dic_Unvisible = Dic_Unvisible
    
    Set Dic_Unvisible = Nothing
    
End Function

Public Function String_LenB(T_Str As String)
'- 文字列のバイト数を返す(マルチバイト(MBCS)に対応)
'+ MBCS: Multibyte Character Set
'+ LenB関数では、全てが2バイト文字として換算されてしまう（[A]が2バイトとして）
    
    String_LenB = LenB(StrConv(T_Str, vbFromUnicode))
    
End Function

Public Function String_to_Array(T_Str As String, Optional T_CrLf As String = vbCrLf, Optional Delimiter As String = ",") As Variant
'- 文字列(テキスト)を配列に変換する
    
    Dim At_S        As Long
    Dim At_E        As Long
    Dim At_CrLf     As Long
    Dim Cnt_Row     As Long
    Dim Cnt_Col     As Long
    Dim DataStr     As String
    Dim DataAry     As Variant
    Dim SplitAry()  As String
    Dim Col_L       As Long
    Dim Col_U       As Long
    Dim i           As Long
    Dim j           As Long
    Dim Len_Str     As Long
    Dim Len_CrLf    As Long
    
    '- 文字列が無かった場合,終了
    If T_Str = "" Then Exit Function
    
    '- 改行コードの文字数を取得
    Len_CrLf = Len(T_CrLf)
    
    '- 行数を取得
    Cnt_Row = 1
    At_CrLf = 1
    Do
        At_CrLf = InStr(At_CrLf, T_Str, T_CrLf) + Len_CrLf
        If At_CrLf = Len_CrLf Then Exit Do
        Cnt_Row = Cnt_Row + 1
    Loop
    
    '- 1行の場合、1次元配列で返す
    If Cnt_Row = 1 Then
        
        SplitAry = Split(T_Str, Delimiter)
        Cnt_Col = UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1
        ReDim DataAry(0 To Cnt_Col - 1)
        
        For i = LBound(SplitAry, 1) To UBound(SplitAry, 1)
            DataAry(i) = SplitAry(i)
        Next
        
    Else
        
        '- 文字列の長さを取得
        Len_Str = Len(T_Str)
        
        At_S = 1
        At_E = 1
        Cnt_Col = 0
        i = 0
        Do
            '- 開始位置が最後の場合,終了位置も最後にする
            If At_S = Len_Str Then
                At_E = Len_Str
                
            Else
            
                '- 終了位置を検索(改行コードの1つ後ろ)
                At_E = InStr(At_S, T_Str, T_CrLf) - 1
                
                '- 終了位置が無ければ、最後まで設定
                If At_E = -1 Then
                    At_E = Len_Str
                End If
                
            End If
            
            '- データがあった場合
            If At_S <= At_E Then
                
                '- データ範囲を切り出す
                DataStr = Mid$(T_Str, At_S, At_E - At_S + 1)
                
                '- データを区切り文字で分割
                SplitAry = Split(DataStr, Delimiter)
                    
                '- 列数をカウントし,格納配列を作成(初回のみ)
                If Cnt_Col = 0 Then
                    Cnt_Col = UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1
                    ReDim DataAry(0 To Cnt_Row - 1, 0 To Cnt_Col - 1)
                End If
                
                Col_L = LBound(SplitAry, 1)
                Col_U = UBound(SplitAry, 1)
                    
                '- 列数が途中で増えている場合、列を拡張して対応
                If Cnt_Col < (Col_U - Col_L + 1) Then
                    Cnt_Col = Col_U - Col_L + 1
                    ReDim Preserve DataAry(0 To Cnt_Row - 1, 0 To Cnt_Col - 1)
                End If
                
                '- 各列の値を格納していく
                For j = Col_L To Col_U
                    DataAry(i, j) = SplitAry(j)
                Next
                    
            '- 改行が続いている場合
            Else
                
                '- 列数を1とする(初回のみ)
                If Cnt_Col = 0 Then
                    Cnt_Col = 1
                    ReDim DataAry(0 To Cnt_Row - 1, 0 To Cnt_Col - 1)
                End If
                
            End If
            
            '- 開始位置を設定
            At_S = At_E + Len_CrLf + 1
            
            '- 次の開始位置が文字列を超えている場合,抜ける
            If Len_Str < At_S Then Exit Do
            
            '- 配列の要素をインクリメント
            i = i + 1
            
        Loop
        
    End If
    
    String_to_Array = DataAry
    
End Function

Public Function String_to_Number(StrNumber As String) As Variant
    
    Dim T_Num       As String
    Dim Len_Num     As Long
    Dim At_Dot      As Long
    Dim Val_Ret     As Variant
    Dim Flg_Under   As Boolean
    Dim Flg_Cur     As Boolean
    Dim Flg_Dbl     As Boolean
    
'    + --もNG
    If IsNumeric(StrNumber) = False Then Exit Function
    
    '- 一旦格納
    T_Num = StrNumber
    
    '- 値調整
    T_Num = Replace(T_Num, ",", "")
    
    '- マイナス判定
    If Left$(T_Num, 1) = "-" Then
        Flg_Under = True
    End If
    
    '- 10進数判定
    If Flg_Under = False Then
        If isNumber(T_Num) = False Then Exit Function
    Else
        If isNumber(Mid$(T_Num, 2)) = False Then Exit Function
    End If
    
    '- 数値部分の長さを取得
    Len_Num = Len(T_Num)
    If Flg_Under = True Then
        Len_Num = Len_Num - 1
    End If
    
    '- ドット一致を取得(小数部判定用)
    At_Dot = InStr(1, T_Num, ".")
    
    '- ドットがあった場合、Dblフラグを立てる
    '+ ドットの有無で強制
    If At_Dot <> 0 Then
        Flg_Dbl = True
    End If
    
    If Flg_Dbl = True Then
        '+ Double   :-1.79769313486232E308 ～ -4.94065645841247E-324(負の値)
        If CStr(CDbl(T_Num)) = T_Num Then
            Val_Ret = CDbl(T_Num)
        Else
            Val_Ret = T_Num
        End If
    Else
        '+ Long     :2,147,483,647 (整数部:10桁)
        If Len_Num < 10 Then
            Val_Ret = CLng(T_Num)
        ElseIf Len_Num = 10 Then
            Val_Ret = CCur(T_Num)
            If -2147483648# <= Val_Ret And Val_Ret <= 2147483647 Then
                Val_Ret = CLng(Val_Ret)
            End If
        Else
            If Len_Num < 15 Then
                Val_Ret = CCur(T_Num)
            ElseIf Len_Num = 15 Then
                '+ Currency :922,337,203,685,477.[5807] (整数部:15桁,小数部:4桁)
                '+ -922,337,203,685,477 ～ 922,337,203,685,477
                If Mid$(T_Num, 1 + Abs(CLng(Flg_Under)), 1) = 9 Then
                    If Flg_Under = False Then
                        If T_Num <= "922337203685477" Then
                            Flg_Cur = True
                        End If
                    Else
                        If T_Num <= "-922337203685477" Then
                            Flg_Cur = True
                        End If
                    End If
                Else
                    Flg_Cur = True
                End If
                
                If Flg_Cur = True Then
                    Val_Ret = CCur(T_Num)
                Else
                    Val_Ret = T_Num
                End If
                
            Else
                Val_Ret = T_Num
            End If
        End If
    End If
    
    String_to_Number = Val_Ret
    
End Function
