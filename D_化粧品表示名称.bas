Attribute VB_Name = "D_化粧品表示名称"
'Option Explicit
'
'Option Explicit
'
'
'' スクレイピングに使用(Webから情報取得)
'
'Private Function PRV_Cstr()
'
'    Dim DataAry As Variant
'
'    DataAry = Selection.Value
'
'    DataAry = Array_CStr(DataAry)
'
'    Call Array_Paste(Selection, DataAry)
'
'End Function
'
'Private Function PRV_Test_Web()
'
'    Dim IE          As SC_IE
'    Dim URL_Open    As String
'    Dim LinkAry     As Variant
'    Dim StrHTML     As String
'    Dim LoopTag     As Object
'    Dim LoopTD      As Object
'    Dim C_Tag       As SC_Tag
'    Dim Next_Max    As Long
'    Dim Next_Idx    As Long
'    Dim Next_Str    As String
'    Dim SplitStr()  As String
'    Dim T_Str       As String
'    Dim DataAry()   As String
'    Dim T_Row       As Long
'    Dim T_Col       As Long
'    Dim Dic_Col     As Scripting.Dictionary
'    Dim Max_OnePage As Long
'    Dim Flg_OnePage As Boolean
'    Dim FindWord    As String
'    Dim Max_Item    As Long
'    Dim Row_End     As Long
'    Dim FindAry     As Variant
'    Dim Idx_Find    As Long
'    Dim Dic_Item    As Scripting.Dictionary
'    Dim Flg_Get     As Boolean
'
'    '- 全成分の数を設定
'    Max_Item = 11883
'
'    '- 1ページに表示される最大件数を設定
'    Max_OnePage = 10
'
'    Set C_Tag = New SC_Tag
'    Set Dic_Col = New Scripting.Dictionary
'    Set IE = New SC_IE
'
'    With Dic_Col
'        Call .Add("成分番号", 1)
'        Call .Add("表示名称", 2)
'        Call .Add("INCI名", 3)
'        Call .Add("定義", 4)
'    End With
'
'    Set Dic_Item = PRV_Get_Dic_Item(ThisWorkbook.Worksheets("成分表示名称リスト"))
'
'    '- URLを設定
'    URL_Open = "http://www.jcia.org/n/biz/ln/b/"
'
'    '- IEをセット
'    Call IE.Open_(URL_Open, False)
'
'    For Idx_Find = 1 To 9
'
'        FindWord = Idx_Find
'
'        '- まずは検索を実行
'        For Each LoopTag In IE.Get_Tag(htg_Input)
'
'            Set C_Tag.Tag = LoopTag
'
'            With C_Tag
'
'                If .Type_ = "text" Then
'
'                    If .Name = "word" Then
'
'                        '- 検索ワードを指定
'                        .Value = FindWord
'
'                    End If
'
'                ElseIf .Type_ = "submit" Then
'
'                    If .ID = "searchBtn" Then
'
'                        .Click
'                        Exit For
'
'                    End If
'
'                End If
'
'            End With
'
'        Next
'
'        '- 次へのマークを設定
'        Next_Str = "?word=" & FindWord & "&pageIdx="
'
'        '- [次へ]の最大数を取得
'        Next_Max = 0
'        For Each LoopTag In IE.Get_Tag(htg_A)
'
'            T_Str = LoopTag.href
'
'            If InStr(1, T_Str, Next_Str) <> 0 Then
'
'                SplitStr = Split(T_Str, "=")
'
'                If Next_Max < CLng(SplitStr(2)) Then
'
'                    Next_Max = CLng(SplitStr(2))
'
'                End If
'
'            End If
'
'        Next
'
'        '- データ格納配列を調整
'        ReDim DataAry(1 To (Next_Max + 1) * Max_OnePage, 1 To Dic_Col.Count)
'
'        '- 1ページのみ取得か設定
'        Flg_OnePage = False
'
'        If Flg_OnePage = True Then
'            Next_Idx = 968
'            URL_Open = "http://www.jcia.org/n/biz/ln/b/?word=" & FindWord & "&pageIdx=" & (Next_Idx - 1)
'            Call IE.Navigate(URL_Open)
'            Call IE.Wait
'        End If
'
'        For Next_Idx = 1 To Next_Max
'
'            '- ステータスバーを設定
'            Application.StatusBar = FindWord & ":" & Next_Idx & "/" & Next_Max & ":" & Dic_Item.Count & "/" & Max_Item
'
'            DoEvents
'            Call App_Wait(1)
'            DoEvents
'
''            IE.Visible = False
'
'            '- データテーブルを,配列に格納
'            If IE.Get_Tag(htg_TBODY) Is Nothing Then
'
''                Call MsgBox("取得に失敗しています", vbCritical): Stop
'
'            Else
'
'                For Each LoopTag In IE.Get_Tag(htg_TBODY)
'
'                    '- 初期化
'                    T_Col = 0
'
'                    If IE_Get_Tag(LoopTag, htg_TD) Is Nothing Then
'
''                        Call MsgBox("取得に失敗しています", vbCritical): Stop
'
'                    Else
'
'                        For Each LoopTD In IE_Get_Tag(LoopTag, htg_TD)
'
'                            Set C_Tag.Tag = LoopTD
'
'                            T_Str = C_Tag.innerText
'
'                            '- 先に列番号を取得
'                            If Dic_Col.Exists(T_Str) = True Then
'                                T_Col = Dic_Col.Item(T_Str)
'                            Else
'
'                                '- データの行数を取得
'                                If InStr(1, T_Str, "検索結果の") <> 0 Then
'
''                                    If Dic_Item.Count = Max_Item Then GoTo Terminate
'
'                                    T_Str = Replace(T_Str, "検索結果の", "")
'                                    T_Str = Replace(T_Str, "番目", "")
'                                    T_Row = CLng(T_Str)
'
'                                Else
'
'                                    '- データの列番号が取得できていた場合,データを格納
'                                    '+ 前提:次は必ずデータ
'                                    If T_Col <> 0 Then
'
'                                        If T_Col = 1 Then
'                                            Flg_Get = Not (Dic_Item.Exists(T_Str))
'                                            If Flg_Get = True Then
'                                                Call Dic_Item.Add(T_Str, 1)
'                                            Else
'                                                Dic_Item.Item(T_Str) = Dic_Item.Item(T_Str) + 1
'                                            End If
'                                        End If
'
'                                        If Flg_Get = True Then
'
'                                            DataAry(T_Row, T_Col) = T_Str
'
'                                        End If
'
'                                        '- 該当列を初期化
'                                        T_Col = 0
'
'                                     End If
'
'                                End If
'
'                            End If
'
'                        Next
'
'                    End If
'
'                Next
'
'            End If
'
'            '- 次のリンクをクリック
'            If Flg_OnePage = False Then
'                For Each LoopTag In IE.Get_Tag(htg_A)
'                    Set C_Tag.Tag = LoopTag
'                    If InStr(1, C_Tag.href, Next_Str & (Next_Idx + 1)) <> 0 Then
'                        C_Tag.Click
'                        Exit For
'                    End If
'                Next
'            End If
'
'        Next
'
'Terminate:
'
'        With Sheet_p
'
'            Row_End = FNC_Range_EndRow(.Cells(1, 1))
'
'            Call Array_Paste(.Cells(Row_End + 1, 1), DataAry)
'
'        End With
'
'    Next
'
'    Call IE.Close_
'
'    '- 解放
'    Set LoopTag = Nothing
'    Set IE = Nothing
'
'End Function
'
'Private Function PRV_Get_Dic_Item(T_Sheet As Worksheet) As Scripting.Dictionary
'
'    Dim Row_End     As Long
'    Dim DataAry     As Variant
'    Dim T_Row       As Long
'    Dim Dic_Item    As Scripting.Dictionary
'    Dim T_Item      As String
'
'    Set Dic_Item = New Scripting.Dictionary
'
'    With T_Sheet
'
'        Row_End = FNC_Range_EndRow(.Cells(1, 1))
'
'        DataAry = FNC_Range_Value(.Range(.Cells(1, 1), .Cells(Row_End, 1)), True)
'
'    End With
'
'    With Dic_Item
'
'        For T_Row = LBound(DataAry, 1) To UBound(DataAry, 1)
'
'            T_Item = DataAry(T_Row)
'
'            If T_Item <> "" Then
'
'                If .Exists(T_Item) = False Then
'
'                    Call .Add(T_Item, Empty)
'
'                End If
'
'            End If
'
'        Next
'
'    End With
'
'    Set PRV_Get_Dic_Item = Dic_Item
'
'    Set Dic_Item = Nothing
'
'End Function
'
'Public Function IE_Set(Optional T_URL As String = "", Optional Visible As Boolean = True) As InternetExplorer
''- オブジェクト作成
'
'    Dim IE      As InternetExplorer
'
'    Set IE = New InternetExplorer
'
'    '- 表示のON/OFF
'    IE.Visible = Visible
'
'    '- URLの設定があった場合
'    If T_URL <> "" Then
'
'        '- サイトを開く
'        Call IE.Navigate(T_URL)
'
'        '- 開かれるまで待つ
'        Call IE_Wait_Navigation(IE)
'
'    End If
'
'    Set IE_Set = IE
'
'    Set IE = Nothing
'
'End Function
'
'Public Function IE_Wait_Navigation(IE As Object)
''- 画面移動の完了待ち
'
''    Do While IE.Busy Or IE.readyState < 4
''
''        DoEvents
''
''    Loop
'
'    '- 必ず1秒は待つ
'    Call App_Wait(1)
'
'    Do While IE.readyState <> 4                            'サイトが開かれるまで待つ（お約束）
'
'        Do While IE.Busy = True                              'サイトが開かれるまで待つ（お約束）
'
'            Call App_Wait(1)
'
'        Loop
'
'    Loop
'
'End Function
'
'Private Function PRV_Convert_URL(IE_URL As E_IE_URL) As String
''- 規定のURLを取得
'
'    Dim Ret_URL     As String
'
'    Select Case IE_URL
'
'        Case url_Yahoo
'            Ret_URL = "http://www.yahoo.co.jp/"
'
'        Case url_Google
'            Ret_URL = "https://www.google.co.jp/"
'
'    End Select
'
'    '- 戻り値
'    PRV_Convert_URL = Ret_URL
'
'End Function
'
'
'Private Function PRV_Get_HTML(IE As InternetExplorer) As String
''- HyperText Markup Language（ハイパーテキスト マークアップ ランゲージ）を取得
'
'    PRV_Get_HTML = IE.Document.Body.InnerHTML
'
'End Function
'
'Private Function PRV_Get_Body(IE As InternetExplorer) As String
''- HyperText Markup Language（ハイパーテキスト マークアップ ランゲージ）を取得
'
'    PRV_Get_Body = IE.Document.Body.innerText
'
'End Function
'
'Public Function IE_CripBoard(IE As InternetExplorer) As Variant
'
'    Dim Text_CB     As Stream
'
'    With IE.Document.parentWindow.ClipBoardData
'
'        ''ClipBoardの内容をクリアする
'        .ClearData "text"
'
'        ''ClipBoardに文字列をセットする
'        .SetData "text", "We are REDS!!"
'
'        ''ClipBoardの文字列を取得する
'        Text_CB = .GetData("text")
'
'    End With
'
'End Function
'
'Sub Googleで検索()
'
'    ' IEを立ち上げて Google を開く
'    Dim IE As Object
''    Set IE = new_ie("http://www.google.co.jp")
'
'    ' 検索キーワードを入力
''    type_val ie, "q", "ホゲラッチョ" 　    '+ Googleの仕様変更に伴い変更 2012/12/01
''    type_val IE, "lst-ib", "ホゲラッチョ"
'
'    ' 検索ボタンクリック
''    submit_click IE, "btnG"
'
'    ' 検索結果の 1 件目のタイトルを表示
'    MsgBox domselec(IE, Array( _
'        "id", "res", _
'        "tag", "li", 0, _
'        "tag", "h3", 0 _
'    )).innerText
'
'    ' IEを閉じる
'    IE.Quit
'    Set IE = Nothing
'
'End Sub
'
'Public Function IE_Get_ID(IE As InternetExplorer, ID As String) As HTMLElementCollection
'    ' 注：IEのgetElementByIdはnameも参照する
'    Set IE_Get_ID = IE.Document.getElementById(ID)
'
'End Function
'
'Public Function IE_Get_TagName(IE As InternetExplorer, NamE_HTML_Tag As String) As HTMLElementCollection
'' getElementsByTagName
'
'    Set IE_Get_TagName = IE.getElementsByTagName(NamE_HTML_Tag)
'
'End Function
'
''' 入力します
''Sub type_val(IE, dom_id, val)
''    gid(IE, dom_id).value = val
''    Sleep 100
''End Sub
''
''' 送信ボタンやリンクをクリック
''Sub submit_click(IE, dom_id)
''    gid(IE, dom_id).Click
''    waitIE IE
''End Sub
'
'' 簡易DOMセレクタ
'Function domselec(IE, arr)
'    Dim parent_obj      As Object
'    Dim child_obj       As Object
'    Dim cur             As Long
'    Dim continue_flag   As Boolean
'    Dim dom_id
'    Dim tag_name
'    Dim index_num
'
'    Set parent_obj = IE.Document
'
'    ' 条件配列内で階層を深めていく
'    cur = 0
'    continue_flag = True
'    Do While continue_flag = True
'
'        ' 適用メソッドの種類を判定
'        If arr(cur) = "id" Then
'
'            ' getElementById
'            dom_id = arr(cur + 1)
'            Set child_obj = parent_obj.getElementById(dom_id)
'
'            ' 条件配列内のカーソルを進める
'            cur = cur + 2
'
'        ElseIf arr(cur) = "tag" Then
'
'            ' getElementsByTagName
'            tag_name = arr(cur + 1)
'            index_num = arr(cur + 2)
'            Set child_obj = parent_obj.getElementsByTagName(tag_name)(index_num)
'
'            ' 条件配列内のカーソルを進める
'            cur = cur + 3
'
'        End If
'
'        ' 取得したオブジェクトを次の階層の親オブジェクトとする
'        Set parent_obj = child_obj
'
'        ' 条件配列の終端まで来たか
'        If cur > UBound(arr) Then
'            continue_flag = False
'        End If
'
'    Loop
'
'    Set domselec = parent_obj
'
'End Function
'
'
'' チェックボックスの状態をセットします
'Sub set_check_state(IE, dom_id, checked_flag)
'    ' 希望通りのチェック状態でなければクリック
''    If Not (gid(IE, dom_id).Checked = checked_flag) Then
''        ie_click IE, dom_id
''    End If
'End Sub
'
'
'' セレクトボックスを文言ベースで選択します
'Sub select_by_label(IE, dom_id, label)
'
'    Dim opts
'    Dim i       As Long
'
'    If Len(label) < 1 Then
'      Exit Sub
'    End If
'
''    Set opts = gid(IE, dom_id).Options
'    For i = 0 To opts.Length - 1
'        ' textが同じか
'        If opts(i).innerText = label Then
'            opts(i).Selected = True
'            Exit Sub
'        End If
'    Next i
'
'End Sub
'
'
'' ラジオボタンを値ベースで選択します
'Sub select_radio_by_val(IE, post_name, Value)
'
'    Dim radios
'    Dim i           As Long
'
'    If Len(Value) < 1 Then
'        Exit Sub
'    End If
'
'    Set radios = IE.Document.getElementsByName(post_name)
'    For i = 0 To radios.Length - 1
'        If radios(i).Value = CStr(Value) Then
'            radios(i).Click
'
''            Sleep 100
'        End If
'    Next i
'
'End Sub
'
'
