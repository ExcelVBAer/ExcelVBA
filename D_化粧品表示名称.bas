Attribute VB_Name = "D_���ϕi�\������"
'Option Explicit
'
'Option Explicit
'
'
'' �X�N���C�s���O�Ɏg�p(Web������擾)
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
'    '- �S�����̐���ݒ�
'    Max_Item = 11883
'
'    '- 1�y�[�W�ɕ\�������ő匏����ݒ�
'    Max_OnePage = 10
'
'    Set C_Tag = New SC_Tag
'    Set Dic_Col = New Scripting.Dictionary
'    Set IE = New SC_IE
'
'    With Dic_Col
'        Call .Add("�����ԍ�", 1)
'        Call .Add("�\������", 2)
'        Call .Add("INCI��", 3)
'        Call .Add("��`", 4)
'    End With
'
'    Set Dic_Item = PRV_Get_Dic_Item(ThisWorkbook.Worksheets("�����\�����̃��X�g"))
'
'    '- URL��ݒ�
'    URL_Open = "http://www.jcia.org/n/biz/ln/b/"
'
'    '- IE���Z�b�g
'    Call IE.Open_(URL_Open, False)
'
'    For Idx_Find = 1 To 9
'
'        FindWord = Idx_Find
'
'        '- �܂��͌��������s
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
'                        '- �������[�h���w��
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
'        '- ���ւ̃}�[�N��ݒ�
'        Next_Str = "?word=" & FindWord & "&pageIdx="
'
'        '- [����]�̍ő吔���擾
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
'        '- �f�[�^�i�[�z��𒲐�
'        ReDim DataAry(1 To (Next_Max + 1) * Max_OnePage, 1 To Dic_Col.Count)
'
'        '- 1�y�[�W�̂ݎ擾���ݒ�
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
'            '- �X�e�[�^�X�o�[��ݒ�
'            Application.StatusBar = FindWord & ":" & Next_Idx & "/" & Next_Max & ":" & Dic_Item.Count & "/" & Max_Item
'
'            DoEvents
'            Call App_Wait(1)
'            DoEvents
'
''            IE.Visible = False
'
'            '- �f�[�^�e�[�u����,�z��Ɋi�[
'            If IE.Get_Tag(htg_TBODY) Is Nothing Then
'
''                Call MsgBox("�擾�Ɏ��s���Ă��܂�", vbCritical): Stop
'
'            Else
'
'                For Each LoopTag In IE.Get_Tag(htg_TBODY)
'
'                    '- ������
'                    T_Col = 0
'
'                    If IE_Get_Tag(LoopTag, htg_TD) Is Nothing Then
'
''                        Call MsgBox("�擾�Ɏ��s���Ă��܂�", vbCritical): Stop
'
'                    Else
'
'                        For Each LoopTD In IE_Get_Tag(LoopTag, htg_TD)
'
'                            Set C_Tag.Tag = LoopTD
'
'                            T_Str = C_Tag.innerText
'
'                            '- ��ɗ�ԍ����擾
'                            If Dic_Col.Exists(T_Str) = True Then
'                                T_Col = Dic_Col.Item(T_Str)
'                            Else
'
'                                '- �f�[�^�̍s�����擾
'                                If InStr(1, T_Str, "�������ʂ�") <> 0 Then
'
''                                    If Dic_Item.Count = Max_Item Then GoTo Terminate
'
'                                    T_Str = Replace(T_Str, "�������ʂ�", "")
'                                    T_Str = Replace(T_Str, "�Ԗ�", "")
'                                    T_Row = CLng(T_Str)
'
'                                Else
'
'                                    '- �f�[�^�̗�ԍ����擾�ł��Ă����ꍇ,�f�[�^���i�[
'                                    '+ �O��:���͕K���f�[�^
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
'                                        '- �Y�����������
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
'            '- ���̃����N���N���b�N
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
'    '- ���
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
''- �I�u�W�F�N�g�쐬
'
'    Dim IE      As InternetExplorer
'
'    Set IE = New InternetExplorer
'
'    '- �\����ON/OFF
'    IE.Visible = Visible
'
'    '- URL�̐ݒ肪�������ꍇ
'    If T_URL <> "" Then
'
'        '- �T�C�g���J��
'        Call IE.Navigate(T_URL)
'
'        '- �J�����܂ő҂�
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
''- ��ʈړ��̊����҂�
'
''    Do While IE.Busy Or IE.readyState < 4
''
''        DoEvents
''
''    Loop
'
'    '- �K��1�b�͑҂�
'    Call App_Wait(1)
'
'    Do While IE.readyState <> 4                            '�T�C�g���J�����܂ő҂i���񑩁j
'
'        Do While IE.Busy = True                              '�T�C�g���J�����܂ő҂i���񑩁j
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
''- �K���URL���擾
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
'    '- �߂�l
'    PRV_Convert_URL = Ret_URL
'
'End Function
'
'
'Private Function PRV_Get_HTML(IE As InternetExplorer) As String
''- HyperText Markup Language�i�n�C�p�[�e�L�X�g �}�[�N�A�b�v �����Q�[�W�j���擾
'
'    PRV_Get_HTML = IE.Document.Body.InnerHTML
'
'End Function
'
'Private Function PRV_Get_Body(IE As InternetExplorer) As String
''- HyperText Markup Language�i�n�C�p�[�e�L�X�g �}�[�N�A�b�v �����Q�[�W�j���擾
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
'        ''ClipBoard�̓��e���N���A����
'        .ClearData "text"
'
'        ''ClipBoard�ɕ�������Z�b�g����
'        .SetData "text", "We are REDS!!"
'
'        ''ClipBoard�̕�������擾����
'        Text_CB = .GetData("text")
'
'    End With
'
'End Function
'
'Sub Google�Ō���()
'
'    ' IE�𗧂��グ�� Google ���J��
'    Dim IE As Object
''    Set IE = new_ie("http://www.google.co.jp")
'
'    ' �����L�[���[�h�����
''    type_val ie, "q", "�z�Q���b�`��" �@    '+ Google�̎d�l�ύX�ɔ����ύX 2012/12/01
''    type_val IE, "lst-ib", "�z�Q���b�`��"
'
'    ' �����{�^���N���b�N
''    submit_click IE, "btnG"
'
'    ' �������ʂ� 1 ���ڂ̃^�C�g����\��
'    MsgBox domselec(IE, Array( _
'        "id", "res", _
'        "tag", "li", 0, _
'        "tag", "h3", 0 _
'    )).innerText
'
'    ' IE�����
'    IE.Quit
'    Set IE = Nothing
'
'End Sub
'
'Public Function IE_Get_ID(IE As InternetExplorer, ID As String) As HTMLElementCollection
'    ' ���FIE��getElementById��name���Q�Ƃ���
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
''' ���͂��܂�
''Sub type_val(IE, dom_id, val)
''    gid(IE, dom_id).value = val
''    Sleep 100
''End Sub
''
''' ���M�{�^���⃊���N���N���b�N
''Sub submit_click(IE, dom_id)
''    gid(IE, dom_id).Click
''    waitIE IE
''End Sub
'
'' �Ȉ�DOM�Z���N�^
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
'    ' �����z����ŊK�w��[�߂Ă���
'    cur = 0
'    continue_flag = True
'    Do While continue_flag = True
'
'        ' �K�p���\�b�h�̎�ނ𔻒�
'        If arr(cur) = "id" Then
'
'            ' getElementById
'            dom_id = arr(cur + 1)
'            Set child_obj = parent_obj.getElementById(dom_id)
'
'            ' �����z����̃J�[�\����i�߂�
'            cur = cur + 2
'
'        ElseIf arr(cur) = "tag" Then
'
'            ' getElementsByTagName
'            tag_name = arr(cur + 1)
'            index_num = arr(cur + 2)
'            Set child_obj = parent_obj.getElementsByTagName(tag_name)(index_num)
'
'            ' �����z����̃J�[�\����i�߂�
'            cur = cur + 3
'
'        End If
'
'        ' �擾�����I�u�W�F�N�g�����̊K�w�̐e�I�u�W�F�N�g�Ƃ���
'        Set parent_obj = child_obj
'
'        ' �����z��̏I�[�܂ŗ�����
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
'' �`�F�b�N�{�b�N�X�̏�Ԃ��Z�b�g���܂�
'Sub set_check_state(IE, dom_id, checked_flag)
'    ' ��]�ʂ�̃`�F�b�N��ԂłȂ���΃N���b�N
''    If Not (gid(IE, dom_id).Checked = checked_flag) Then
''        ie_click IE, dom_id
''    End If
'End Sub
'
'
'' �Z���N�g�{�b�N�X�𕶌��x�[�X�őI�����܂�
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
'        ' text��������
'        If opts(i).innerText = label Then
'            opts(i).Selected = True
'            Exit Sub
'        End If
'    Next i
'
'End Sub
'
'
'' ���W�I�{�^����l�x�[�X�őI�����܂�
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
