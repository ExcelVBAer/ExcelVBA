Attribute VB_Name = "M_Number"
Option Explicit

Public Function Number_10_to_2(Val_Decimal As Long) As Long
'- 10�i����2�i��
    
    '+ �֐��ł͌��E(511)������
'    Number_10_to_2 = CLng(Application.WorksheetFunction.Dec2Bin(Var_Decimal))
    
    Dim T_Val       As Long
    Dim Val_Bin     As String
    Dim BaseNum     As Long
    
    '- �i����ݒ�
    BaseNum = 2
    
    '- ��U�i�[
    T_Val = Val_Decimal
    
    '- �ϊ�
    Val_Bin = ""
    Do
        
        Val_Bin = CStr(T_Val Mod BaseNum) & Val_Bin
        
        T_Val = T_Val \ BaseNum
        
    Loop Until T_Val = 0
    
    '- �߂�l
    Number_10_to_2 = CLng(Val_Bin)
    
End Function

Public Function Number_2_to_10(Val_Binary As Long) As Long
'- 2�i����10�i��
    
    '+ �֐��ł͌��E(511)������
'    Convert_2_to_10 = CLng(Application.WorksheetFunction.Bin2Dec(Var_Binary))
    
    Dim i       As Long
    Dim Val_Dec As Long
    Dim T_Val   As String
    Dim BaseNum As Long
    Dim Len_Str As Long
    
    '- �i����ݒ�
    BaseNum = 2
    
    '- ��U�i�[
    '+ ���]�����ĉ�����v�Z����
    T_Val = StrReverse(CStr(Val_Binary))
    Len_Str = Len(T_Val)
    
    '- �ϊ�
    Val_Dec = 0
    For i = 1 To Len_Str
        
        Val_Dec = Val_Dec + (CLng(Mid$(T_Val, i, 1)) * (BaseNum ^ (i - 1)))
        
    Next
    
    '- �߂�l
    Number_2_to_10 = Val_Dec
    
End Function

Public Function Number_10_to_16(Val_Decimal As Long, Optional Add_Head As Boolean = False) As String
    
    Dim Head        As String
    
    If Add_Head = True Then
        Head = "&H"
    End If
    
    Number_10_to_16 = Head & Hex(Val_Decimal)
    
End Function

Public Function Number_16_to_10(Val_Hex As String) As Long
    
    Dim Head        As String
    
    Head = "&H"
    
    '- ���������s�v�̏ꍇ�͏���
    If Len(Val_Hex) > 2 Then
        
        If Left$(Val_Hex, 2) = Head Then
            
            Head = ""
            
        End If
        
    End If
    
    Number_16_to_10 = CDec(Head & Val_Hex)
    
End Function

Public Function Number_10_to_nn(Val_Decimal As Long, BaseNum As Long) As Long
'- 10�i����n�i��
    
    Dim T_Val       As Long
    Dim Val_nn      As String
    
    '- ��U�i�[
    T_Val = Val_Decimal
    
    '- �ϊ�
    Val_nn = ""
    Do
        
        Val_nn = CStr(T_Val Mod BaseNum) & Val_nn
        
        T_Val = T_Val \ BaseNum
        
    Loop Until T_Val = 0
    
    '- �߂�l
    Number_10_to_nn = CLng(Val_nn)
    
End Function

Public Function Number_Time_to_Serial(Val_Time As Double) As Double
'���Ԃ̃V���A���l�ϊ�
    
    Dim Num_Date        As Long
    Dim Num_Hour        As Long
    Dim Num_Minute      As Long
    Dim Num_Temp        As Double
    Dim Num_Serial      As Double
    
    '- �Ώےl��0�̏ꍇ�A0��Ԃ�
    If Val_Time = 0 Then
        
        Num_Serial = 0
    
    '- �Ώےl��0�ȊO�̏ꍇ
    Else
        
        '- ���̐��A���Ԃ̐��A���̑��[�������߂�
        Num_Date = Application.WorksheetFunction.RoundDown(Val_Time / 24, 0)
        
        Num_Hour = Application.WorksheetFunction.RoundDown(Val_Time - (Num_Date * 24), 0)
        
        Num_Temp = Val_Time - Application.WorksheetFunction.RoundDown(Val_Time, 0)
        
        Num_Minute = Application.WorksheetFunction.Round(Num_Temp * 60, 0)
        
        '- �ߒl
        Num_Serial = Num_Date + CDbl(TimeSerial(Num_Hour, Num_Minute, 0))
    
    End If
    
    Number_Time_to_Serial = Num_Serial
    
End Function
