Attribute VB_Name = "M_Date"
Option Explicit
'変更点
Public Enum E_Date
    Years
    Months
    Weeks
    Days
    Hours
    Minutes
    Seconds
End Enum

Public Enum E_Round
    Up      '+ 切り上げ
    Down    '+ 切り捨て
    UpDown  '+ 四捨五入
End Enum

Public Function Date_Day_First(T_Date As Date, Optional Add_Month As Long = 0) As Date
    
    Date_Day_First = DateSerial(Year(T_Date), Month(T_Date) + Add_Month, 1)
    
End Function

Public Function Date_Day_Last(T_Date As Date, Optional Add_Month As Long = 0) As Date
    
    Dim TempDate    As Date
    
    TempDate = T_Date
    TempDate = Date_Day_First(T_Date, 1 + Add_Month)
    TempDate = DateAdd("d", -1, TempDate)
    
    Date_Day_Last = TempDate
    
End Function

Public Function Date_DatePart(Optional Year As Long = 2000, Optional Month As Long = 1, Optional Day As Long = 1) As Date
    
    Dim T_Year      As Long
    Dim T_Month     As Long
    Dim T_Day       As Long
    Dim R_Month     As Long
    Dim Date_Ret    As Date
    Dim Add_Year    As Long
    
    If Year = 0 Then
        T_Year = 1
    Else
        T_Year = Year
    End If
    
    If Month = 0 Then
        T_Month = 1
    Else
        T_Month = Month
    End If
    
    If Day = 0 Then
        T_Day = 1
    Else
        T_Day = Day
    End If
    
    Date_Ret = DateSerial(T_Year, T_Month, T_Day)
    
    '- 全年月日を調整した日付を返す
    Date_DatePart = Date_Ret
    
End Function

Public Function Date_TimePart(Optional Hour As Long = 0, Optional Minute As Long = 0, Optional Second As Long = 0) As Date
    
    Dim T_Second    As Long
    Dim T_Minute    As Long
    Dim T_Hour      As Long
    Dim T_Day       As Long
    Dim R_Second    As Long
    Dim R_Minute    As Long
    Dim R_Hour      As Long
    Dim Date_Ret    As Date
    
    '- 時分秒を日時分秒換算でそれぞれ集計
    T_Second = Second
    T_Minute = Minute + Date_Second_to_Minute(T_Second)
    T_Hour = Hour + Date_Minute_to_Hour(T_Minute)
    T_Day = Date_Hour_to_Day(T_Hour)
    
    '- 時分秒の各値を取得
    R_Hour = T_Hour Mod 24
    R_Minute = T_Minute Mod 60
    R_Second = T_Second Mod 60
    
    '- 時分秒を取得
    Date_Ret = TimeSerial(R_Hour, R_Minute, R_Second)
    
    '- 日を追加
    Date_Ret = DateAdd("d", T_Day, Date_Ret)
    
    Date_TimePart = Date_Ret
    
End Function

Public Function Date_Add(T_Date As Date, Optional Year As Long = 0, Optional Month As Long = 0, Optional Day As Long = 0, _
                         Optional Week As Long = 0, Optional Hour As Long = 0, Optional Minute As Long = 0, Optional Second As Long = 0) As Date
'+ 【以下注意】
'+ ※うるう年の場合：DateAddでは2/29→(年を加or減)→2/28　となる
'+ ※うるう年の場合：Date関数では2/29→(年を加or減)→3/1　となる
'+ ※うるう年の場合：DateSerialでは2/29→(年を加or減)→3/1　となる
    
    Dim Date_Ret    As Date
    
    Date_Ret = T_Date
    
    Date_Ret = PRV_Date_Add("yyyy", Year, Date_Ret)
    Date_Ret = PRV_Date_Add("m", Month, Date_Ret)
    Date_Ret = PRV_Date_Add("d", Day, Date_Ret)
    
    Date_Ret = PRV_Date_Add("ww", Week, Date_Ret)
    
    Date_Ret = PRV_Date_Add("h", Hour, Date_Ret)
    Date_Ret = PRV_Date_Add("n", Minute, Date_Ret)
    Date_Ret = PRV_Date_Add("s", Second, Date_Ret)
    
    Date_Add = Date_Ret
    
End Function

Private Function PRV_Date_Add(AddType As String, AddVal As Long, T_Date As Date) As Date
    
    Dim R_Date  As Date
    
    If AddVal <> 0 Then
        
        R_Date = DateAdd(AddType, AddVal, T_Date)
        
    Else
        
        R_Date = T_Date
        
    End If
    
    PRV_Date_Add = R_Date
    
End Function

Public Function Date_Diff(DateType As E_Date, Date1 As Date, Date2 As Date) As Long
        
    Dim DateStr     As String
    Dim T_Diff      As Long
    
    DateStr = DateType_Code_to_String(DateType)
    
    '+ オーバーフローに対応
On Error GoTo Err
    
    T_Diff = DateDiff(DateStr, Date1, Date2)
    
    Date_Diff = T_Diff
    
    Exit Function
Err:
    Call MsgBox("Err:OverFlow", vbCritical)
    
End Function

Private Function DateType_Code_to_String(DateType As E_Date) As String
        
    Dim RetStr      As String
    
    Select Case DateType
    
    Case E_Date.Years
        RetStr = "yyyy"
        
    Case E_Date.Months
        RetStr = "m"
        
    Case E_Date.Weeks
        RetStr = "ww"
    
    Case E_Date.Days
        RetStr = "d"
        
    Case E_Date.Hours
        RetStr = "h"
    
    Case E_Date.Minutes
        RetStr = "n"
    
    Case E_Date.Seconds
        RetStr = "s"
        
    End Select
    
    DateType_Code_to_String = RetStr
    
End Function

Public Function Date_Second_to_Minute(T_Second As Long) As Long
    
    Date_Second_to_Minute = Int(T_Second / 60)
    
End Function

Public Function Date_Minute_to_Hour(T_Minute As Long) As Long
    
    Date_Minute_to_Hour = Int(T_Minute / 60)
    
End Function

Public Function Date_Hour_to_Day(T_Hour As Long) As Long
    
    Date_Hour_to_Day = Int(T_Hour / 24)
    
End Function

Public Function Date_Time_to_Second(T_Date As Date) As Long
    
    Dim T_Hour      As Long
    Dim T_Minute    As Long
    Dim T_Second    As Long
    Dim Second_Sum  As Long
    
    T_Hour = Hour(T_Date)
    T_Minute = Minute(T_Date)
    T_Second = Second(T_Date)
    
    Second_Sum = T_Second + T_Minute * 60 + T_Hour * 60 * 60
    
End Function

Public Function Date_Time_to_Minute(T_Date As Date, Optional RoundCase As E_Round = E_Round.Down) As Long
    
    Dim T_Hour      As Long
    Dim T_Minute    As Long
    Dim T_Second    As Long
    Dim T_Min_Sec   As Long
    Dim Minute_Sum  As Long
    
    T_Hour = Hour(T_Date)
    T_Minute = Minute(T_Date)
    T_Second = Second(T_Date)
    
    Select Case RoundCase
    
    Case E_Round.Up
        T_Min_Sec = IIf(T_Second = 0, 0, 1)
        
    Case E_Round.Down
        T_Min_Sec = 0
        
    Case E_Round.UpDown
        T_Min_Sec = IIf(T_Second < 30, 0, 1)
        
    End Select
    
    Minute_Sum = T_Min_Sec + T_Minute + T_Hour * 60
    
End Function

Public Function Date_IsYear(T_Year As Variant, Optional Min As Long = 1899, Optional Max As Long = 9999) As Boolean
    
    If isNumber(T_Year) = False Then Exit Function
    
    If Min <= CLng(T_Year) And CLng(T_Year) <= Max Then
        Date_IsYear = True
    End If
    
End Function

Public Function Date_IsMonth(T_Month As Variant, Optional Min As Long = 1, Optional Max As Long = 12) As Boolean
    
    If isNumber(T_Month) = False Then Exit Function
    
    If Min <= CLng(T_Month) And CLng(T_Month) <= Max Then
        Date_IsMonth = True
    End If
    
End Function

Public Function Date_IsDay(T_Day As Variant, Optional Min As Long = 1, Optional Max As Long = 31, Optional T_YearMonth As Date = #12/31/1899#) As Boolean
    
    Dim Min_Day     As Long
    Dim Max_Day     As Long
    
    If isNumber(T_Day) = False Then Exit Function
    
    If T_YearMonth = #12/31/1899# Then
        Min_Day = Min
        Max_Day = Max
    Else
        Min_Day = 1
        Max_Day = Date_Day_Last(T_YearMonth)
    End If
    
    If Min_Day <= CLng(T_Day) And CLng(T_Day) <= Max_Day Then
        Date_IsDay = True
    End If
    
End Function

Public Function Date_IsHour(T_Hour As Variant, Optional Min As Long = 0, Optional Max As Long = 23) As Boolean
    
    If isNumber(T_Hour) = False Then Exit Function
    
    If Min <= T_Hour And T_Hour <= Max Then
        Date_IsHour = True
    End If
    
End Function

Public Function Date_IsMinute(T_Minute As Variant, Optional Min As Long = 0, Optional Max As Long = 59) As Boolean
    
    If isNumber(T_Minute) = False Then Exit Function
    
    If Min <= T_Minute And T_Minute <= Max Then
        Date_IsMinute = True
    End If
    
End Function

Public Function Date_IsSecond(T_Second As Variant, Optional Min As Long = 0, Optional Max As Long = 59) As Boolean
    
    If isNumber(T_Second) = False Then Exit Function
    
    If Min <= T_Second And T_Second <= Max Then
        Date_IsSecond = True
    End If
    
End Function

Public Function Date_IsDateTime(Expression As Variant, Optional Min_Year As Long = 2000, Optional Max_Year As Long = 2099) As Boolean
    
    Dim Date_Str        As String
    Dim SplitAry()      As String
    Dim Flg_DateTime    As Boolean
    Dim T_Date_Sp       As String
    Dim T_Time_Sp       As String
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, " ") <> 0 Then
        
        SplitAry = Split(Date_Str, " ")
        
        If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 2 Then
            
            T_Date_Sp = SplitAry(0)
            T_Time_Sp = SplitAry(1)
            
            If Date_IsDate(T_Date_Sp, Min_Year, Max_Year) = True Then
                
                If Date_IsTime(T_Time_Sp) = True Then
                    
                    Flg_DateTime = True
                    
                End If
                
            End If
            
        End If
                
    End If
    
    Date_IsDateTime = Flg_DateTime
    
End Function

Public Function Date_IsDate(Expression As Variant, Optional Min_Year As Long = 2000, Optional Max_Year As Long = 2099, Optional Delimiter As String = "/") As Boolean
    
    Dim Date_Str    As String
    Dim T_Year      As Variant
    Dim T_Month     As Variant
    Dim T_Day       As Variant
    Dim SplitAry()  As String
    Dim Date_YM     As Date
    Dim Flg_Date    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, ":") = 0 Then
        
        If InStr(1, Date_Str, Delimiter) <> 0 Then
            
            SplitAry = Split(Date_Str, Delimiter)
            
            If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 3 Then
                
                T_Year = SplitAry(0)
                T_Month = SplitAry(1)
                T_Day = SplitAry(2)
                
                If Date_IsYear(T_Year, Min_Year, Max_Year) = True Then
                    If Date_IsMonth(T_Month) = True Then
                        Date_YM = DateSerial(T_Year, T_Month, 1)
                        If Date_IsDay(T_Day, T_YearMonth:=Date_YM) = True Then
                            
                            Flg_Date = True
                            
                        End If
                    End If
                End If
                
            End If
            
        End If
        
    End If
    
    Date_IsDate = Flg_Date
    
End Function

Public Function Date_IsDate_YearMonth(Expression As Variant, Optional Min_Year As Long = 2000, Optional Max_Year As Long = 2099, Optional Delimiter As String = "/") As Boolean
    
    Dim Date_Str    As String
    Dim T_Year      As Variant
    Dim T_Month     As Variant
    Dim SplitAry()  As String
    Dim Flg_Date    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, Delimiter) <> 0 Then
        
        SplitAry = Split(Date_Str, Delimiter)
        
        If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 2 Then
            
            T_Year = SplitAry(0)
            T_Month = SplitAry(1)
            
            If Date_IsYear(T_Year, Min_Year, Max_Year) = True Then
                If Date_IsMonth(T_Month) = True Then
                        
                    Flg_Date = True
                    
                End If
            End If
            
        End If
        
    End If
    
    Date_IsDate_YearMonth = Flg_Date
    
End Function

Public Function Date_IsDate_MonthDay(Expression As Variant, Optional Delimiter As String = "/") As Boolean
    
    Dim Date_Str    As String
    Dim T_Month     As Variant
    Dim T_Day       As Variant
    Dim SplitAry()  As String
    Dim Flg_Date    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, Delimiter) <> 0 Then
        
        SplitAry = Split(Date_Str, Delimiter)
        
        If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 2 Then
            
            T_Month = SplitAry(0)
            T_Day = SplitAry(1)
            
            If Date_IsMonth(T_Month) = True Then
                If Date_IsDay(T_Day) = True Then
                    
                    Flg_Date = True
                    
                End If
            End If
            
        End If
        
    End If
    
    Date_IsDate_MonthDay = Flg_Date
    
End Function

Public Function Date_IsTime(Expression As Variant, Optional Delimiter As String = "/") As Boolean
    
    Dim Date_Str    As String
    Dim SplitAry()  As String
    Dim T_Hour      As Variant
    Dim T_Minute    As Variant
    Dim T_Second    As Variant
    Dim Flg_Time    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, ":") <> 0 Then
        
        If InStr(1, Date_Str, Delimiter) = 0 Then
            
            SplitAry = Split(Date_Str, ":")
            
            If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 3 Then
                
                T_Hour = SplitAry(0)
                T_Minute = SplitAry(1)
                T_Second = SplitAry(2)
                    
                If Date_IsHour(T_Hour) = True Then
                    If Date_IsMinute(T_Minute) = True Then
                        If Date_IsSecond(T_Second) = True Then
                            
                            Flg_Time = True
                            
                        End If
                    End If
                End If
                
            End If
            
        End If
        
    End If
    
    Date_IsTime = Flg_Time
    
End Function

Public Function Date_IsTime_HourMinute(Expression As Variant) As Boolean
        
    Dim Date_Str    As String
    Dim SplitAry()  As String
    Dim T_Hour      As Variant
    Dim T_Minute    As Variant
    Dim Flg_Time    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, ":") <> 0 Then
        
        SplitAry = Split(Date_Str, ":")
        
        If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 2 Then
            
            T_Hour = SplitAry(0)
            T_Minute = SplitAry(1)
                
            If Date_IsHour(T_Hour) = True Then
                If Date_IsMinute(T_Minute) = True Then
                        
                    Flg_Time = True
                    
                End If
            End If
            
        End If
        
    End If
    
    Date_IsTime_HourMinute = Flg_Time
    
End Function

Public Function Date_IsTime_HourMinute_Lite(Expression As Variant) As Boolean
'+ Date型に時刻を格納可能と判定
'+ 12/1,12.1でも処理はされてしまうので、注意が必要
    
    Dim T_Time      As Date
    Dim Flg_Time    As Boolean
    
    '- 判定初期化：True
    Flg_Time = True
    
    '- 時刻として取得してみる
    T_Time = -1
    On Error Resume Next
    T_Time = TimeValue(Expression)
    On Error GoTo 0
    
    '- 取得できなかった場合、エラー
    If T_Time = -1 Then
        Flg_Time = False
    End If
    
    Date_IsTime_HourMinute_Lite = Flg_Time
    
End Function

Public Function Date_IsTime_MinuteSecond(Expression As Variant) As Boolean
    
    Dim Date_Str    As String
    Dim SplitAry()  As String
    Dim T_Minute    As Variant
    Dim T_Second    As Variant
    Dim Flg_Time    As Boolean
    
    If IsDate(Expression) = False Then Exit Function
    
    Date_Str = LCase(CStr(Expression))
    
    If InStr(1, Date_Str, ":") <> 0 Then
        
        SplitAry = Split(Date_Str, ":")
        
        If UBound(SplitAry, 1) - LBound(SplitAry, 1) + 1 = 2 Then
            
            T_Minute = SplitAry(0)
            T_Second = SplitAry(1)
            
            If Date_IsMinute(T_Minute) = True Then
                If Date_IsSecond(T_Second) = True Then
                    
                    Flg_Time = True
                    
                End If
            End If
            
        End If
        
    End If
    
    Date_IsTime_MinuteSecond = Flg_Time
    
End Function

Private Function IsNumber(Expression As Variant) As Boolean
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
