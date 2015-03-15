Attribute VB_Name = "M_Math"
Option Explicit

Public Function Math_Round(num As Double, Optional At_Round As Long = 0) As Double
'+ Int,Fixは負数で小数点以下の切り方が異なる
'+ Int:負･･･常に切り上げ
'+ Fix:負･･･常に切り捨て
'+ このRoudは、絶対値として四捨五入する
    
    Math_Round = Application.WorksheetFunction.Round(num, At_Round)
    
End Function

Public Function Math_RoundDown(num As Double, Optional At_Round As Long = 0) As Double
    
    Math_RoundDown = Application.WorksheetFunction.RoundDown(num, At_Round)
    
End Function

Public Function Math_RoundUp(num As Double, Optional At_Round As Long = 0) As Double
    
    Math_RoundUp = Application.WorksheetFunction.RoundUp(num, At_Round)
    
End Function

Public Function Math_Floor(num As Double, RndDownUnit As Double) As Double
'+ 指定の値単位で切り捨てる
    Math_Floor = Application.WorksheetFunction.Floor(num, RndDownUnit)
    
End Function

Public Function Math_Ceiling(num As Double, RndUpUnit As Double) As Double
'+ 指定の値単位で切り上げる
    
    Math_Ceiling = Application.WorksheetFunction.Ceiling(num, RndUpUnit)
    
End Function

Public Function Math_Upper(Num1 As Variant, Num2 As Variant) As Variant
    
    If IsNumeric(Num1) = False Then Exit Function
    If IsNumeric(Num2) = False Then Exit Function
    
    If Num1 < Num2 Then
        Math_Upper = Num2
    Else
        Math_Upper = Num1
    End If
    
End Function

Public Function Math_Lower(Num1 As Variant, Num2 As Variant) As Variant
    
    If IsNumeric(Num1) = False Then Exit Function
    If IsNumeric(Num2) = False Then Exit Function
    
    If Num1 < Num2 Then
        Math_Lower = Num1
    Else
        Math_Lower = Num2
    End If
    
End Function

Public Function Math_Max(ParamArray Vals()) As Double
    
    Math_Max = Application.WorksheetFunction.Max(Vals)
    
End Function

Public Function Math_Min(ParamArray Vals()) As Double
    
    Math_Min = Application.WorksheetFunction.Min(Vals)
    
End Function

Public Function Math_Abs(num As Variant) As Variant
'+ 絶対値

    If IsNumeric(num) = False Then Exit Function
    
    Math_Abs = Math.Abs(num)
    
End Function

Public Function Math_Sgn(num As Variant) As Long
'+ 符号判定(正：1、負：-1、0：0)
    
    If IsNumeric(num) = False Then Exit Function
    
    Math_Sgn = Math.Sgn(num)
    
End Function

Public Function Math_Exp(num As Double) As Double
'+ eを底とする数式のべき乗（指数関数）を計算します
    
    If 709 < num Then Exit Function
    
    Math_Exp = Math.Exp(num)
    
End Function

Public Function Math_Log(num As Double) As Double
    
    Math_Log = Math.Log(num)
    
End Function

Public Function Math_Rnd(num As Double) As Single
    
    Math_Rnd = Math.Rnd(num)
    
End Function

Public Function Math_Randmize()
    
    Call Math.Randomize(Math.Rnd)
    
End Function

Public Function Math_Pow(Num1 As Double, Num2 As Double) As Double
'+ べき乗
    
    Math_Pow = Application.WorksheetFunction.Power(Num1, Num2)
    
End Function

Public Function Math_Sqr(num As Double) As Double
'+ 平方根
    
    Math_Sqr = Math.Sqr(num)
    
End Function

Public Function Math_Sin(num As Double) As Double
'+ サイン
    
    Math_Sin = Math.Sin(num)
    
End Function

Public Function Math_Cos(num As Double) As Double
'+ コサイン
    
    Math_Cos = Math.Cos(num)
    
End Function

Public Function Math_Tan(num As Double) As Double
'+ タンジェント
    
    Math_Tan = Math.Tan(num)
    
End Function

Public Function Math_Quotient_Max(Num1 As Double, Num2 As Double) As Double
'+ 除算の最大値を返す
    
    Dim Num_Quo     As Long
    Dim Num_Mod     As Long
    Dim Cnt_Quo     As Long
    
    If Num1 = 0 Or Num2 = 0 Then Exit Function
    
    Cnt_Quo = 0
    Do
        
        Cnt_Quo = Cnt_Quo + 1
        
        Num_Mod = Num1 Mod (Num2 * Cnt_Quo)
        
        Num_Quo = Num1 \ (Num2 * Cnt_Quo)
        
    Loop Until Num_Mod <= Num2
    
    Math_Quotient_Max = Num_Quo
    
End Function

Public Function Math_GCD(Num1 As Double, Num2 As Double) As Double
'+ 最大公約数
    
    Dim Num_Mod     As Double
    Dim Mod_Max     As Double
    
    If Num1 = 0 Or Num2 = 0 Then Exit Function
    
    Num_Mod = Num1 Mod Num2
    
    If Num_Mod = 0 Then
        Mod_Max = Num2
    Else
        Mod_Max = Math_GCD(Num2, Num_Mod)
    End If
    
    Math_GCD = Mod_Max
    
End Function

