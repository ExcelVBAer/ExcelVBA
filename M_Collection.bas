Attribute VB_Name = "M_Collection"
Option Explicit

Public Function Collection_to_Array(T_Colect As Collection) As Variant
    
    Dim T_Dim       As Long
    Dim DataAry     As Variant
    Dim LoopItem    As Variant
    Dim LoopItem2   As Variant
    Dim T_Colect2   As Collection
    Dim i           As Long
    Dim j           As Long
    Dim Cnt_Col     As Long
    Dim Cnt_Col2    As Long
    
    If T_Colect Is Nothing Then Exit Function
    
    Cnt_Col = T_Colect.Count
    If Cnt_Col = 0 Then Exit Function
    
    If TypeName(T_Colect.Item(1)) <> "Collection" Then
        T_Dim = 1
    Else
        If TypeName(T_Colect.Item(1).Item(1)) <> "Collection" Then
            T_Dim = 2
        End If
    End If
    
    Select Case T_Dim
    
    Case 1
        
        ReDim DataAry(0 To Cnt_Col - 1)
        
        i = 0
        For Each LoopItem In T_Colect
            DataAry(i) = LoopItem
            i = i + 1
        Next
        
    Case 2
        
        '- óÒêîÇéÊìæ
        For Each LoopItem In T_Colect
            If LoopItem Is Nothing = False Then
                Cnt_Col2 = LoopItem.Count
                Exit For
            End If
        Next
        If Cnt_Col2 = 0 Then
            Cnt_Col2 = 1
        End If
            
        ReDim DataAry(0 To Cnt_Col - 1, 0 To Cnt_Col2 - 1)
        
        i = 0
        For Each LoopItem In T_Colect
            
            j = 0
            For Each LoopItem2 In LoopItem
                
                If (Cnt_Col2 - 1) < j Then
                    ReDim Preserve DataAry(0 To Cnt_Col - 1, 0 To j)
                End If
                
                DataAry(i, j) = LoopItem2
                
                j = j + 1
                
            Next
            
            i = i + 1
            
        Next
            
    End Select
    
    Collection_to_Array = DataAry
    
End Function
