Attribute VB_Name = "M_Git"
Option Explicit

'Reference : Microsoft Scripting Runtime
'Reference : Microsoft ActiveX Data Objects 2.x Library

Private FSO     As New Scripting.FileSystemObject

Public Enum E_Encode
    Unknown
    ShiftJIS
    UTF8
    UTF16
End Enum

Public Function Git_ConvertUTF8()
    
    Dim Path_Repository As String
    Dim T_File          As File
    Dim Path_File       As String
    
    Path_Repository = "H:\00_VBA\04_色々\ExcelVBA\"
    
    If FSO.FolderExists(Path_Repository) = False Then Exit Function
    
    For Each T_File In FSO.GetFolder(Path_Repository).Files
        
        Path_File = T_File.Path
        
        Select Case FSO.GetExtensionName(Path_File)
            
            Case "bas", "cls", "txt"
            
                Call ShiftJIS_to_UTF8(Path_File)
                
        End Select
        
    Next
    
End Function

Private Function ShiftJIS_to_UTF8(File As String)
    
    Dim destWithBOM As ADODB.Stream
    Dim Cnt_Bom     As Long
    Dim T_Type      As E_Encode
    Dim T_Char      As String
    Dim T_Char_UTF8 As String
    
    T_Char_UTF8 = File_BOM_Type_to_Name(UTF8)
    
    Set destWithBOM = New ADODB.Stream
    With destWithBOM
        
        T_Type = File_Encode(File)
        
        If T_Type <> UTF8 Then
            
            '- utf-8の器を作っておく
            .Type = adTypeText
            .Charset = T_Char_UTF8
            .Open
            .Position = 0
            
            ' ファイルをS-JIS で開いてセット
            T_Char = File_BOM_Type_to_Name(T_Type)
            Dim TempStream  As ADODB.Stream
            Set TempStream = New ADODB.Stream
            With TempStream
                .Type = adTypeText
                .Charset = T_Char
                .Open
                Call .LoadFromFile(File)
                Call .CopyTo(destWithBOM)
                .Close
            End With
            
            '- 位置を頭に移動させておく
            '+ Copy後で下になっているので、そのまま保存しようとすると空のテキストになってしまう
            .Position = 0
            
            '- UTF8で保存
            Dim dest  As ADODB.Stream
            Set dest = New ADODB.Stream
            With dest
                .Type = adTypeBinary
                .Open
                Call destWithBOM.CopyTo(dest)
                Call .SaveToFile(File, adSaveCreateOverWrite)
                .Close
            End With
            
            .Close
            
        End If
        
'        '- BOM消去
'        '+ BOMが無いとShiftJIS→UTF8が再度繰り返され、エンコードがおかしくなるのでBOMは残しておく
'        Call File_BOM_Delete(File)
        
    End With
    
    Set destWithBOM = Nothing
    Set dest = Nothing
    Set TempStream = Nothing
    
End Function

Private Function File_BOM_Type_to_Name(T_Type As E_Encode) As String
    
    Dim T_Name      As String
    
    Select Case T_Type
    Case UTF8:      T_Name = "utf-8"
    Case UTF16:     T_Name = "utf-16"
    Case ShiftJIS:  T_Name = "shift-jis"
    Case Else:      T_Name = "shift-jis"
    End Select
    
    File_BOM_Type_to_Name = T_Name
    
End Function

Public Function File_BOM_Delete(File As String)
    
    Dim C_Stream    As ADODB.Stream
    Dim dest        As ADODB.Stream
    Dim Cnt_Bom     As Long
    
    If FSO.FileExists(File) = False Then Exit Function
    
    Set C_Stream = New ADODB.Stream
    
    ' BOM消去
    With C_Stream
        
        Cnt_Bom = File_BOM_Count(File)
        
        If 0 < Cnt_Bom Then
            .Type = adTypeBinary
            .Open
            Call .LoadFromFile(File)
            .Position = Cnt_Bom
            
            Set dest = New ADODB.Stream
            With dest
                .Type = adTypeBinary
                .Open
                .Position = 0
                Call C_Stream.CopyTo(dest)
                Call .SaveToFile(File, adSaveCreateOverWrite)
                .Close
            End With
            
            .Close
            
        End If
        
    End With
    
    Set C_Stream = Nothing
    Set dest = Nothing
    
End Function

Private Function File_BOM_from_File(File As String) As String()
        
    Dim C_Stream    As ADODB.Stream
    Dim Head1       As String
    Dim Head2       As String
    Dim Head3       As String
    Dim HeadBom     As String
    Dim T_Type      As E_Encode
    
    If FSO.FileExists(File) = False Then Exit Function
    
    Set C_Stream = New ADODB.Stream
        
    With C_Stream
        
        '- BOMの有無を確認
        .Type = adTypeBinary
        .Open
        Call .LoadFromFile(File)
        .Position = 0 '読込開始位置
        On Error Resume Next
        Head1 = UCase(Right("0" & Hex(AscB(.Read(1))), 2))
        Head2 = UCase(Right("0" & Hex(AscB(.Read(1))), 2))
        Head3 = UCase(Right("0" & Hex(AscB(.Read(1))), 2))
        On Error GoTo 0
        .Close
        
    End With
    
    T_Type = File_BOM_to_EncodeType(Head1, Head2, Head3)
    
    Select Case T_Type
    Case UTF8:  HeadBom = Head1 & "," & Head2 & "," & Head3
    Case UTF16: HeadBom = Head1 & "," & Head2 & "," & ""
    Case Else:  HeadBom = ",,,"
    End Select
    
    File_BOM_from_File = Split(HeadBom, ",")
    
    Set C_Stream = Nothing
    
End Function

Private Function File_Encode(File As String) As E_Encode
    
    Dim BomAry()    As String
    Dim T_Type      As E_Encode
    
    If FSO.FileExists(File) = False Then Exit Function
    
    BomAry = File_BOM_from_File(File)
    
    T_Type = File_BOM_to_EncodeType(BomAry(0), BomAry(1), BomAry(2))
    
    File_Encode = T_Type
    
End Function

Private Function File_BOM_Count(File As String) As Long
    
    Dim BomAry()    As String
    Dim Cnt_Bom     As Long
    Dim i           As Long
        
    If FSO.FileExists(File) = False Then Exit Function
    
    BomAry = File_BOM_from_File(File)
    
    Cnt_Bom = 0
    For i = LBound(BomAry, 1) To UBound(BomAry, 1)
        If BomAry(i) <> "" Then
            Cnt_Bom = Cnt_Bom + 1
        End If
    Next
    
    File_BOM_Count = Cnt_Bom
    
End Function

Private Function File_BOM_to_EncodeType(Head1 As String, Head2 As String, Head3 As String) As E_Encode
    
    Dim T_Type  As E_Encode
    
    'utf-16
    If (Head1 = "FF" And Head2 = "FE") Or _
       (Head1 = "FE" And Head2 = "FF") Then
        
        T_Type = UTF16
    
    'utf-8
    ElseIf (Head1 = "EF" And _
            Head2 = "BB" And _
            Head3 = "BF") Then
        
        T_Type = UTF8
        
    'BOM無し
    Else
        T_Type = Unknown
        
    End If
    
    File_BOM_to_EncodeType = T_Type
    
End Function
