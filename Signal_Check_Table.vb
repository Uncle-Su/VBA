Public Type DICT_TYPE
    Name As String
    Col As Byte
End Type

Dim Row As Long
Sub Macro1()
    Dim Fso As Object, p$, sFolder$
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Dim dict() As DICT_TYPE
    p = "D:\WORK-STEC\Software Repository\SAIC_BCM_BC03_SW_Latest\src"
    Row = 0
    
    Call Capture_Variable("D:\WORK-STEC\Software Repository\SAIC_BCM_BC03_SW_Latest\src\GEN_COMMON\OS_IO_Debounce_Generated.cfg", "Input", "Output")
    Call Capture_Variable("D:\WORK-STEC\Software Repository\SAIC_BCM_BC03\03-SW\src\OS_CFG\OS_IO_CAN_LIN_Table.cfg", "Net_Input", "Net_Output")
    Call Capture_Variable("D:\WORK-STEC\Software Repository\SAIC_BCM_BC03_SW_Latest\src\GEN_COMMON\OS_IO_Internal_Generated.cfg", "Int_Input", "Int_Output")
    'ReDim dict(4 To 4) As DICT_TYPE
    Call Init_Col(dict)
'    sFile = ".cfg"
'    Call GetFolder(p, sFile, Fso, sFolder)

    sFile = ".c"
    Call GetFolder(p, sFile, Fso, sFolder, dict)


'    sFile = ".h"
'    Call GetFolder(p, sFile, Fso, sFolder)
    
'    If sFolder <> "" Then
'    Else
'    End If
    Set Fso = Nothing
End Sub
Private Sub GetFolder(ByVal sPath$, ByVal sFile$, Fso As Object, sFolder$, dic() As DICT_TYPE)
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    
    Dim GetStr As String
    Dim Ignore As Boolean
    Dim k
    Dim kk
    Dim sh As Worksheet
    Dim code$
    Dim code_complete As Boolean
    Dim i As Long, j As Long
    Set sh = ActiveSheet
    Set Folder = Fso.GetFolder(sPath)
    For Each File In Folder.Files
        If InStrRev(File.Name, ".") > 0 Then
            'Debug.Print LCase(Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".") + 1))
            'k = Split(File.Name, ".")
            'k (UBound(k))  '   is the suffix name
            If LCase(Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".") + 1)) = LCase(sFile) Then
                'Debug.Print sPath + "\" + File.Name
                
                'here is result
                Open sPath + "\" + File.Name For Input As #10
                Ignore = False
                Do Until EOF(10)
                    Line Input #10, GetStr
                    
                    If GetStr <> "" Then
                        If Not ingore Then
                            k = Split(GetStr, "//")
                            
                            If InStr(k(0), ";") > 0 Then
                                code_complete = True
                            Else
                                code_complete = False
                            End If
                            
'                            If code_complete Then
'                                code = k(0)
'                            Else
                                code = code + k(0)
'                            End If
                            
                            
                            If code_complete Then
                                If InStr(code, "OS_IO_Get_InputState(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_InputState(", 1)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(1, i) And sh.Cells(j, 2) <> "" Then
                                                
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "1:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                                
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_InputState(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_InputState(", 2)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "2:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_InputStateBits(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_InputStateBits(", 2)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(1, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                kk(2, i) = Replace(kk(2, i), "0x", "&h")
                                                kk(2, i) = Replace(kk(2, i), "0X", "&h")
                                                kk(2, i) = CStr(CLng(kk(2, i)))
                                                If CStr(sh.Cells(j, 3)) <> kk(2, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "3:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(2, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_InputStateBits(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_InputStateBits(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                
                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr(sh.Cells(j, 3)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    
                                                    Debug.Print "4:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_InputStateByte(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_InputStateByte(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
'                                                Dim u_temp%
'                                                u_temp = (1 + (sh.Cells(j, 3) - 1) \ 8)
'                                                If CStr((1 + (sh.Cells(j, 3) - 1) \ 8)) <> kk(3, i) Then
'                                                        Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
'                                                End If

                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr((1 + (sh.Cells(j, 3) - 1) \ 8)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "5:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_InputStateByte(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_InputStateByte(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr((1 + (sh.Cells(j, 3) - 1) \ 8)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "6:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_OutputState(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_OutputState(", 1)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(1, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "11:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_OutputState(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_OutputState(", 2)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "12:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_OutputStateBits(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_OutputStateBits(", 2)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(1, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                kk(2, i) = Replace(kk(2, i), "0x", "&h")
                                                kk(2, i) = Replace(kk(2, i), "0X", "&h")
                                                kk(2, i) = CStr(CLng(kk(2, i)))
                                                If CStr(sh.Cells(j, 3)) <> kk(2, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "13:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(2, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_OutputStateBits(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_OutputStateBits(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr(sh.Cells(j, 3)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    
                                                    Debug.Print "14:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_OutputStateByte(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_OutputStateByte(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr((1 + (sh.Cells(j, 3) - 1) \ 8)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "15:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Set_OutputStateByte(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Set_OutputStateByte(", 3)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                    
                                                kk(3, i) = Replace(kk(3, i), "0x", "&h")
                                                kk(3, i) = Replace(kk(3, i), "0X", "&h")
                                                kk(3, i) = Replace(kk(3, i), "WIPERDELAYPOSITIONS", "4")
                                                
                                                kk(3, i) = CStr(CLng(kk(3, i)))
                                                If CStr((1 + (sh.Cells(j, 3) - 1) \ 8)) <> kk(3, i) Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "16:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3)) & "  " & kk(3, i)
                                                End If
                                            End If
                                        Next
                                    Next
                                End If

                                If InStr(code, "OS_IO_Get_AD_Input_Range(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_AD_Input_Range(", 5)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(4, i) And sh.Cells(j, 2) <> "" Then
                                                sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "O"
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "21:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                                
                                            End If
                                        Next
                                    Next
                                End If

                                If InStr(code, "OS_IO_Get_EventOnIndex(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_EventOnIndex(", 4)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "22:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_EventOffIndex(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_EventOffIndex(", 4)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "23:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_Output_EventOnIndex(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_Output_EventOnIndex(", 4)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "24:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                If InStr(code, "OS_IO_Get_Output_EventOffIndex(") > 0 Then
                                    kk = Get_Augument(code, "OS_IO_Get_Output_EventOffIndex(", 4)
                                    For i = 1 To UBound(kk, 2)
                                        For j = 2 To sh.UsedRange.Rows.Count
                                            If sh.Cells(j, 2) = kk(2, i) And sh.Cells(j, 2) <> "" Then
                                                If sh.Cells(j, Get_Col_By_Name(File.Name, dic)) <> "O" Then
                                                    sh.Cells(j, Get_Col_By_Name(File.Name, dic)) = "I"
                                                End If
                                                If CStr(sh.Cells(j, 3)) <> "1" Then
                                                    Call Set_Yellow(sh.Name, j, Get_Col_By_Name(File.Name, dic))
                                                    Debug.Print "25:" & sh.Cells(j, 2) & ":" & CStr(sh.Cells(j, 3))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                                
                                code = ""
                            End If
                        End If
                        If InStr(GetStr, "/*") > 0 Then
                            Ignore = True
                        End If
                        If InStr(GetStr, "*/") > 0 Then
                            Ignore = False
                        End If
                    End If
                Loop
                Close #10
                
                'End of result
            End If
        End If
    Next
    For Each SubFolder In Folder.SubFolders
        Call GetFolder(SubFolder.Path, sFile, Fso, sFolder, dic)
    Next
    Set Folder = Nothing
    Set File = Nothing
    Set SubFolder = Nothing
End Sub
Sub test()
'    Open ThisWorkbook.Path + "\app.c" For Input As #4
'
'    Do Until EOF(4)
'        Line Input #4, GetStr
'        Debug.Print GetStr
'        Debug.Print Trim(GetStr)
'    Loop
'
'    Close #4
'    Debug.Print Trim("      fasf        ")
'    ss = " "
'    k = Split(ss, "//")
'    Debug.Print InStr(4, ss, "/")
'    Dim a() As String
'
'    ReDim a(0 To 5, 0 To 5) As String
'     a(0, 1) = 6
'    Debug.Print a(0, 1)
'    s1$ = "a(a(1,2,3),c,a(5,6,7))"
'    s2$ = "a"
'    k = Get_Augument(s1, s2, 3)
'    Debug.Print k(1, 1)
'    Debug.Print Trim("      fasf        ")
    kk = "0X08"
    Debug.Print CStr(CLng(Replace(kk, "0x", "&h")))
End Sub
Function Get_Augument(str As String, fun As String, aug_num As Long)
    Dim i%, j%
    Dim i_au%, i_num%, i_fun%
    Dim one$
    Dim found As Boolean
    Dim result() As String
    Dim pos1$

    i_aug = 0
    i_fun = 0
    
    i = 0
    i = InStr(1, str, fun)
    Do While i > 0
        found = False
        i_fun = i_fun + 1
        i_aug = 0
        ReDim Preserve result(1 To aug_num, 1 To i_fun) As String
        For j = i To Len(str)
            one = Mid(str, j, 1)
            If one = "(" Then
                i_num = i_num + 1
                found = True
            End If
            
            If i_num = 1 And (one = "," Or one = ")") Then
                pos2 = j
                i_aug = i_aug + 1
                result(i_aug, i_fun) = Trim(Replace(Mid(str, pos1 + 1, pos2 - pos1 - 1), vbTab, ""))
            End If
            
            
            If i_num = 1 And (one = "," Or one = "(") Then
                pos1 = j
            End If
            
            If one = ")" Then
                i_num = i_num - 1
            End If
            
            If i_num = 0 And found = True Then
                Exit For
            End If
        Next
    
        i = InStr(i + 1, str, fun)
    Loop
    
    Get_Augument = result
    
End Function

Sub Capture_Variable(Path As String, str1$, str2$)
    Dim GetStr As String, tmpstr As String
    Dim start1 As Boolean, end1 As Boolean
    Dim start2 As Boolean, end2 As Boolean
    Dim start As Long
    Dim CurStr$, CurLen%
    
    CurStr = ""
    CurLen = 0
    Open Path For Input As #10
    Do Until EOF(10)
        Line Input #10, GetStr
        If Left(GetStr, 2) <> "//" And InStr(GetStr, "=") <= 0 And InStr(GetStr, "LAST_NONE_DEBOUNCED") <= 0 Then
            If InStr(GetStr, "{") > 0 Or (InStr(GetStr, "#define ") And InStr(GetStr, "\") > 0) > 0 Then
                start = start + 1
                Row = Row + 1
                Select Case start
                    Case 1
                        Cells(Row + 1, 1) = str1
                    Case 2
                        Cells(Row + 1, 1) = str2
                End Select
            End If
            
            tmpstr = Get_Mid_Str(GetStr, "^", ",")
            If tmpstr <> "" Then
                If InStr(LCase(tmpstr), "reserved") <= 0 Then
                
                    If CurStr = "" Then
                        CurStr = tmpstr
                        CurLen = 1
                    ElseIf Check_Exist(tmpstr, CurStr & "_[1-9]") Then
                        CurLen = CurLen + 1
                    Else
                        Row = Row + 1
                        Cells(Row, 2) = CurStr
                        Cells(Row, 3) = CurLen
                        
                        'new signal
                        CurStr = tmpstr
                        CurLen = 1
                    End If
                End If
            End If
        End If
    Loop
    Close #10
End Sub
Function Get_Col_By_Name(Name As String, dic() As DICT_TYPE) As Long
    Dim i%
    Dim sh As Worksheet
    
    Set sh = ActiveSheet
    
    If Name <> "" And InStr(Name, " ") <> 1 Then
        If UBound(dic) = 4 And dic(UBound(dic)).Name = "" Then
            ReDim dic(4 To 4) As DICT_TYPE
            dic(UBound(dic)).Name = Name
            dic(UBound(dic)).Col = 4
            sh.Cells(1, dic(UBound(dic)).Col) = Name
            Get_Col_By_Name = dic(UBound(dic)).Col
            Exit Function
            
        ElseIf UBound(dic) >= 4 Then
            For i = 4 To UBound(dic)
                If dic(i).Name = Name Then
                    Get_Col_By_Name = dic(i).Col
                    Exit Function
                End If
            Next
            
            ReDim Preserve dic(4 To UBound(dic) + 1) As DICT_TYPE
            dic(UBound(dic)).Name = Name
            dic(UBound(dic)).Col = UBound(dic)
            sh.Cells(1, dic(UBound(dic)).Col) = Name
            Get_Col_By_Name = dic(UBound(dic)).Col
            Exit Function
        
        End If

    Else
        Get_Col_By_Name = 0
    End If
    
End Function
Sub Init_Col(dic() As DICT_TYPE)
    Dim i%
    Dim sh As Worksheet
    Set sh = ActiveSheet
    ReDim dic(4 To 4) As DICT_TYPE
    If sh.UsedRange.Columns.Count >= 4 Then
        For i = 4 To sh.UsedRange.Columns.Count
            Call Get_Col_By_Name(sh.Cells(1, i), dic)
        Next
    End If
End Sub
Function IsReDim(ByRef MyArray()) As Boolean
    On Error GoTo Z
    Dim szTmp
    szTmp = Join(MyArray, ",")
    IsReDim = LenB(szTmp) > 0
    Exit Function
Z:
    IsReDim = False
End Function
Function Get_Mid_Str(ByVal Source As String, ByVal PreStr As String, ByVal PostStr As String) As String
    Dim Temp_Reg As New RegExp
    With Temp_Reg
        .Global = True
        .IgnoreCase = True
    End With
    Dim Temp_MC As MatchCollection
    Dim Temp_Str As String
    
    Temp_Str = PreStr + "\w+" + PostStr
    Temp_Reg.Pattern = Temp_Str
    Set Temp_MC = Temp_Reg.Execute(Source)
    If Temp_MC.Count > 0 Then
        Temp_Str = Temp_MC.Item(0)
        
        Temp_Reg.Pattern = PreStr
        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Temp_Reg.Pattern = PostStr
        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Get_Mid_Str = Temp_Str
    Else
        Get_Mid_Str = ""
    End If

End Function
Function Get_End_Str(ByVal Source As String, ByVal PreStr As String) As String
    Dim Temp_Reg As New RegExp
    With Temp_Reg
        .Global = True
        .IgnoreCase = True
    End With
    Dim Temp_MC As MatchCollection
    Dim Temp_Str As String
    
    Temp_Str = PreStr + ".*"
    Temp_Reg.Pattern = Temp_Str
    Set Temp_MC = Temp_Reg.Execute(Source)
    If Temp_MC.Count > 0 Then
        Temp_Str = Temp_MC.Item(0)
        
        Temp_Reg.Pattern = PreStr
        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Get_End_Str = Temp_Str
    Else
        Get_End_Str = ""
    End If

End Function

Function Get_Mid_Str_S(ByVal Source As String, ByVal PreStr As String, ByVal PostStr As String) As String
    Dim Temp_Reg As New RegExp
    With Temp_Reg
        .Global = True
        .IgnoreCase = True
    End With
    Dim Temp_MC As MatchCollection
    Dim Temp_Str As String
    
    Temp_Str = PreStr + "\S+" + PostStr
    Temp_Reg.Pattern = Temp_Str
    Set Temp_MC = Temp_Reg.Execute(Source)
    If Temp_MC.Count > 0 Then
        Temp_Str = Temp_MC.Item(0)
        
        Temp_Reg.Pattern = PreStr
        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Temp_Reg.Pattern = PostStr
        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Get_Mid_Str_S = Temp_Str
    Else
        Get_Mid_Str_S = ""
    End If

End Function

Function Check_Exist(ByVal Source As String, ByVal PreStr As String) As Boolean
    Dim Temp_Reg As New RegExp
    With Temp_Reg
        .Global = True
        .IgnoreCase = True
    End With
    Dim Temp_MC As MatchCollection
    Dim Temp_Str As String
    
    Temp_Str = PreStr
    Temp_Reg.Pattern = Temp_Str
    Set Temp_MC = Temp_Reg.Execute(Source)
    If Temp_MC.Count > 0 Then
'        Temp_Str = Temp_MC.Item(0)
'
'        Temp_Reg.Pattern = PreStr
'        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
'
'        Temp_Reg.Pattern = PostStr
'        Temp_Str = Temp_Reg.Replace(Temp_Str, "")
        
        Check_Exist = True
    Else
        Check_Exist = False
    End If

End Function

Sub Set_Yellow(Sheetname As String, Row As Long, Column As Long)
    
    Sheets(Sheetname).Select
    Cells(Row, Column).Select
    'Set Yellow
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub Set_Blue(Sheetname As String, Row As Long, Column As Long)
    
    Sheets(Sheetname).Select
    Cells(Row, Column).Select
    'Set Blue
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

End Sub

Sub Set_Red(Sheetname As String, Row As Long, Column As Long)
    
    Sheets(Sheetname).Select
    Cells(Row, Column).Select
    'Set Red
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub Set_Gray(Sheetname As String, Row As Long, Column As Long)

    Sheets(Sheetname).Select
    Cells(Row, Column).Select
    'Set gray
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub

Sub Set_NonColor(Sheetname As String, Row As Long, Column As Long)
    
    Sheets(Sheetname).Select
    Cells(Row, Column).Select
    'Set Non Color
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Function Get_Path() As String
    Dim Path As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Function
        Path = .SelectedItems(1) & "\"
    End With
    
    Get_Path = Path
End Function
