Public Type SIGNAL_MAP
    SigName As String
    CanIn As String
    LinIn As String
    CanOut As String
    LinOut As String
    CanInFrm As String
    LinInFrm As String
    CanOutFrm As String
    LinOutFrm As String
End Type
Public Type CAN_LIN_MAP
    AppName As String
    DbcName As String
    FrmName As String
End Type
Public Type HARDWARE_MAP
    AppName As String
    HwName As String
End Type

Sub db()
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient  ' if this item doesn't set, recordcount will be -1
    Dim cmd As New ADODB.Command
    Dim i As Long        '

    cn.Provider = "Microsoft.ACE.OLEDB.12.0"        'access 2007/2010/2013 64bit version use this ace engine
                                                    ' maybe 32bit version sw can use "Microsoft.Jet.Oledb.4.0"
    cn.Open (ThisWorkbook.path & "\数据库\class.accdb")
    
    rs.ActiveConnection = cn
    'rs.CursorType = adOpenDynamic       'if adOpenStatic , recordcount can be used
    'rs.Open ("select * from sleep_sample"), , adOpenDynamic, adLockPessimistic
    rs.Open ("select * from sleep_sample,sleep_sample_re"), , adOpenDynamic, adLockPessimistic
    
    Debug.Print rs.RecordCount
    rs.MoveFirst
    Debug.Print rs.AbsolutePosition
    
    ActiveSheet.Cells.Clear
    For i = 1 To rs.Fields.Count
        Cells(i, 3) = rs.Fields.Item(0).Value
        Cells(i, 5) = rs.Fields.Item(1).Value
        Cells(i, 7) = rs.Fields.Item(2).Value
        Cells(i, 9) = rs.Fields.Item(3).Value
        
        Cells(i, 2) = rs(0).Value
        Debug.Print rs.AbsolutePosition
        rs.MoveNext
    Next
    rs.MoveFirst
    For i = 1 To rs.Fields.Count
        Cells(i, 4) = rs.Fields(0).Value
        Cells(i, 6) = rs.Fields(1).Value
        Cells(i, 8) = rs.Fields(2).Value
        Cells(i, 10) = rs.Fields(3).Value
        'Cells(i, 2) = rs(0).Value
        rs.MoveNext
    Next
    
    
    For i = 1 To rs.Fields.Count
         Cells(i, 1) = rs.Fields(i - 1).Name
         'i = i + 1
    Next
    
    Range("A2").CopyFromRecordset rs
    rs.Close
    cn.Close
    
    
'    If Not rs.EOF Then
'        MsgBox ("成绩已经存在于数据库中")
'    Else
'        For i = 3 To 18
'            rs.addnew
'            rs!Name = Sheet1.Cells(i, 1)
'            rs!chinese = Sheet1.Cells(i, 2)
'            rs!Math = Sheet1.Cells(i, 3)
'            rs!english = Sheet1.Cells(i, 4)
'            Sheet1.Cells(18, 2) = Sheet1.Cells(18, 2) + Sheet1.Cells(i, 2)
'            Sheet1.Cells(18, 3) = Sheet1.Cells(18, 3) + Sheet1.Cells(i, 3)
'            Sheet1.Cells(18, 3) = Sheet1.Cells(18, 4) + Sheet1.Cells(i, 4)
'
'            If i = 18 Then
'
'                rs!chinese = Sheet1.Cells(18, 2)
'                rs!english = Sheet1.Cells(18, 4)
'                rs!Math = Sheet1.Cells(18, 3)
'                rs!平均分 = CInt((rs!chinese + rs!english + rs!Math) / 3)
'                rs!总分 = rs!chinese + rs!english + rs!Math
'                Sheet1.Cells(18, 5) = rs!平均分
'            Else
'                rs!总分 = CInt(Sheet1.Cells(i, 2)) + CInt(Sheet1.Cells(i, 3)) + CInt(Sheet1.Cells(i, 4))
'                rs!平均分 = CInt(rs!总分 / 3)
'                Sheet1.Cells(i, 5) = rs!平均分
'            End If
'
'            rs.addnew
'            rs!Name = Sheet1.Cells(i, 7)
'            rs!chinese = Sheet1.Cells(i, 8)
'            rs!Math = Sheet1.Cells(i, 9)
'            rs!english = Sheet1.Cells(i, 10)
'
'            Sheet1.Cells(18, 8) = Sheet1.Cells(18, 8) + Sheet1.Cells(i, 8)
'            Sheet1.Cells(18, 9) = Sheet1.Cells(18, 9) + Sheet1.Cells(i, 9)
'            Sheet1.Cells(18, 10) = Sheet1.Cells(18, 10) + Sheet1.Cells(i, 10)
'
'            If i = 18 Then
'                rs!chinese = allchineseB
'                rs!english = allenglishB
'                rs!Math = allmathB
'                rs!平均分 = CInt((rs!chinese + rs!english + rs!Math) / 3)
'                Sheet1.Cells(18, 11) = rs!平均分
'                rs!总分 = rs!chinese + rs!english + rs!Math
'            Else
'                rs!总分 = CInt(Sheet1.Cells(i, 8)) + CInt(Sheet1.Cells(i, 9)) + CInt(Sheet1.Cells(i, 10))
'                rs!平均分 = CInt(rs!总分 / 3)
'                Sheet1.Cells(i, 11) = rs!平均分
'            End If
'            rs.Update
'
'            rs.addnew
'            rs!Name = Sheet1.Cells(i, 13)
'            rs!chinese = Sheet1.Cells(i, 14)
'            rs!Math = Sheet1.Cells(i, 15)
'            rs!english = Sheet1.Cells(i, 16)
'
'            Sheet1.Cells(18, 14) = Sheet1.Cells(18, 14) + Sheet1.Cells(i, 14)
'            Sheet1.Cells(18, 15) = Sheet1.Cells(18, 15) + Sheet1.Cells(i, 15)
'            Sheet1.Cells(18, 16) = Sheet1.Cells(18, 16) + Sheet1.Cells(i, 16)
'
'            If i = 18 Then
'                rs!chinese = allchineseC
'                rs!english = allenglishC
'                rs!Math = allmathC
'                rs!平均分 = CInt((rs!chinese + rs!english + rs!Math) / 3)
'                rs!总分 = rs!chinese + rs!english + rs!Math
'                Sheet1.Cells(i, 17) = rs!平均分
'            Else
'                rs!总分 = CInt(Sheet1.Cells(i, 14)) + CInt(Sheet1.Cells(i, 15)) + CInt(Sheet1.Cells(i, 16))
'                rs!平均分 = CInt(rs!总分 / 3)
'                Sheet1.Cells(i, 17) = rs!平均分
'            End If
'            rs.Update
'        Next
'
'        rs.Close
'        MsgBox ("成绩已经成功写入数据库")
'    End If
    
    Set cn = Nothing
    Set rs = Nothing
End Sub
Sub db2_excel()
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim Fpath As String, CnStr As String, SqlStr As String
    Dim sname() As String
    
    rs.CursorLocation = adUseClient  ' if this item doesn't set, recordcount will be -1
    Dim cmd As New ADODB.Command
    Dim i As Long
    
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"        'access 2007/2010/2013 64bit version use this ace engine
                                                    ' maybe 32bit version sw can use "Microsoft.Jet.Oledb.4.0"
    'cn.Properties("Extended  Properties").Value = "Excel 12.0"
    Fpath = ThisWorkbook.path & "\数据库\Excel_1.xlsx"
    CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & Fpath & ";" & _
            "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
    cn.ConnectionString = CnStr

    'cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Extended Properties=""Excel 12.0 Macro;HDR=NO"";" & "Data Source=" & ThisWorkbook.Path & "\数据库\Excel_1.xlsx;"
    
    cn.Open
    
    p = 0
    If True Then
        Set rs = cn.OpenSchema(adSchemaTables, TABLE_NAME)
            
            ActiveSheet.Cells.Clear
            For k = 1 To rs.Fields.Count
                Cells(p + 1, k) = rs.Fields(k - 1).Name
            Next
        Debug.Print rs.RecordCount
        Do Until rs.EOF
            Debug.Print rs!TABLE_NAME & "    ";   'read Table_Name(SheetName)   '.' is replaced by '#'
            Debug.Print rs!TABLE_TYPE & "    ";
            Debug.Print rs!TABLE_CATALOG & "    ";
            Debug.Print rs!TABLE_SCHEMA;
            Debug.Print
            Debug.Print rs(1) & "    ";
            Debug.Print rs(2) & "    ";
            Debug.Print rs(3) & "    ";
            Debug.Print rs(4) & "    ";
            Debug.Print rs(5) & "    ";
            Debug.Print rs(6) & "    ";
            Debug.Print rs(7)
            Debug.Print rs.Fields(2).Value
            
            p = p + 1
            Debug.Print p
            For k = 1 To rs.Fields.Count
                Cells(p + 2, k) = rs.Fields(k - 1).Value
            Next
            rs.MoveNext
        Loop
    Else
        rs.ActiveConnection = cn
        SqlStr = "Select * FROM [BC03_Locking System$]"
        rs.Open SqlStr, , adOpenDynamic, adLockPessimistic
    End If

    Debug.Print rs.RecordCount
    Debug.Print rs.Fields.Count
    rs.MoveFirst
    Debug.Print rs.AbsolutePosition
    'Sheets("Sheet1").Range("A1:G100").CopyFromRecordset cn.Execute(SqlStr)

    'rs.CursorType = adOpenDynamic       'if adOpenStatic , recordcount can be used
    
    ActiveSheet.Cells.Clear
    For i = 1 To rs.Fields.Count
        Cells(1, i) = rs.Fields.Item(0).Value
        Cells(3, i) = rs(0).Value
        Debug.Print rs.AbsolutePosition
        rs.MoveNext
    Next
    
    
    For i = 1 To rs.Fields.Count
         Cells(1, i) = rs.Fields(i - 1).Name
         'i = i + 1
    Next
    ActiveSheet.Cells.Clear
    Range("A2").CopyFromRecordset rs

    rs.Close
    cn.Close
        
    Set cn = Nothing
    Set rs = Nothing
End Sub

Sub 合并工作簿数据()
Dim cnn As New Connection, rs As New Recordset
Dim cn As Object
Dim pname As String
Dim strSql As String, myname As String
myname = Dir(ThisWorkbook.path & "\*.xlsm")
Range("A2:XFD1048576").Clear
While myname <> ""
pname = ThisWorkbook.path & "\" & myname '
If myname <> "合并结果.xlsm" Then
If myname <> "合并结果.xlsm" Then Workbooks.Open pname
k = k + 1

Range("A1").Select
Selection.End(xlDown).Select '要保证A列没有空单格才能这样使用
Selection.End(xlDown).Select
Selection.End(xlUp).Offset(1, 0).Select
lrow = Selection.row
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Extended Properties=Excel 12.0 Macro;" & "Data Source=" & pname
sname = GetSheetList(pname)
strSql = "Select 身份证号码,原姓名,乡编码,村编码,组编码,ColumnF,医疗证号,ColumnH,ColumnI,ColumnJ,联系电话,ColumnL,ColumnM,ColumnN,ColumnO,与户主关系 FROM [" & sname & "$]" ' Where 姓名 like '%" & str1 & "%'"
'strSql = "Select 身份证号码,原姓名,乡编码,村编码,组编码,医疗证号,联系电话,与户主关系 FROM [" & sname & "$]" ' Where 姓名 like '%" & str1 & "%'"
rs.Open strSql, cnn, adOpenStatic
Sheet2.Cells(lrow, 1).CopyFromRecordset rs
cnn.Close
'cn.close
End If
If myname <> "合并结果.xlsm" Then Workbooks(myname).Close True
myname = Dir
Wend
Set cnn = Nothing
Set cn = Nothing
End Sub
Function GetSheetList(strFile As String) As String
Dim adoxCatalog As ADOX.Catalog
Dim tbSheet As ADOX.Table
Dim i As Integer
ReDim SheetNum(10)
Set adoxCatalog = New ADOX.Catalog
GetSheetList = ""
i = 0
adoxCatalog.ActiveConnection = ("Provider=Microsoft.Jet.OLEDB.4.0;extended properties='Excel 8.0;HDR=No';data source=" & strFile)
For Each tbSheet In adoxCatalog.Tables
If tbSheet.Type = "TABLE" Then
SheetNum(i) = Left(tbSheet.Name, InStr(1, tbSheet.Name, "$") - 1)
i = i + 1
End If
Next
GetSheetList = Trim(SheetNum(0)) '只获取工作簿第一个工作表的名称
Set adoxCatalog = Nothing
End Function

Sub Get_Dif()
    Dim i&, j&, k&
    Dim sh As Worksheet, sh_o As Worksheet, val1 As Byte, val2 As Byte
    Set sh = Sheets("Sheet4")
    Set sh_o = Sheets("Sheet2")
    For i = 1 To 70
        For j = 75 To 165
            If sh.Cells(i, 2) & sh.Cells(i, 3) = sh.Cells(j, 2) Then
                For k = 5 To sh.Cells(i, 100).End(xlToLeft).Column
                    
                    sh_o.Cells(i + 1, 1) = sh.Cells(j, 2)
                    val1 = CByte("&h" & sh.Cells(i, k))
                    val2 = CByte("&h" & sh.Cells(j, k))
                    If val1 <> val2 Then
                        sh_o.Cells(i + 1, k) = Hex(val1 Xor val2)
                    End If
                    Debug.Print sh.Cells(i, k) & " ";
                Next
                Debug.Print
                Exit For
            End If
        Next
    Next
    sh_o.Cells(1, 1) = "ID"
    For i = 1 To 32
        sh_o.Cells(1, 4 + i) = "Byte" & i
    Next
End Sub

Public Function get_lib(InputPath As String, SigMap() As SIGNAL_MAP, Prm() As String, HwMap() As HARDWARE_MAP, Command() As String)

        DID_filename = "DID_info.xlsm"
        command_filename = "Command Lib.xlsx"
        Mapping_filename = "CAN_LIN_Generate_BC03.xlsm"
        hardware_filename = "Hardware_Mapping.xlsm"
    
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        rs.CursorLocation = adUseClient  ' if this item doesn't set, recordcount will be -1
    
        cn.Provider = "Microsoft.ACE.OLEDB.12.0"        'access 2007/2010/2013 64bit version use this ace engine
                                                    ' maybe 32bit version sw can use "Microsoft.Jet.Oledb.4.0"
        'cn.Properties("Extended  Properties").Value = "Excel 12.0"
    
    
    'start get did_info
        Input_FPath = InputPath + DID_filename
        Fpath = Input_FPath
        CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & Fpath & ";" & _
                "Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
                                                                         'HDR indicate first Row is Field Name
        cn.ConnectionString = CnStr
        cn.Open

        rs.ActiveConnection = cn
        rs.Open "select ParameterName from [DID_Table$]", , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表
        
        rs.MoveFirst
        j = 0
        For i = 0 To rs.RecordCount - 1
            If (Not IsNull(rs.Fields.Item("ParameterName").Value)) And _
                LCase(rs.Fields.Item("ParameterName").Value) <> "reserved" Then
                j = j + 1
                ReDim Preserve Prm(j - 1) As String
                Prm(j - 1) = rs.Fields.Item("ParameterName").Value
            End If
            rs.MoveNext
        Next
        
        rs.Close
        cn.Close
        
    'End Did_inf
    
   'start get command info
        Input_FPath = InputPath + command_filename
        Fpath = Input_FPath
        CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & Fpath & ";" & _
                "Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
                                                                         'HDR indicate first Row is Field Name
        cn.ConnectionString = CnStr
        cn.Open

        rs.ActiveConnection = cn
        rs.Open "select Command from [Command Pool$]", , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表
        
        rs.MoveFirst
        j = 0
        For i = 0 To rs.RecordCount - 1
            If (Not IsNull(rs.Fields.Item("Command").Value)) Then
                j = j + 1
                ReDim Preserve Command(j - 1) As String
                Command(j - 1) = rs.Fields.Item("Command").Value
            End If
            rs.MoveNext
        Next
        
        rs.Close
        cn.Close
        
    'End command info
    
    'start get hardware _info
        Input_FPath = InputPath + hardware_filename
        Fpath = Input_FPath
        CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & Fpath & ";" & _
                "Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
                                                                         'HDR indicate first Row is Field Name
        cn.ConnectionString = CnStr
        cn.Open

        rs.ActiveConnection = cn
        rs.Open "select [App Name],[Alias] from [Hardware$]", , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表
        
        
        rs.MoveFirst
        j = 0
        rs.MoveNext
        rs.MoveNext
        rs.MoveNext
        For i = 3 To rs.RecordCount - 1
            If Not IsNull(rs.Fields.Item("App Name").Value) Then
                
                ReDim Preserve HwMap(j) As HARDWARE_MAP
                HwMap(j).AppName = rs.Fields.Item("App Name").Value
                
                If Not IsNull(rs.Fields.Item("Alias").Value) Then
                    HwMap(j).HwName = rs.Fields.Item("Alias").Value
                Else
                    HwMap(j).HwName = ""
                End If
                j = j + 1
            End If
            

            rs.MoveNext
        Next
        rs.Close
        cn.Close
        
    'End hardware_info
        
    'start get network _info
        Input_FPath = InputPath + Mapping_filename
        Fpath = Input_FPath
        CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & Fpath & ";" & _
                "Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
                                                                         'HDR indicate first Row is Field Name
        cn.ConnectionString = CnStr
        cn.Open

        rs.ActiveConnection = cn
    'Get signal mapping
        sql_str = "select Name,[CAN Input],[LIN Input],[CAN Output],[LIN Output] from [Inputs$] where [CAN Input]<>'' or [LIN Input]<>'' or [CAN Output]<>'' or [LIN Output]<>'' " & _
                  "Union " & _
                  "select Name,[CAN Input],[LIN Input],[CAN Output],[LIN Output] from [Outputs$] where [CAN Input]<>'' or [LIN Input]<>'' or [CAN Output]<>'' or [LIN Output]<>''"
        rs.Open sql_str, , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表,同时检索多张表会慢很多
        
        
        ReDim Preserve SigMap(rs.RecordCount - 1) As SIGNAL_MAP
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            SigMap(i).SigName = rs.Fields.Item("Name").Value
            If Not IsNull(rs.Fields.Item("CAN Input").Value) Then
                SigMap(i).CanIn = rs.Fields.Item("CAN Input").Value
            End If
            If Not IsNull(rs.Fields.Item("LIN Input").Value) Then
                SigMap(i).LinIn = rs.Fields.Item("LIN Input").Value
            End If
            If Not IsNull(rs.Fields.Item("CAN Output").Value) Then
                SigMap(i).CanOut = rs.Fields.Item("CAN Output").Value
            End If
            If Not IsNull(rs.Fields.Item("LIN Output").Value) Then
                SigMap(i).LinOut = rs.Fields.Item("LIN Output").Value
            End If
            rs.MoveNext
        Next
        rs.Close
    'End of get signal mapping
        
    'get CAN_LIN mapping

        sql_str = "select [Signal Name],[DBC Name],[Frame Name] from [CAN_$] where [Signal Name] <>'' "
        'where [CAN Input]<>''  or [LIN Input]<>'' or [CAN Output]<>'' or [LIN Output]<>''        ',[Outputs$]
        rs.Open sql_str, , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表,同时检索多张表会慢很多
        
        Dim Can_Link() As CAN_LIN_MAP
        ReDim Preserve Can_Link(rs.RecordCount - 1) As CAN_LIN_MAP
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            Can_Link(i).AppName = rs.Fields.Item("Signal Name").Value
            If Not IsNull(rs.Fields.Item("DBC Name").Value) Then
                Can_Link(i).DbcName = rs.Fields.Item("DBC Name").Value
            End If
            Can_Link(i).FrmName = rs.Fields.Item("Frame Name").Value
            rs.MoveNext
        Next
        rs.Close
                
        sql_str = "select [Signal Name],[DBC Name],[Frame Name] from [LIN_$] where [Signal Name] <>''"
        'where [CAN Input]<>''  or [LIN Input]<>'' or [CAN Output]<>'' or [LIN Output]<>''        ',[Outputs$]
        rs.Open sql_str, , adOpenDynamic, adLockPessimistic     '"select * from [DID_Table$],[VT Config$]"可以检索多个表,同时检索多张表会慢很多
        
        Dim Lin_Link() As CAN_LIN_MAP
        ReDim Preserve Lin_Link(rs.RecordCount - 1) As CAN_LIN_MAP
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            Lin_Link(i).AppName = rs.Fields.Item("Signal Name").Value
            If Not IsNull(rs.Fields.Item("DBC Name").Value) Then
                Lin_Link(i).DbcName = rs.Fields.Item("DBC Name").Value
            End If
            Lin_Link(i).FrmName = rs.Fields.Item("Frame Name").Value
            rs.MoveNext
        Next
        rs.Close
        cn.Close
    'end get CAN_LIN mapping
        
    'do match
        For i = 0 To UBound(SigMap)
            If SigMap(i).CanIn <> "" Then
                Found = 0
                For j = 0 To UBound(Can_Link)
                    If Can_Link(j).AppName = SigMap(i).CanIn Then
                        SigMap(i).CanIn = Can_Link(j).DbcName
                        SigMap(i).CanInFrm = Can_Link(j).FrmName
                        Found = 1
                        Exit For
                    End If
                Next
                If Found = 0 Then
                    SigMap(i).CanIn = ""
                End If
            End If
            If SigMap(i).CanOut <> "" Then
                Found = 0
                For j = 0 To UBound(Can_Link)
                    If Can_Link(j).AppName = SigMap(i).CanOut Then
                        SigMap(i).CanOut = Can_Link(j).DbcName
                        SigMap(i).CanOutFrm = Can_Link(j).FrmName
                        Found = 1
                        Exit For
                    End If
                Next
                If Found = 0 Then
                    SigMap(i).CanOut = ""
                End If
            End If
            If SigMap(i).LinIn <> "" Then
                Found = 0
                For j = 0 To UBound(Lin_Link)
                    If Lin_Link(j).AppName = SigMap(i).LinIn Then
                        SigMap(i).LinIn = Lin_Link(j).DbcName
                        SigMap(i).LinInFrm = Lin_Link(j).FrmName
                        Found = 1
                        Exit For
                    End If
                Next
                If Found = 0 Then
                    SigMap(i).LinIn = ""
                End If
            End If
            If SigMap(i).LinOut <> "" Then
                Found = 0
                For j = 0 To UBound(Lin_Link)
                    If Lin_Link(j).AppName = SigMap(i).LinOut Then
                        SigMap(i).LinOut = Lin_Link(j).DbcName
                        SigMap(i).LinOutFrm = Lin_Link(j).FrmName
                        Found = 1
                        Exit For
                    End If
                Next
                If Found = 0 Then
                    SigMap(i).LinOut = ""
                End If
            End If
        Next
    'end do match

    'End network_info
End Function

Public Sub Testcase_Convert()
    Dim path As String
    Dim SigMap_t() As SIGNAL_MAP
    Dim Prm_t() As String
    Dim HwMap_t() As HARDWARE_MAP
    Dim Cmd_t() As String
    
    Dim LIN_map(8) As String
    
    LIN_map(0) = ""      ' index = DBC channel,  value = Hardware channel
    LIN_map(1) = "1"     ' index = DBC channel,  value = Hardware channel
    LIN_map(2) = "5"     ' index = DBC channel,  value = Hardware channel
    LIN_map(3) = "3"     ' index = DBC channel,  value = Hardware channel
    LIN_map(4) = ""      ' index = DBC channel,  value = Hardware channel
    LIN_map(5) = ""      ' index = DBC channel,  value = Hardware channel
    LIN_map(6) = ""      ' index = DBC channel,  value = Hardware channel
    LIN_map(7) = ""      ' index = DBC channel,  value = Hardware channel
    LIN_map(8) = "6"     ' index = DBC channel,  value = Hardware channel
    
    
    
    result = MsgBox("请选择数据库目录", vbYesNo, "操作提示")
    'vbOK=1,vbCancel=2,vbAbort=3,vbRetry=4,vbIgnore=5,vbYes=6,vbNo=7
    If result = vbNo Then
        Exit Sub
    End If
        
    path = Get_Path()
    If path = "" Then
        MsgBox "No path is chosen!"
        Exit Sub
    End If
    
    result = MsgBox("请选择Testcase文件", vbYesNo, "操作提示")
    If result = vbNo Then
        Exit Sub
    End If
    FileDir = SelectFile()
    If FileDir = "" Then
        MsgBox "No file is chosen!"
        Exit Sub
    End If
        
    Call get_lib(path, SigMap_t, Prm_t, HwMap_t, Cmd_t)
    kk = Split(FileDir, ".")
    DestDir = kk(0) & "_Grt." & kk(UBound(kk))
    FileCopy FileDir, DestDir
    Dim wb As Workbook
    Dim sht As Worksheet
    Set wb = Workbooks.Open(DestDir)
    For Each sht In wb.Worksheets
        sht.Columns("G:G").Insert
        sht.Columns("E:E").Insert
        sht.Columns("B:B").Insert
        Dim Found As Boolean
        Debug.Print sht.UsedRange.Rows.Count
        For i = 2 To sht.UsedRange.Rows.Count
            For k = 1 To 10
                If k = 3 Or k = 7 Or k = 10 Then
                    'insert description
                    If sht.Cells(i, k).Value <> "" Then
                        Found = False
                        'match prm
                        If Not Found Then
                            For j = 0 To UBound(Prm_t)
                                If LCase(Replace(sht.Cells(i, k).Value, " ", "")) = LCase(Replace(Prm_t(j), " ", "")) Then
                                    sht.Cells(i, k).Value = Prm_t(j)
                                    sht.Cells(i, k - 1).Value = "DID"
                                    Found = True
                                    Exit For
                                End If
                            Next
                        End If
        
                        'end prm
                        
                        'match command
                        If Not Found Then
                            For j = 0 To UBound(Cmd_t)
                                If LCase(Replace(sht.Cells(i, k).Value, " ", "")) = LCase(Replace(Cmd_t(j), " ", "")) Then
                                    sht.Cells(i, k).Value = Cmd_t(j)
                                    sht.Cells(i, k - 1).Value = "Command"
                                    Found = True
                                    Exit For
                                End If
                            Next
                        End If
                        'end command
                        
                        'match HW
                        If Not Found Then
                            For j = 0 To UBound(HwMap_t)
                                If LCase(Replace(sht.Cells(i, k).Value, " ", "")) = LCase(Replace(HwMap_t(j).AppName, " ", "")) Then
                                    sht.Cells(i, k).Value = HwMap_t(j).AppName
                                    sht.Cells(i, k - 1).Value = "Hw::" & HwMap_t(j).HwName
                                    Found = True
                                    Exit For
                                End If
                            Next
                        End If
                        'end hw
                        
                        'match NetWork
                        If Not Found Then
                            For j = 0 To UBound(SigMap_t)
                                If LCase(Replace(sht.Cells(i, k).Value, " ", "")) = LCase(Replace(SigMap_t(j).SigName, " ", "")) Then
                                    sht.Cells(i, k).Value = SigMap_t(j).SigName
                                    If k = 3 Or k = 7 Then
                                        If SigMap_t(j).CanIn <> "" Then
                                            sht.Cells(i, k - 1).Value = "CAN::" & SigMap_t(j).CanInFrm
                                        ElseIf SigMap_t(j).LinIn <> "" Then
                                            sht.Cells(i, k - 1).Value = "LIN" & Right(SigMap_t(j).LinIn, 1) & "::" & SigMap_t(j).LinInFrm
                                        End If
                                    Else
                                        If SigMap_t(j).CanOut <> "" Then
                                            sht.Cells(i, k - 1).Value = "CAN::" & SigMap_t(j).CanOutFrm
                                        ElseIf SigMap_t(j).LinOut <> "" Then
                                            sht.Cells(i, k - 1).Value = "LIN" & Right(SigMap_t(j).LinOut, 1) & "::" & SigMap_t(j).LinOutFrm
                                        End If
                                    End If
                                    
                                    Found = True
                                    Exit For
                                End If
                            Next
                        End If
                        'end network
                    End If
                End If
            Next
        Next
    Next
    

    
    Application.DisplayAlerts = False '关闭提示
    wb.Close True
    Application.DisplayAlerts = True  '恢复打开提示
    
End Sub

Private Function Get_Path() As String
    Dim path As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Function
        path = .SelectedItems(1) & "\"
    End With
    
    Get_Path = path
End Function
Private Function SelectFile() As String
    Dim path As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear   '清除文件过滤器
        If .Show = False Then Exit Function
        path = .SelectedItems(1)
    End With
    
    SelectFile = path
End Function
