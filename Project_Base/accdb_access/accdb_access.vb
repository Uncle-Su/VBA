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
         Cells(i, 1) = rs.Fields(i - 1).name
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
    Dim fpath As String, CnStr As String, SqlStr As String
    Dim sname() As String
    
    rs.CursorLocation = adUseClient  ' if this item doesn't set, recordcount will be -1
    Dim cmd As New ADODB.Command
    Dim i As Long
    
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"        'access 2007/2010/2013 64bit version use this ace engine
                                                    ' maybe 32bit version sw can use "Microsoft.Jet.Oledb.4.0"
    'cn.Properties("Extended  Properties").Value = "Excel 12.0"
    fpath = ThisWorkbook.path & "\数据库\Excel_1.xlsx"
    CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & fpath & ";" & _
            "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"";"     'IMEX indicate the mode of using Excel file   0-output  1- input 2-mix
    cn.ConnectionString = CnStr

    'cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Extended Properties=""Excel 12.0 Macro;HDR=NO"";" & "Data Source=" & ThisWorkbook.Path & "\数据库\Excel_1.xlsx;"
    
    cn.Open
    
    p = 0
    If True Then
        Set rs = cn.OpenSchema(adSchemaTables, TABLE_NAME)
            
            ActiveSheet.Cells.Clear
            For k = 1 To rs.Fields.Count
                Cells(p + 1, k) = rs.Fields(k - 1).name
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
         Cells(1, i) = rs.Fields(i - 1).name
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
SheetNum(i) = Left(tbSheet.name, InStr(1, tbSheet.name, "$") - 1)
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
