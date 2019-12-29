Option Explicit

Const vbext_ct_stdmodule = 1
Const vbext_ct_classmodule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_document = 100

Private Type STRING_USED
    name As String
    used As Boolean
End Type

Sub VBA_Export()
    Dim objVBProject As Object
    Dim comp As Object
    Dim strReturn As String

    'For Each objVBProject In Application.VBE.VBProjects
        For Each comp In ThisWorkbook.VBProject.VBComponents
            If comp.Type = vbext_ct_stdmodule Or (comp.Type = vbext_ct_document And LCase(comp.name) = "thisworkbook") Then
                If comp.CodeModule.CountOfLines > 0 Then
                    Open ThisWorkbook.path + "\" + comp.name + ".vb" For Output As #1
                    Print #1, comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                    Close #1
                End If
            End If
        Next
    'Next
End Sub

Sub VBA_Module_Clear()
    Dim objVBProject As Object
    Dim comp As Object
    
    'For Each objVBProject In Application.VBE.VBProjects
        For Each comp In ThisWorkbook.VBProject.VBComponents
            If comp.Type = vbext_ct_stdmodule Then
                ThisWorkbook.VBProject.VBComponents.Remove comp
            End If
        Next
    'Next
End Sub

Sub VBA_Module_Clear_Other()
    Dim owb As Workbook
    Dim fpath As String, fname As String
    fpath = "D:\05-CodeRepository\old\CVT25_ALL_old\Engineering\Software\CBoot\Testing\Requirement\1199.9_SW_Requirements_Test_Spec of Geely_Cboot"
    fname = "1199.9_SW_Requirements_Test_Specification_BYD_CBOOT_CVT25.xlsm"
    'owb = Application.Workbooks.Open(fpath & "\" & fname)
    Set owb = Application.Workbooks(fname)
    
    Dim comp As Object
    
    'For Each objVBProject In Application.VBE.VBProjects
        For Each comp In owb.VBProject.VBComponents
            If comp.Type = vbext_ct_stdmodule Then
                owb.VBProject.VBComponents.Remove comp
            End If
        Next
    'Next
End Sub

Sub VBA_Import()
    Dim objVBProject As Object
    Dim comp As Object
    Dim strReturn As String
    Dim vba_files() As String
    Dim comp_names() As STRING_USED, vba_names() As STRING_USED
    Dim file As Variant
    Dim i As Integer, j As Integer, line_index As Integer
    Dim init_line_cnt As Integer
    

    vba_files = Get_Files(ThisWorkbook.path + "/*.vb")
    
    'delete the suffix of vba_files and update vba_names()
    ReDim vba_names(UBound(vba_files)) As STRING_USED
    For i = 0 To UBound(vba_files)
        vba_names(i).name = Replace(vba_files(i), ".vb", "")
        vba_names(i).used = False
    Next
    
    ReDim comp_names(ThisWorkbook.VBProject.VBComponents.Count - 1) As STRING_USED
    'get the name of component
    For i = 0 To ThisWorkbook.VBProject.VBComponents.Count - 1
        comp_names(i).name = ThisWorkbook.VBProject.VBComponents.Item(i + 1).name
        comp_names(i).used = False
    Next
    
    'update used state
    For i = 0 To UBound(vba_names)
        For j = 0 To UBound(comp_names)
            If vba_names(i).name = comp_names(j).name Then
                vba_names(i).used = True
                comp_names(j).used = True
            End If
        Next
    Next
    
    
    For i = 0 To UBound(vba_names)
        If LCase(vba_names(i).name) <> "thisworkbook" Then
            'add new component
            If vba_names(i).used = False Then
                ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_stdmodule).name = vba_names(i).name
            End If
            
            'update exist components' code
            init_line_cnt = ThisWorkbook.VBProject.VBComponents(vba_names(i).name).CodeModule.CountOfLines
            line_index = 0
            
            Dim temp_code As String
            Open Application.ActiveWorkbook.path + "\" + vba_files(i) For Input As #1
            While Not EOF(1)
                Line Input #1, temp_code
                line_index = line_index + 1
                ThisWorkbook.VBProject.VBComponents(vba_names(i).name).CodeModule.InsertLines init_line_cnt + line_index, temp_code
            Wend
            
            If init_line_cnt > 0 Then
                ThisWorkbook.VBProject.VBComponents(vba_names(i).name).CodeModule.DeleteLines 1, init_line_cnt
            End If
            Close #1
        End If
    Next
    
    'delete unused components
    For i = 0 To UBound(comp_names)
        If comp_names(i).used = False And ThisWorkbook.VBProject.VBComponents(comp_names(i).name).Type = vbext_ct_stdmodule Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(comp_names(i).name)
        End If
    Next
End Sub

Function Get_Files(path As String) As String()
    Dim file() As String
    Dim tempstr As String
    Dim file_num As Integer
    tempstr = Dir(path, vbNormal)
    file_num = 0
    While tempstr <> ""
        ReDim Preserve file(file_num) As String
        file(file_num) = tempstr
        file_num = file_num + 1
        tempstr = Dir()
    Wend
    
    Get_Files = file
End Function

Sub VBA_Update_VBA_M()

    Dim vba_files() As String
    Dim i As Integer, line_index As Integer
    Dim init_line_cnt As Integer

    vba_files = Get_Files(ThisWorkbook.path + "/*.vb")
    For i = 0 To UBound(vba_files)
        If LCase(vba_files(i)) = "thisworkbook.vb" Then
            line_index = 0
            Dim temp_code As String
            Dim temp_name As String
            temp_name = Replace(vba_files(i), ".vb", "")
            init_line_cnt = ThisWorkbook.VBProject.VBComponents(temp_name).CodeModule.CountOfLines
            Open Application.ActiveWorkbook.path + "\" + vba_files(i) For Input As #1
            While Not EOF(1)
                Line Input #1, temp_code
                line_index = line_index + 1
                ThisWorkbook.VBProject.VBComponents(temp_name).CodeModule.InsertLines init_line_cnt + line_index, temp_code
            Wend
            If init_line_cnt > 0 Then
                ThisWorkbook.VBProject.VBComponents(temp_name).CodeModule.DeleteLines 1, init_line_cnt
            End If
            Close #1
        End If
    Next

End Sub



Private Sub Workbook_open()
    'Call VBA_Module_Clear
    'Call VBA_Import
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Call VBA_Export
End Sub

