Imports System.IO
Imports System.Collections.ObjectModel
Public Class Form1
    Public Const UppLim = 1000000
    Public FolderPath As String
    Public RawData() As String = New String(UppLim - 1) {}
    Public FinallyData() As String = New String(UppLim - 1) {}
    Public RawCount As Long
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RawCount = 0
        ListBox3.Items.Add(Now + "请选择文件名修改文件夹路径")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            FolderPath = FolderBrowserDialog1.SelectedPath
            ListBox3.Items.Add(Now + "已选择文件名修改文件夹路径:" + FolderPath)
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            Button2.Enabled = True
            Button3.Enabled = False
            ListBox4.Enabled = False
            ListBox4.SelectedIndex = -1
        Else
            ListBox3.Items.Add(Now + "已取消文件名修改文件夹路径选择")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim FileList As ReadOnlyCollection(Of String)
        Dim AppPath As String, Comp_1 As String, Comp_2 As String
        Dim i As Long, j As Long
        Dim FolderCheck As Boolean = True
        Button3.Enabled = False
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox4.SelectedIndex = -1
        ListBox3.Items.Add(Now + "正在确认文件夹路径是否有误")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        If Boo_DirExist(FolderPath) = 0 Then
            ListBox3.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            MsgBox("错误:未找到该文件夹路径,请重新指定", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Exit Sub
        End If
        AppPath = Directory.GetCurrentDirectory()
        If Mid(AppPath, Len(AppPath), 1) <> "\" Then
            AppPath += "\"
        End If
        If Mid(FolderPath, Len(FolderPath), 1) <> "\" Then
            FolderPath += "\"
        End If
        If Len(FolderPath) = Len(AppPath) Then
            Comp_1 = ""
            Comp_2 = ""
            For i = 1 To Len(FolderPath)
                If "a" <= Mid(FolderPath, i, 1) And Mid(FolderPath, i, 1) <= "z" Then
                    Comp_1 += Chr(Asc(Mid(FolderPath, i, 1)) - 32)
                Else
                    Comp_1 += Mid(FolderPath, i, 1)
                End If
                If "a" <= Mid(AppPath, i, 1) And Mid(AppPath, i, 1) <= "z" Then
                    Comp_2 += Chr(Asc(Mid(AppPath, i, 1)) - 32)
                Else
                    Comp_2 += Mid(AppPath, i, 1)
                End If
            Next i
            If StrComp(Comp_1, Comp_2) = 0 Then FolderCheck = False
        End If
        If FolderCheck = False Then
            ListBox3.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            MsgBox("错误:文件夹路径不能为程序所在文件夹", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Exit Sub
        End If
        ListBox3.Items.Add(Now + "文件夹路径无误,开始枚举本层文件夹下的文件")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        FileList = My.Computer.FileSystem.GetFiles(FolderPath)
        If FileList.Count = 0 Then
            RawCount = 0
            ListBox3.Items.Add(Now + "文件列表已输出完成,文件夹总文件个数:" + CStr(FileList.Count) + "|实际导入个数:" + CStr(RawCount))
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            MsgBox("本层文件夹下没有文件", 0 + vbInformation + vbSystemModal, "批量修改文件名")
        Else
            RawCount = 0
            For i = 0 To FileList.Count - 1
                For j = Len(My.Computer.FileSystem.GetName(FileList(i))) To 1 Step -1
                    If Mid(My.Computer.FileSystem.GetName(FileList(i)), j, 1) = "." Then
                        RawData(RawCount) = My.Computer.FileSystem.GetName(FileList(i))
                        RawCount += 1
                        Exit For
                    End If
                Next j
                If RawCount = UppLim Then
                    ListBox3.Items.Add(Now + "文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件")
                    ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                    MsgBox("文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件", 0 + vbExclamation + vbYesNo, "文件导入个数过多")
                    Exit For
                End If
            Next i
            ListBox3.Items.Add(Now + "文件夹枚举完成,开始输出文件列表")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            ListBox1.Items.Add(Now + "|修改前文件名如下:")
            For i = 0 To RawCount - 1
                ListBox1.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + RawData(i))
            Next i
            If RawCount > UppLim Then RawCount = UppLim
            ListBox3.Items.Add(Now + "文件列表已输出完成,文件夹总文件个数:" + CStr(FileList.Count) + "|实际导入个数:" + CStr(RawCount))
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            ListBox4.Enabled = True
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Long, j As Long, tmp As Long, num As Long
        Dim RepCheck As Boolean = False
        Dim flag_arr() As Long = New Long(UppLim - 1) {}
        Dim temp As String
        For i = 0 To UppLim - 1
            FinallyData(i) = ""
        Next i
        ListBox3.Items.Add(Now + "开始按照规则进行文件名修改编排")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        If ListBox4.SelectedIndex = 0 And TextBox1.Text <> "" Then '0 1-在开头增加指定字符串
            For i = 0 To RawCount - 1
                FinallyData(i) = TextBox1.Text + RawData(i)
            Next i
        ElseIf ListBox4.SelectedIndex = 2 And TextBox1.Text <> "" And TextBox3.Text <> "" Then '2 2-1-在中间增加指定字符串(不含后缀)
            If Val(TextBox3.Text) = 0 Then
                For i = 0 To RawCount - 1
                    FinallyData(i) = TextBox1.Text + RawData(i)
                Next i
            Else
                For i = 0 To RawCount - 1
                    For j = Len(RawData(i)) To 1 Step -1
                        If Mid(RawData(i), j, 1) = "." Then Exit For
                    Next j
                    temp = Mid(RawData(i), 1, j - 1)
                    FinallyData(i) = Mid(temp, 1, Val(TextBox3.Text)) + TextBox1.Text + Mid(temp, Val(TextBox3.Text) + 1) + Mid(RawData(i), j)
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 3 And TextBox1.Text <> "" And TextBox3.Text <> "" Then '3 2-2-在中间增加指定字符串(包含后缀)
            If Val(TextBox3.Text) = 0 Then
                For i = 0 To RawCount - 1
                    FinallyData(i) = TextBox1.Text + RawData(i)
                Next i
            Else
                For i = 0 To RawCount - 1
                    FinallyData(i) = Mid(RawData(i), 1, Val(TextBox3.Text)) + TextBox1.Text + Mid(RawData(i), Val(TextBox3.Text) + 1)
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 5 And TextBox1.Text <> "" Then '5 3-1-在末尾增加指定字符串(不含后缀)
            For i = 0 To RawCount - 1
                For j = Len(RawData(i)) To 1 Step -1
                    If Mid(RawData(i), j, 1) = "." Then Exit For
                Next j
                FinallyData(i) = Mid(RawData(i), 1, j - 1) + TextBox1.Text + Mid(RawData(i), j)
            Next i
        ElseIf ListBox4.SelectedIndex = 6 And TextBox1.Text <> "" Then '6 3-2-在末尾增加指定字符串(包含后缀)
            For i = 0 To RawCount - 1
                FinallyData(i) = RawData(i) + TextBox1.Text
            Next i
        ElseIf ListBox4.SelectedIndex = 8 And TextBox2.Text <> "" Then '8 4-1-在开头删除指定长度(不含后缀)
            For i = 0 To RawCount - 1
                For j = Len(RawData(i)) To 1 Step -1
                    If Mid(RawData(i), j, 1) = "." Then Exit For
                Next j
                If Val(TextBox2.Text) >= j Then
                    FinallyData(i) = Mid(RawData(i), j)
                Else
                    FinallyData(i) = Mid(RawData(i), Val(TextBox2.Text) + 1)
                End If
            Next i
        ElseIf ListBox4.SelectedIndex = 9 And TextBox2.Text <> "" Then '9 4-2-在开头删除指定长度(包含后缀)
            For i = 0 To RawCount - 1
                j = Len(RawData(i))
                If Val(TextBox2.Text) >= j Then
                    ListBox3.Items.Add(Now + "指定的长度大于文件名的长度,文件名将保留最后一个字符")
                    ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                    FinallyData(i) = Mid(RawData(i), j)
                Else
                    FinallyData(i) = Mid(RawData(i), Val(TextBox2.Text) + 1)
                End If
            Next i
        ElseIf ListBox4.SelectedIndex = 11 And TextBox2.Text <> "" And TextBox3.Text <> "" Then '11 5-1-在中间删除指定长度(不含后缀)
            If Val(TextBox3.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的位置数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    For j = Len(RawData(i)) To 1 Step -1
                        If Mid(RawData(i), j, 1) = "." Then Exit For
                    Next j
                    temp = Mid(RawData(i), 1, j - 1)
                    FinallyData(i) = Mid(temp, 1, Val(TextBox3.Text) - 1) + Mid(temp, Val(TextBox2.Text) + Val(TextBox3.Text)) + Mid(RawData(i), j)
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 12 And TextBox2.Text <> "" And TextBox3.Text <> "" Then '12 5-2-在中间删除指定长度(包含后缀)
            If Val(TextBox3.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的位置数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    j = Len(RawData(i))
                    If Val(TextBox2.Text) + Val(TextBox3.Text) - 1 >= j And Val(TextBox3.Text) = 1 Then
                        ListBox3.Items.Add(Now + "指定的长度大于文件名的长度,文件名将保留最后一个字符")
                        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                        FinallyData(i) = Mid(RawData(i), j)
                    Else
                        FinallyData(i) = Mid(RawData(i), 1, Val(TextBox3.Text) - 1) + Mid(RawData(i), Val(TextBox2.Text) + Val(TextBox3.Text))
                    End If
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 14 And TextBox2.Text <> "" Then '14 6-1-在末尾删除指定长度(不含后缀)
            For i = 0 To RawCount - 1
                For j = Len(RawData(i)) To 1 Step -1
                    If Mid(RawData(i), j, 1) = "." Then Exit For
                Next j
                temp = Mid(RawData(i), 1, j - 1)
                If Len(temp) >= Val(TextBox2.Text) Then
                    FinallyData(i) = Mid(temp, 1, Len(temp) - Val(TextBox2.Text)) + Mid(RawData(i), j)
                Else
                    FinallyData(i) = Mid(RawData(i), j)
                End If
            Next i
        ElseIf ListBox4.SelectedIndex = 15 And TextBox2.Text <> "" Then '15 6-2-在末尾删除指定长度(包含后缀)
            For i = 0 To RawCount - 1
                If Len(RawData(i)) >= Val(TextBox2.Text) Then
                    FinallyData(i) = Mid(RawData(i), 1, Len(RawData(i)) - Val(TextBox2.Text))
                Else
                    ListBox3.Items.Add(Now + "指定的长度大于文件名的长度,文件名将保留第一个字符")
                    ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                    FinallyData(i) = Mid(RawData(i), 1, 1)
                End If
            Next i
        ElseIf ListBox4.SelectedIndex = 17 And TextBox1.Text <> "" And TextBox4.Text <> "" Then '17 7-1-从左开始删除指定序次的指定字符串(不含后缀)
            If Val(TextBox4.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的序次数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    For j = Len(RawData(i)) To 1 Step -1
                        If Mid(RawData(i), j, 1) = "." Then Exit For
                    Next j
                    temp = Mid(RawData(i), 1, j - 1)
                    tmp = 0
                    For j = 1 To Len(temp) - Len(TextBox1.Text) + 1
                        If StrComp(Mid(temp, j, Len(TextBox1.Text)), TextBox1.Text) = 0 Then
                            tmp += 1
                            If tmp = Val(TextBox4.Text) Then
                                FinallyData(i) = Mid(temp, 1, j - 1) + Mid(temp, j + Len(TextBox1.Text)) + Mid(RawData(i), Len(temp) + 1)
                                Exit For
                            End If
                        End If
                    Next j
                    If tmp < Val(TextBox4.Text) Then
                        FinallyData(i) = RawData(i)
                    End If
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 18 And TextBox1.Text <> "" And TextBox4.Text <> "" Then '18 7-2-从左开始删除指定序次的指定字符串(包含后缀)
            If Val(TextBox4.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的序次数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    tmp = 0
                    For j = 1 To Len(RawData(i)) - Len(TextBox1.Text) + 1
                        If StrComp(Mid(RawData(i), j, Len(TextBox1.Text)), TextBox1.Text) = 0 Then
                            tmp += 1
                            If tmp = Val(TextBox4.Text) Then
                                FinallyData(i) = Mid(RawData(i), 1, j - 1) + Mid(RawData(i), j + Len(TextBox1.Text))
                                Exit For
                            End If
                        End If
                    Next j
                    If tmp < Val(TextBox4.Text) Or StrComp(RawData(i), TextBox1.Text) = 0 Then
                        FinallyData(i) = RawData(i)
                    End If
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 20 And TextBox1.Text <> "" And TextBox4.Text <> "" Then '20 8-1-从右开始删除指定序次的指定字符串(不含后缀)
            If Val(TextBox4.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的序次数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    For j = Len(RawData(i)) To 1 Step -1
                        If Mid(RawData(i), j, 1) = "." Then Exit For
                    Next j
                    temp = Mid(RawData(i), 1, j - 1)
                    tmp = 0
                    For j = Len(temp) - Len(TextBox1.Text) + 1 To 1 Step -1
                        If StrComp(Mid(temp, j, Len(TextBox1.Text)), TextBox1.Text) = 0 Then
                            tmp += 1
                            If tmp = Val(TextBox4.Text) Then
                                FinallyData(i) = Mid(temp, 1, j - 1) + Mid(temp, j + Len(TextBox1.Text)) + Mid(RawData(i), Len(temp) + 1)
                                Exit For
                            End If
                        End If
                    Next j
                    If tmp < Val(TextBox4.Text) Then
                        FinallyData(i) = RawData(i)
                    End If
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 21 And TextBox1.Text <> "" And TextBox4.Text <> "" Then '21 8-2-从右开始删除指定序次的指定字符串(包含后缀)
            If Val(TextBox4.Text) = 0 Then
                ListBox3.Items.Add(Now + "指定的序次数值输入错误(要求大于0),程序已中止,请在更正输入后再次点击执行")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                Exit Sub
            Else
                For i = 0 To RawCount - 1
                    tmp = 0
                    For j = Len(RawData(i)) - Len(TextBox1.Text) + 1 To 1 Step -1
                        If StrComp(Mid(RawData(i), j, Len(TextBox1.Text)), TextBox1.Text) = 0 Then
                            tmp += 1
                            If tmp = Val(TextBox4.Text) Then
                                FinallyData(i) = Mid(RawData(i), 1, j - 1) + Mid(RawData(i), j + Len(TextBox1.Text))
                                Exit For
                            End If
                        End If
                    Next j
                    If tmp < Val(TextBox4.Text) Or StrComp(RawData(i), TextBox1.Text) = 0 Then
                        FinallyData(i) = RawData(i)
                    End If
                Next i
            End If
        ElseIf ListBox4.SelectedIndex = 23 And TextBox1.Text <> "" Then  '23 9-修改全部后缀名至同一
            For i = 0 To RawCount - 1
                For j = Len(RawData(i)) To 1 Step -1
                    If Mid(RawData(i), j, 1) = "." Then Exit For
                Next j
                FinallyData(i) = Mid(RawData(i), 1, j) + TextBox1.Text
            Next i
        Else
            ListBox3.Items.Add(Now + "选择的功能有参数未指定,程序已中止,请在全部指定后再次点击执行")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            Exit Sub
        End If
        ListBox3.Items.Add(Now + "文件名修改编排已完成,开始检查文件名合法性")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        For i = 0 To RawCount - 1
            If Len(FolderPath + FinallyData(i)) > 252 Then
                FinallyData(i) = Mid(FinallyData(i), Len(FinallyData(i)) + Len(FolderPath) - 251)
                ListBox3.Items.Add(Now + "出现文件名长度超出Windows系统限制,已进行处理")
                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            End If
        Next i
        For i = 0 To UppLim - 1
            flag_arr(i) = 0
        Next i
        num = 1
        For i = 0 To RawCount - 2
            If flag_arr(i) = 0 Then
                For j = i + 1 To RawCount - 1
                    If StrComp(FinallyData(i), FinallyData(j)) = 0 Then
                        If RepCheck = False Then
                            flag_arr(i) = num
                            ListBox3.Items.Add(Now + "出现文件名重复,已进行处理")
                            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                            RepCheck = True
                        End If
                        num += 1
                        flag_arr(j) = num
                        ListBox3.Items.Add(Now + "出现文件名重复,已进行处理")
                        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                    End If
                Next j
                If RepCheck = True Then
                    num = 1
                    RepCheck = False
                End If
            End If
        Next i
        For i = 0 To RawCount - 1
            If flag_arr(i) <> 0 Then FinallyData(i) = CStr(flag_arr(i)) + "_" + FinallyData(i)
        Next i
        ListBox3.Items.Add(Now + "重复文件名检查已完成,开始输出修改后文件名列表")
        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        ListBox2.Items.Clear()
        ListBox2.Items.Add(Now + "|修改后文件名如下:")
        For i = 0 To RawCount - 1
            ListBox2.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + FinallyData(i))
        Next i
        Button4.Enabled = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim i As Long, ChgFail As Long = 0, Unmodified As Long = 0
        Dim IgnFlag As Boolean = False
        Dim Result As MsgBoxResult
        Result = MsgBox("开始批量修改文件名,请确认所有改动操作准确无误", 0 + vbExclamation + vbYesNo, "警告")
        If Result = vbYes Then
            Button3.Enabled = False
            Button4.Enabled = False
            ListBox4.Enabled = False
            ListBox4.SelectedIndex = -1
            ListBox3.Items.Add(Now + "已选择是,开始批量修改文件名")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            For i = 0 To RawCount - 1
                If File.Exists(FolderPath + RawData(i)) = False Then
                    ChgFail += 1
                    ListBox3.Items.Add(Now + "文件:" + RawData(i) + "修改失败")
                    ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                    If IgnFlag = False Then
                        MsgBox("错误:" + RawData(i) + vbCrLf + "文件修改失败,文件不存在", 0 + vbCritical + vbSystemModal, "修改文件名失败")
                        ListBox3.Items.Add(Now + "已发生修改失败,开始询问是否继续执行修改操作")
                        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                        Result = MsgBox("已发生修改失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来的所有修改失败警告", 0 + vbExclamation + vbYesNo, "警告")
                        If Result = vbYes Then
                            IgnFlag = True
                            ListBox3.Items.Add(Now + "已选择是,将继续执行并忽视修改失败警告")
                            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                        Else
                            ListBox3.Items.Add(Now + "已选择否,批量修改文件名操作已终止")
                            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                            Exit Sub
                        End If
                    End If
                Else
                    Try
                        If StrComp(FolderPath + RawData(i), FolderPath + FinallyData(i)) <> 0 Then
                            Rename(FolderPath + RawData(i), FolderPath + FinallyData(i))
                        Else
                            Unmodified += 1
                        End If
                    Catch ex As Exception
                        ChgFail += 1
                        ListBox3.Items.Add(Now + "文件:" + RawData(i) + "修改失败")
                        ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                        If IgnFlag = False Then
                            MsgBox("错误:" + RawData(i) + vbCrLf + "文件修改失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "修改文件名失败")
                            ListBox3.Items.Add(Now + "已发生修改失败,开始询问是否继续执行修改操作")
                            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                            Result = MsgBox("已发生修改失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来的所有修改失败警告", 0 + vbExclamation + vbYesNo, "警告")
                            If Result = vbYes Then
                                IgnFlag = True
                                ListBox3.Items.Add(Now + "已选择是,将继续执行并忽视修改失败警告")
                                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                            Else
                                ListBox3.Items.Add(Now + "已选择否,批量修改文件名操作已终止")
                                ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                                Exit Sub
                            End If
                        End If
                    End Try
                    ListBox3.Items.Add(Now + "文件:" + RawData(i) + "修改成功")
                    ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
                End If
            Next i
            ListBox3.Items.Add(Now + "批量修改文件名完成,计划修改:" + CStr(RawCount) + ",成功:" + CStr(RawCount - Unmodified - ChgFail) + ",失败:" + CStr(ChgFail) + ",未修改:" + CStr(Unmodified))
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            MsgBox("批量修改文件名完成,计划修改:" + CStr(RawCount) + ",成功:" + CStr(RawCount - Unmodified - ChgFail) + ",失败:" + CStr(ChgFail) + ",未修改:" + CStr(Unmodified), 0 + vbInformation + vbSystemModal, "批量修改文件名完成")
        Else
            ListBox3.Items.Add(Now + "已选择否,批量修改文件名操作已终止")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
        End If
    End Sub
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        ListBox2.Items.Clear()
        Button3.Enabled = True
        Button4.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        If ListBox4.SelectedIndex = 0 Then '0 1-在开头增加指定字符串
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 2 Then '2 2-1-在中间增加指定字符串(不含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 3 Then '3 2-2-在中间增加指定字符串(包含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 5 Then '5 3-1-在末尾增加指定字符串(不含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 6 Then '6 3-2-在末尾增加指定字符串(包含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 8 Then '8 4-1-在开头删除指定长度(不含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 9 Then '9 4-2-在开头删除指定长度(包含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 11 Then '11 5-1-在中间删除指定长度(不含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 12 Then '12 5-2-在中间删除指定长度(包含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 14 Then  '14 6-1-在末尾删除指定长度(不含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 15 Then '15 -2-在末尾删除指定长度(包含后缀)
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        ElseIf ListBox4.SelectedIndex = 17 Then  '17 7-1-从左开始删除指定序次的指定字符串(不含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = True
        ElseIf ListBox4.SelectedIndex = 18 Then '18 7-2-从左开始删除指定序次的指定字符串(包含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = True
        ElseIf ListBox4.SelectedIndex = 20 Then  '20 8-1-从右开始删除指定序次的指定字符串(不含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = True
        ElseIf ListBox4.SelectedIndex = 21 Then '21 8-2-从右开始删除指定序次的指定字符串(包含后缀)
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = True
        ElseIf ListBox4.SelectedIndex = 23 Then  '23 9-修改全部后缀名至同一
            TextBox1.Enabled = True
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        Else '非有效操作全部禁用
            Button3.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim i As Integer, count As Integer
        Dim numindex(260) As Integer
        Dim temp As String
        Button4.Enabled = False
        If TextBox1.Text.Contains("\"c) Or TextBox1.Text.Contains("/"c) Or TextBox1.Text.Contains(":"c) Or TextBox1.Text.Contains("*"c) Or
            TextBox1.Text.Contains("?"c) Or TextBox1.Text.Contains(""""c) Or TextBox1.Text.Contains("<"c) Or TextBox1.Text.Contains(">"c) Or TextBox1.Text.Contains("|"c) Then
            ListBox3.Items.Add(Now + "请注意,键入了非法字符,以下非法字符不被接受\/:*?""<>|")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            count = 0
            For i = 1 To Len(TextBox1.Text)
                If Mid(TextBox1.Text, i, 1) = "\" Or Mid(TextBox1.Text, i, 1) = "/" Or Mid(TextBox1.Text, i, 1) = ":" Or Mid(TextBox1.Text, i, 1) = "*" Or
            Mid(TextBox1.Text, i, 1) = "?" Or Mid(TextBox1.Text, i, 1) = "<" Or Mid(TextBox1.Text, i, 1) = ">" Or Mid(TextBox1.Text, i, 1) = """" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = TextBox1.Text
            For i = 0 To count - 1
                If numindex(i) <> 1 And numindex(i) <> Len(TextBox1.Text) Then
                    temp = Mid(TextBox1.Text, 1, numindex(i) - 1) + Mid(TextBox1.Text, numindex(i) + 1)
                ElseIf numindex(i) = 1 Then
                    temp = Mid(TextBox1.Text, numindex(i) + 1)
                ElseIf numindex(i) = Len(TextBox1.Text) Then
                    temp = Mid(TextBox1.Text, 1, numindex(i) - 1)
                End If
            Next i
            TextBox1.Text = temp
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim i As Integer, count As Integer
        Dim flag As Boolean = False
        Dim numindex(260) As Integer
        Dim temp As String
        Button4.Enabled = False
        For i = 1 To Len(TextBox2.Text)
            If Mid(TextBox2.Text, i, 1) < "0" Or Mid(TextBox2.Text, i, 1) > "9" Then
                flag = True
                Exit For
            End If
        Next i
        If flag = True Then
            ListBox3.Items.Add(Now + "请注意,键入了非法字符,数字以外的字符不被接受")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            count = 0
            For i = 1 To Len(TextBox2.Text)
                If "0" <= Mid(TextBox2.Text, i, 1) And Mid(TextBox2.Text, i, 1) <= "9" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = ""
            For i = 0 To count - 1
                temp += Mid(TextBox2.Text, numindex(i), 1)
            Next i
            TextBox2.Text = temp
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim i As Integer, count As Integer
        Dim flag As Boolean = False
        Dim numindex(260) As Integer
        Dim temp As String
        Button4.Enabled = False
        For i = 1 To Len(TextBox3.Text)
            If Mid(TextBox3.Text, i, 1) < "0" Or Mid(TextBox3.Text, i, 1) > "9" Then
                flag = True
                Exit For
            End If
        Next i
        If flag = True Then
            ListBox3.Items.Add(Now + "请注意,键入了非法字符,数字以外的字符不被接受")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            count = 0
            For i = 1 To Len(TextBox3.Text)
                If "0" <= Mid(TextBox3.Text, i, 1) And Mid(TextBox3.Text, i, 1) <= "9" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = ""
            For i = 0 To count - 1
                temp += Mid(TextBox3.Text, numindex(i), 1)
            Next i
            TextBox3.Text = temp
        End If
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Dim i As Integer, count As Integer
        Dim flag As Boolean = False
        Dim numindex(260) As Integer
        Dim temp As String
        Button4.Enabled = False
        For i = 1 To Len(TextBox4.Text)
            If Mid(TextBox4.Text, i, 1) < "0" Or Mid(TextBox4.Text, i, 1) > "9" Then
                flag = True
                Exit For
            End If
        Next i
        If flag = True Then
            ListBox3.Items.Add(Now + "请注意,键入了非法字符,数字以外的字符不被接受")
            ListBox3.SelectedItem = ListBox3.Items(ListBox3.Items.Count - 1)
            count = 0
            For i = 1 To Len(TextBox4.Text)
                If "0" <= Mid(TextBox4.Text, i, 1) And Mid(TextBox4.Text, i, 1) <= "9" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = ""
            For i = 0 To count - 1
                temp += Mid(TextBox4.Text, numindex(i), 1)
            Next i
            TextBox4.Text = temp
        End If
    End Sub
    Private Shared Function Boo_DirExist(StrPath As String) As Boolean
        Boo_DirExist = Directory.Exists(StrPath)
    End Function
    Private Shared Function NumLength(Num As String) As String
        If Len(Num) = 1 Then
            NumLength = "     " + Num
        ElseIf Len(Num) = 2 Then
            NumLength = "    " + Num
        ElseIf Len(Num) = 3 Then
            NumLength = "   " + Num
        ElseIf Len(Num) = 4 Then
            NumLength = "  " + Num
        ElseIf Len(Num) = 5 Then
            NumLength = " " + Num
        Else
            NumLength = Num
        End If
    End Function
End Class