<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        ListBox1 = New ListBox()
        FolderBrowserDialog1 = New FolderBrowserDialog()
        ListBox2 = New ListBox()
        Button2 = New Button()
        Button1 = New Button()
        ListBox3 = New ListBox()
        TextBox1 = New TextBox()
        Button3 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        TextBox2 = New TextBox()
        Label3 = New Label()
        TextBox3 = New TextBox()
        TextBox4 = New TextBox()
        Label4 = New Label()
        Button4 = New Button()
        ListBox4 = New ListBox()
        Button5 = New Button()
        Button6 = New Button()
        Button7 = New Button()
        SuspendLayout()
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.HorizontalScrollbar = True
        ListBox1.ItemHeight = 17
        ListBox1.Location = New Point(12, 12)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(465, 582)
        ListBox1.TabIndex = 13
        ' 
        ' ListBox2
        ' 
        ListBox2.FormattingEnabled = True
        ListBox2.HorizontalScrollbar = True
        ListBox2.ItemHeight = 17
        ListBox2.Location = New Point(896, 12)
        ListBox2.Name = "ListBox2"
        ListBox2.Size = New Size(465, 582)
        ListBox2.TabIndex = 14
        ' 
        ' Button2
        ' 
        Button2.Enabled = False
        Button2.Location = New Point(483, 42)
        Button2.Name = "Button2"
        Button2.Size = New Size(121, 24)
        Button2.TabIndex = 12
        Button2.Text = "获取文件列表"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(483, 12)
        Button1.Name = "Button1"
        Button1.Size = New Size(121, 24)
        Button1.TabIndex = 11
        Button1.Text = "打开文件夹"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' ListBox3
        ' 
        ListBox3.FormattingEnabled = True
        ListBox3.HorizontalScrollbar = True
        ListBox3.ItemHeight = 17
        ListBox3.Location = New Point(12, 600)
        ListBox3.Name = "ListBox3"
        ListBox3.Size = New Size(1349, 106)
        ListBox3.TabIndex = 18
        ' 
        ' TextBox1
        ' 
        TextBox1.Enabled = False
        TextBox1.Location = New Point(593, 430)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(297, 23)
        TextBox1.TabIndex = 20
        ' 
        ' Button3
        ' 
        Button3.Enabled = False
        Button3.Location = New Point(483, 72)
        Button3.Name = "Button3"
        Button3.Size = New Size(121, 24)
        Button3.TabIndex = 21
        Button3.Text = "保存修改内容"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(483, 433)
        Label1.Name = "Label1"
        Label1.Size = New Size(104, 17)
        Label1.TabIndex = 22
        Label1.Text = "输入指定的字符串"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(483, 462)
        Label2.Name = "Label2"
        Label2.Size = New Size(92, 17)
        Label2.TabIndex = 23
        Label2.Text = "输入指定的长度"
        ' 
        ' TextBox2
        ' 
        TextBox2.Enabled = False
        TextBox2.Location = New Point(593, 459)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(297, 23)
        TextBox2.TabIndex = 24
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(483, 491)
        Label3.Name = "Label3"
        Label3.Size = New Size(92, 17)
        Label3.TabIndex = 25
        Label3.Text = "输入指定的位置"
        ' 
        ' TextBox3
        ' 
        TextBox3.Enabled = False
        TextBox3.Location = New Point(593, 488)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(297, 23)
        TextBox3.TabIndex = 26
        ' 
        ' TextBox4
        ' 
        TextBox4.Enabled = False
        TextBox4.Location = New Point(593, 517)
        TextBox4.Name = "TextBox4"
        TextBox4.Size = New Size(297, 23)
        TextBox4.TabIndex = 27
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(483, 520)
        Label4.Name = "Label4"
        Label4.Size = New Size(92, 17)
        Label4.TabIndex = 28
        Label4.Text = "输入指定的序次"
        ' 
        ' Button4
        ' 
        Button4.Enabled = False
        Button4.Location = New Point(483, 102)
        Button4.Name = "Button4"
        Button4.Size = New Size(121, 24)
        Button4.TabIndex = 31
        Button4.Text = "开始批量修改"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' ListBox4
        ' 
        ListBox4.Enabled = False
        ListBox4.FormattingEnabled = True
        ListBox4.ItemHeight = 17
        ListBox4.Items.AddRange(New Object() {"1-在开头增加指定字符串", "", "2-1-在中间增加指定字符串(不含后缀)", "2-2-在中间增加指定字符串(包含后缀)", "", "3-1-在末尾增加指定字符串(不含后缀)", "3-2-在末尾增加指定字符串(包含后缀)", "", "4-1-在开头删除指定长度(不含后缀)", "4-2-在开头删除指定长度(包含后缀)", "", "5-1-在中间删除指定长度(不含后缀)", "5-2-在中间删除指定长度(包含后缀)", "", "6-1-在末尾删除指定长度(不含后缀)", "6-2-在末尾删除指定长度(包含后缀)", "", "7-1-从左开始删除指定序次的指定字符串(不含后缀)", "7-2-从左开始删除指定序次的指定字符串(包含后缀)", "", "8-1-从右开始删除指定序次的指定字符串(不含后缀)", "8-2-从右开始删除指定序次的指定字符串(包含后缀)", "", "9-修改全部后缀名至同一"})
        ListBox4.Location = New Point(610, 12)
        ListBox4.Name = "ListBox4"
        ListBox4.Size = New Size(280, 412)
        ListBox4.TabIndex = 32
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(483, 340)
        Button5.Name = "Button5"
        Button5.Size = New Size(121, 24)
        Button5.TabIndex = 33
        Button5.Text = "程序使用说明"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' Button6
        ' 
        Button6.Location = New Point(483, 370)
        Button6.Name = "Button6"
        Button6.Size = New Size(121, 24)
        Button6.TabIndex = 34
        Button6.Text = "关于本程序"
        Button6.UseVisualStyleBackColor = True
        ' 
        ' Button7
        ' 
        Button7.Location = New Point(483, 400)
        Button7.Name = "Button7"
        Button7.Size = New Size(121, 24)
        Button7.TabIndex = 35
        Button7.Text = "程序更新日志"
        Button7.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1373, 717)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(Button5)
        Controls.Add(ListBox4)
        Controls.Add(Button4)
        Controls.Add(Label4)
        Controls.Add(TextBox4)
        Controls.Add(TextBox3)
        Controls.Add(Label3)
        Controls.Add(TextBox2)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(Button3)
        Controls.Add(TextBox1)
        Controls.Add(ListBox3)
        Controls.Add(ListBox1)
        Controls.Add(ListBox2)
        Controls.Add(Button2)
        Controls.Add(Button1)
        FormBorderStyle = FormBorderStyle.Fixed3D
        MaximizeBox = False
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "批量修改文件名V1.0.5"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents ListBox3 As ListBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Button4 As Button
    Friend WithEvents ListBox4 As ListBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
End Class
