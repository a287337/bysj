Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmipt_testwell
    '*********************************************************************************************************************************************
    '                                                   关于井斜数据输入、修改、管理窗体的说明
    '                                                                                           秦彦斌 2022年1月19日最后整理
    ' 说明：
    '     （1）在软件运行时，需要计算任意点的井斜角和方位角，系统提供了两种算法，一种是线性插值，另一种是三次样条函数插值。在用三次样条函数插值计算时，
    ' 需要计算一些插值运算系数，这些系数与测斜数据有关，存放在“三次井眼样条函数参数表”中，具体计算见子程序cal_jx_fw_cs2及注释。
    '     （2）当测井数据变化时，需要重新计算生成“三次井眼样条函数参数表”，故设模块级开关变量update_Mm，用以描述测斜数据是否发生了变化，若变化，在窗
    ' 体退出前，重新计算生成“三次井眼样条函数参数表”。重新生成“三次井眼样条函数参数表”开关，True--需要重新计算，False--不需要重新计算，
    '     （3）设计数器，记连续未读着数据的Excel行数，超过10跳出循环，避免长时间读空数。
    '
    '
    ' 程序升级记事：
    ' 20200712,完成升级到VS2008。
    '     （1）模块级链接cn_userdb、OleDbDataAdapter-ad、DataSet-dst、表dst.Tables("Testwell_Table")，模块中大家都能用。
    '     （2）油气井表用Reader，一次性，省内存及资源
    '     （3）在函数 Crt_testwell_Ad中设置好InsertCommand、UpdateCommand 和 DeleteCommand 属性，利用此OleDbDataAdapter可解决井斜数据的查、增、改、删功能。
    '     （4）“添加/更改[&A]”按钮改动内存表dst.Tables("Testwell_Table")，保存按钮利用OleDbDataAdapter写入数据库。
    '     （5）两种方法实现Excel数据导入。
    '     （6）用于 Extended Properties 值的有效 Excel 版本见注释及附近的源程序。
    '     （7）Me.FormClosed设置、写入相应的信息，记录井斜表的变动情况，以为其它程序所用。
    '     （8）顺便升级了cal_jx_fw_cs.vb中cal_xs_caljxfwcs()函数，data_str.vb中write_chanshu()函数。
    '     （9）摸索掌握了VS中操作ACCESS数据库方法、操作Excel方法、BindingSource、dataview等控件用法。
    '     （10）有用的网络资料：
    '        VB.NET学习笔记：ADO.NET操作ACCESS数据库——OleDbDataAdapter的Update方法更新数据库的秘密（行状态RowState和行版本 DataRowVersion）
    '          https://blog.csdn.net/zyjq52uys/article/details/88948594
    '        VB.NET学习笔记：ADO.NET操作ACCESS数据库——数据集DataSet数据管理（DataAdapter查询更新数据库）
    '          https://blog.csdn.net/zyjq52uys/article/details/88665663
    '        VB.NET学习笔记：ADO.NET操作ACCESS数据库——ADO.NET数据访问接口
    '          https://blog.csdn.net/zyjq52uys/article/details/88546802
    '        VB.NET学习笔记：ADO.NET操作ACCESS数据库——使用OleDbDataReader对象
    '          https://blog.csdn.net/zyjq52uys/article/details/88941132
    '        学习记录：VB.NET.操作ACCESS数据库
    '          https://blog.csdn.net/iamtsfw/article/details/91819136
    '          vb.net操作数据库之ACCESS(一)
    '          https://blog.csdn.net/lengyff/article/details/44886515
    '        OleDbDataAdapter
    '          https://wenku.baidu.com/view/49969a3e580216fc700afd19.html
    '        OleDbDataAdapter 具体使用案例
    '         https://blog.csdn.net/longtenggenssupreme/article/details/72357350
    '        BindingSource控件介绍
    '         https://blog.csdn.net/byondocean/article/details/6867214
    ' 20220119，重新布置界面控件排列，实现可最大化（缩放）。

    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private update_Mm As Boolean
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmipt_testwell_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '模块级开关变量update_Mm赋初值
        update_Mm = False
        Me.Text = well_name & "井井眼轨迹数据输入与管理"

        '用ADO.NET给井基本数据赋值
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位 from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            Text1.Text = RECreader.Item("地理位置").ToString
            Text2.Text = RECreader.Item("构造位置").ToString
            Text3.Text = RECreader.Item("井别").ToString
            Text4.Text = RECreader.Item("设计井深m").ToString
            Text5.Text = RECreader.Item("完钻井深m").ToString
            Text6.Text = RECreader.Item("完钻层位").ToString
        End If
        RECreader.Close()
        cn_userdb.Close()

        '测井数据文本框赋初值
        Text7.Text = CStr(0)
        Text8.Text = CStr(0.0#)
        Text9.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Call my_gridview1_flash()
    End Sub
    '*********************************************************************************************************************************************
    '设置并填充DataGridView1
    '*********************************************************************************************************************************************
    Private Sub my_gridview1_flash()
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        ad = Crt_testwell_Ad(cn_userdb)
        dst.Tables("Testwell_Table").Clear()
        ad.Fill(dst, "Testwell_Table")
        BindingSource1.DataSource = dst.Tables("Testwell_Table")
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        '设置并填充DataGridView1
        DataGridView1.MultiSelect = False
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '点击“退出”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        If dst.HasChanges Then
            msg_prompt = "数据有改动，是否保存，然后退出？"
            msg_buttons = 3 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Select Case msg_return
                Case 6
                    ad.Update(dst.Tables("Testwell_Table"))
                    '重新生成“三次井眼样条函数参数表”开关，True--需要重新计算，False--不需要重新计算。
                    update_Mm = True
                    Me.Close()
                Case 7
                    Me.Close()
                Case 2
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '“退出”时，判断重新生成“三次井眼样条函数参数表”开关，True--需要重新计算，False--不需要重新计算。
    '*********************************************************************************************************************************************
    Private Sub frmipt_testwell_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If update_Mm = True Then
            Call cal_xs_caljxfwcs()
            Call write_chanshu("套管磨损计算-是否重算磨损量", "是")
            '调试语句
            'msg_prompt = "数据有改动，已写参数。"
            'msg_buttons = 0 + 64
            'msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '点击“保存”按钮事件
    '*********************************************************************************************************************************************
    Private Sub CMDsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMDsave.Click
        If dst.HasChanges Then
            ad.Update(dst.Tables("Testwell_Table"))
            dst.AcceptChanges()
            '重新生成“三次井眼样条函数参数表”开关，True--需要重新计算，False--不需要重新计算。
            update_Mm = True
            msg_prompt = "数据改动已保存。"
            msg_buttons = 0 + 64
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            msg_prompt = "数据没有改动，不用保存。"
            msg_buttons = 0 + 64
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '*********************************************************************************************************************************************
    '                                                 创建一个 井斜数据OleDbDataAdapter
    '                                                                                              秦彦斌 2019年9月13日中秋节完成
    '
    '   在此函数中设置好：SelectCommand 属性返回数据源中的数据；InsertCommand、UpdateCommand 和 DeleteCommand 属性用于管理数据源中的更改
    '   如此，利用此OleDbDataAdapter可解决井斜数据的查、增、改、删功能
    '*********************************************************************************************************************************************
    Private Function Crt_testwell_Ad(ByVal connection As OleDbConnection) As OleDbDataAdapter
        Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter()
        Dim command As OleDbCommand
        Dim SQL_command As String

        'Create the SelectCommand.
        SQL_command = "select [序   号],[井  深(m)],[井斜角(°)],[方位角(°)] from 测井数据表 where [井号] = ? order by [井  深(m)]"
        command = New OleDbCommand(SQL_command, connection)
        command.Parameters.Add("wellname", OleDbType.VarChar, 50).Value = well_name
        dataAdapter.SelectCommand = command

        'Create the  DeleteCommand.
        command = New OleDbCommand("DELETE * from 测井数据表 where [井  深(m)] = ? and 井号 = ? ", connection)
        command.Parameters.Add("jingshen", OleDbType.Double, 15, "井  深(m)").SourceVersion = DataRowVersion.Original
        command.Parameters.Add("wellname", OleDbType.VarChar, 50).Value = well_name
        dataAdapter.DeleteCommand = command

        ' Create the UpdateCommand.
        command = New OleDbCommand("UPDATE 测井数据表 set [序   号] = ? ,[井斜角(°)] = ? ,[方位角(°)] = ? where [井  深(m)]= ? and 井号 = ? ", connection)
        command.Parameters.Add("xuhao", OleDbType.Integer, 15, "序   号")
        command.Parameters.Add("jxiejiao", OleDbType.Double, 15, "井斜角(°)")
        command.Parameters.Add("fangweijiao", OleDbType.Double, 15, "方位角(°)")
        command.Parameters.Add("jingshen", OleDbType.Double, 15, "井  深(m)")
        command.Parameters.Add("wellname", OleDbType.VarChar, 50).Value = well_name
        dataAdapter.UpdateCommand = command

        ' Create the InsertCommand.
        command = New OleDbCommand("insert into 测井数据表 ([序   号],[井  深(m)],[井斜角(°)],[方位角(°)],[井号]) values (?,?,?,?,?)", connection)
        command.Parameters.Add("xuhao", OleDbType.Integer, 15, "序   号")
        command.Parameters.Add("jingshen", OleDbType.Double, 15, "井  深(m)")
        command.Parameters.Add("jxiejiao", OleDbType.Double, 15, "井斜角(°)")
        command.Parameters.Add("fangweijiao", OleDbType.Double, 15, "方位角(°)")
        command.Parameters.Add("wellname", OleDbType.VarChar, 50).Value = well_name
        dataAdapter.InsertCommand = command
        Return dataAdapter
    End Function
    '*********************************************************************************************************************************************
    '点击“帮助”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '点击“添加/更改[&A]”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Dim foundRows() As DataRow
        Dim tw_newRow As DataRow
        Dim rowfilter As String
        If Val(Text7.Text) = 0 Or (Not IsNumeric(Text7.Text)) Or (Not IsNumeric(Text8.Text)) Or (Not IsNumeric(Text9.Text)) Or (Not IsNumeric(Text10.Text)) Then
            msg_prompt = "序号不应为零，井深、井斜角、方位角应是数字，请输入合理数据！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        rowfilter = "[井  深(m)] = " & CStr(Text8.Text)
        foundRows = dst.Tables("Testwell_Table").Select(rowfilter)
        If foundRows.Length = 0 Then
            tw_newRow = dst.Tables("Testwell_Table").NewRow()
            tw_newRow.Item("序   号") = CType(Text7.Text, Integer)
            tw_newRow.Item("井  深(m)") = CType(Text8.Text, Double)
            tw_newRow.Item("井斜角(°)") = CType(Text9.Text, Double)
            tw_newRow.Item("方位角(°)") = CType(Text10.Text, Double)
            dst.Tables("Testwell_Table").Rows.Add(tw_newRow)
        Else
            foundRows(0).Item("序   号") = CType(Text7.Text, Integer)
            foundRows(0).Item("井  深(m)") = CType(Text8.Text, Double)
            foundRows(0).Item("井斜角(°)") = CType(Text9.Text, Double)
            foundRows(0).Item("方位角(°)") = CType(Text10.Text, Double)
        End If
        DataGridView1.ResetBindings()
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '点击“删除”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        BindingSource1.Current().delete()
    End Sub

    '*********************************************************************************************************************************************
    'DataGridView1选中行发生改变事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        If IsNothing(Me.BindingSource1.Current) Then
            Text7.Text = CStr(0)
            Text8.Text = CStr(0.0#)
            Text9.Text = CStr(0.0#)
            Text10.Text = CStr(0.0#)
        Else
            If IsDBNull(Me.BindingSource1.Current("序   号")) Then
                Text7.Text = CStr(0)
                Text8.Text = CStr(0.0#)
                Text9.Text = CStr(0.0#)
                Text10.Text = CStr(0.0#)
            Else
                Text7.Text = CStr(CType(BindingSource1.Current("序   号"), Integer))
                Text8.Text = CStr(CType(BindingSource1.Current("井  深(m)"), Double))
                Text9.Text = CStr(CType(BindingSource1.Current("井斜角(°)"), Double))
                Text10.Text = CStr(CType(BindingSource1.Current("方位角(°)"), Double))
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击“用将Excel作为ODBC数据源的方法导入...”按钮事件
    '    用把Excel文件作为一个ODBC数据源的方法
    '*********************************************************************************************************************************************
    Private Sub cmdInput_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInput.Click
        Dim sheet As String
        Dim excelpath As String
        Dim filepath As String
        Dim excel_ConString As String
        Dim excel_cn As System.Data.OleDb.OleDbConnection
        Dim excel_ad As New System.Data.OleDb.OleDbDataAdapter
        Dim excel_tb As New DataTable
        Dim row As DataRow
        Dim tw_newRow As DataRow
        Dim recCount As Integer
        Dim quitcount As Integer
        Dim xh As Integer
        Dim jsh As Double
        Dim jxj As Double
        Dim fwj As Double

        On Error GoTo errhandler
        '*********************************************************************************************************************************************
        '判断测井数据表表中是否有数据，若没有，执行导入，若有，提问是否重新导入
        '*********************************************************************************************************************************************
        recCount = dst.Tables("Testwell_Table").Rows.Count
        If recCount > 0 Then
            msg_prompt = CStr(well_name) & "井斜数据表已有数据，是否要重新导入?"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                Exit Sub
            End If
        End If
        '*********************************************************************************************************************************************
        '向用户要文件名
        '*********************************************************************************************************************************************
        excelpath = ""
        filepath = My.Application.Info.DirectoryPath & "\算例"
        OpenFileDialog1.InitialDirectory = filepath & "\"
        OpenFileDialog1.Filter = "油井文件(*.xls)|*.xls|All Files|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.Title = "导入"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            excelpath = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If
        If Dir(excelpath) = "" Then
            MsgBox("找不到文件！", , "系统提示")
            Exit Sub
        End If
        '*********************************************************************************************************************************************
        '将井斜数据通过excel_ad读入到表excel_tb中
        '*********************************************************************************************************************************************
        '用于 Extended Properties 值的有效 Excel 版本。
        '    对于 Microsoft Excel 8.0 (97)、9.0 (2000) 和 10.0 (2002) 工作簿，请使用 Excel 8.0。 
        '    对于 Microsoft Excel 5.0 和 7.0 (95) 工作簿，请使用 Excel 5.0。 
        '    对于 Microsoft Excel 5.0 工作簿，请使用 Excel 4.0。 
        '    对于 Microsoft Excel 3.0 工作簿，请使用 Excel 3.0。
        '*********************************************************************************************************************************************
        excel_ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source=  " & excelpath & ";Extended Properties='Excel 8.0;HDR=Yes'"
        If RadioButton1.Checked = True Then
            excel_ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source=  " & excelpath & ";Extended Properties='Excel 8.0;HDR=Yes'"
        End If
        If RadioButton2.Checked = True Then
            excel_ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source=  " & excelpath & ";Extended Properties='Excel 5.0;HDR=Yes'"
        End If
        If RadioButton3.Checked = True Then
            excel_ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source=  " & excelpath & ";Extended Properties='Excel 4.0;HDR=Yes'"
        End If
        If RadioButton4.Checked = True Then
            excel_ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source=  " & excelpath & ";Extended Properties='Excel 3.0;HDR=Yes'"
        End If
        excel_cn = New System.Data.OleDb.OleDbConnection(excel_ConString)
        excel_cn.Open()
        sheet = "井斜数据表"
        SQL_command = "Select * FROM [" & sheet & "$]"
        '生成命令并执行
        excel_ad.SelectCommand = New OleDbCommand(SQL_command, excel_cn)
        excel_ad.Fill(excel_tb)
        excel_cn.Close()
        If excel_tb.Rows.Count = 0 Then
            MsgBox("没读到数据，请检查文件中是否有内容或文件是否关闭！", , "系统提示")
            Exit Sub
        End If
        '*********************************************************************************************************************************************
        '清空dst.Tables("Testwell_Table")
        '*********************************************************************************************************************************************
        For Each row In dst.Tables("Testwell_Table").Select
            row.Delete()
        Next
        '*********************************************************************************************************************************************
        '将表excel_tb中的数据写入 dst.Tables("Testwell_Table")中,刷新显示
        '用(row.Item(0)中用数字序号可容错Excel中标题行改动。
        '*********************************************************************************************************************************************
        recCount = 0
        quitcount = 0
        For Each row In excel_tb.Select
            xh = Val(row.Item(0).ToString)
            jsh = Val(row.Item(1).ToString)
            jxj = Val(row.Item(2).ToString)
            fwj = Val(row.Item(3).ToString)
            If Not (xh = 0 And jsh = 0 And jxj = 0 And fwj = 0) Then
                tw_newRow = dst.Tables("Testwell_Table").NewRow()
                tw_newRow.Item("序   号") = xh
                tw_newRow.Item("井  深(m)") = jsh
                tw_newRow.Item("井斜角(°)") = jxj
                tw_newRow.Item("方位角(°)") = fwj
                dst.Tables("Testwell_Table").Rows.Add(tw_newRow)
                quitcount = 0
                recCount = recCount + 1
            Else
                quitcount = quitcount + 1
            End If
            If quitcount >= 10 Then
                Exit For
            End If
        Next
        DataGridView1.ResetBindings()
        DataGridView1.Refresh()
        DataGridView1.Show()
        MsgBox("共导入" & CStr(recCount) & "条记录！", , "系统提示")
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "用将Excel作为ODBC数据源的方法读取Excel文件失败，可能是office版本问题，请换另外的导入方法试试。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击“用打开EXcel对象的方法导入...”按钮事件
    '    用打开EXcel对象的方法导入Excel文件数据
    '*********************************************************************************************************************************************
    Private Sub cmdImport2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport2.Click
        Dim excelpath As String
        Dim filepath As String
        Dim row As DataRow
        Dim tw_newRow As DataRow
        Dim recCount As Integer
        Dim xh As Integer
        Dim jsh As Double
        Dim jxj As Double
        Dim fwj As Double
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim i As Integer
        Dim quitcount As Integer
        On Error GoTo errhandler
        '*********************************************************************************************************************************************
        '判断测井数据表表中是否有数据，若没有，执行导入，若有，提问是否重新导入
        '*********************************************************************************************************************************************
        recCount = dst.Tables("Testwell_Table").Rows.Count
        If recCount > 0 Then
            msg_prompt = CStr(well_name) & "井斜数据表已有数据，是否要重新导入?"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                Exit Sub
            End If
        End If
        '*********************************************************************************************************************************************
        '向用户要文件名
        '*********************************************************************************************************************************************
        excelpath = ""
        filepath = My.Application.Info.DirectoryPath & "\算例"
        OpenFileDialog1.InitialDirectory = filepath & "\"
        OpenFileDialog1.Filter = "油井文件(*.xls)|*.xls|All Files|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.Title = "导入"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            excelpath = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If
        If Dir(excelpath) = "" Then
            MsgBox("找不到文件！", , "系统提示")
            Exit Sub
        End If
        '*********************************************************************************************************************************************
        '清空dst.Tables("Testwell_Table")
        '*********************************************************************************************************************************************
        For Each row In dst.Tables("Testwell_Table").Select
            row.Delete()
        Next
        '*********************************************************************************************************************************************
        '逐行读Excel表，并写入 dst.Tables("Testwell_Table")中,刷新显示
        '*********************************************************************************************************************************************
        xlBook = xlApp.Workbooks.Open(excelpath) '打开存在的execl文件
        xlApp.Visible = False '设置EXCEL对象不可见（或可见）
        xlSheet = xlBook.Worksheets("井斜数据表") '设置活动工作表
        recCount = 0
        For i = 1 To xlSheet.UsedRange.Rows.Count
            xh = Val(xlSheet.Cells._Default(i, 1).value)
            jsh = Val(xlSheet.Cells._Default(i, 2).value)
            jxj = Val(xlSheet.Cells._Default(i, 3).value)
            fwj = Val(xlSheet.Cells._Default(i, 4).value)
            If Not (xh = 0 And jsh = 0 And jxj = 0 And fwj = 0) Then
                tw_newRow = dst.Tables("Testwell_Table").NewRow()
                tw_newRow.Item("序   号") = xh
                tw_newRow.Item("井  深(m)") = jsh
                tw_newRow.Item("井斜角(°)") = jxj
                tw_newRow.Item("方位角(°)") = fwj
                dst.Tables("Testwell_Table").Rows.Add(tw_newRow)
                recCount = recCount + 1
                quitcount = 0
            Else
                quitcount = quitcount + 1
            End If
            If quitcount >= 10 Then
                Exit For
            End If
        Next i
        xlBook.Close()
        DataGridView1.ResetBindings()
        DataGridView1.Refresh()
        DataGridView1.Show()
        MsgBox("共导入" & CStr(recCount) & "条记录！", , "系统提示")
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    用打开Excel对象的方法读取Excel文件失败，可能是你计算机上卸载了旧版本MS-Office后又新装了新版本，" _
            & "或者计算机上装过（了）WPS，反正这三者中的两个或三个在注册表里冲突了，处理起来非常复杂和怪异，请换另外的导入方法试试吧。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class
