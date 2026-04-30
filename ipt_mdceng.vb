Option Strict Off
Option Explicit On
Imports System.Data.OleDb

Friend Class ipt_mdceng
    '*********************************************************************************************************************************************
    '                                                   关于目的层输入与管理界面的说明
    ' 重要说明：本窗体的目的层被删除时，会删除所有关联于本层的作业，包括当前zuoye_name设置的作业。此时程序将全局变量zuoye_name置为空""
    ' 本窗体可能使zuoye_name变为"" 
    ' 
    '                                                                                         李润洲 2025年7月22日最后整理   
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    '目的层  李润洲2025年6月15日
    Private mdc_Table As New DataTable
    '岩石岩性  李润洲2025年6月15日
    Private ysyx_Table As New DataTable
    Private SQL_command As String
    Private midMDCDepth As Single  '目的层中部测深
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub ipt_skd_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
        zct_main.my_refresh()
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub ipt_mdceng_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        '目的层必须数据
        '位置
        Text7.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        '深度
        TextBox4.Text = CStr(0.0#)
        TextBox5.Text = CStr(0.0#)
        TextBox6.Text = CStr(0.0#)
        '地层参数
        TextBox1.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text9.Text = CStr(0.0#)
        TextBox2.Text = CStr(0.0#)
        TextBox3.Text = CStr(0.0#)
        TextBox7.Text = CStr(0.0#)

        '岩石数据
        Text13.Text = ""
        Text110.Text = CStr(0.0#)
        Text17.Text = CStr(0.0#)
        Text18.Text = CStr(0.0#)
        Text20.Text = CStr(0.0#)
        Text19.Text = ""

        TextBox8.Text = ""

        Label_0.Text = well_name & "井油气井参数"
        Label2.Text = well_name & "井作业目的层参数"
        Me.Text = well_name & "井" & "作业目的层数据输入与修改"
        '给井基本数据赋值
        Try
            Using cn_userdb As New OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位 from 油气井表 where 井号='" & well_name & "'"
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If RECreader.Read Then
                            Text1.Text = RECreader.Item("地理位置").ToString
                            Text2.Text = RECreader.Item("构造位置").ToString
                            Text3.Text = RECreader.Item("井别").ToString
                            Text4.Text = RECreader.Item("设计井深m").ToString
                            Text5.Text = RECreader.Item("完钻井深m").ToString
                            Text6.Text = RECreader.Item("完钻层位").ToString
                        End If
                    End Using
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "访问井基本数据表时出错。" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try

        '目的层数据列表DataGridView1设置并填充函数fill_skdgrid()
        Call fill_mdcgrid()
        '岩石岩性数据列表DataGridView2设置并填充函数fill_ysyxgrid()
        Call fill_ysyxgrid()
        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
    End Sub
    '*********************************************************************************************************************************************
    '目的层数据列表DataGridView1设置并填充函数fill_mdcgrid()
    '*********************************************************************************************************************************************
    Private Sub fill_mdcgrid()
        Dim i As Integer
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select * from 目的层参数表 where 井号='" & well_name & "'order by 起始深度m"
                Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                    mdc_Table.Clear()
                    ad.Fill(mdc_Table)
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "访问井目的层参数表时出错。" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try

        BindingSource1.DataSource = mdc_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Width = (DataGridView1.Width - 26) * 0.05
        Next i
        DataGridView1.Refresh()
        DataGridView1.Show()
        Call set_mdc_texts()
    End Sub
    '*********************************************************************************************************************************************
    '岩石岩性数据列表设置并填充函数fill_ysyxgrid()----DataGridView2
    '*********************************************************************************************************************************************
    Private Sub fill_ysyxgrid()
        Try
            Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                cn_basedb.Open()
                SQL_command = "select 岩石岩性,岩石强度Mpa,压力敏感系数,变形系数,弹性模量Mpa,泊松比,备注  from 岩石数据表 order by 岩石强度Mpa"
                Using ad As New OleDbDataAdapter(SQL_command, cn_basedb)
                    ysyx_Table.Clear()
                    ad.Fill(ysyx_Table)
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "访问井岩石数据表时出错。" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
        BindingSource2.DataSource = ysyx_Table
        DataGridView2.ClearSelection()
        DataGridView2.DataSource = BindingSource2
        DataGridView2.ResetBindings()
        DataGridView2.AutoGenerateColumns = True
        DataGridView2.AllowUserToAddRows = False
        DataGridView2.AllowUserToDeleteRows = False
        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.MultiSelect = False
        DataGridView2.RowHeadersWidth = 24
        DataGridView2.ReadOnly = True
        DataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Dim i As Integer
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(i).Width = (DataGridView2.Width - 26) * 0.14
        Next i
        DataGridView2.Refresh()
        DataGridView2.Show()
        Call set_ys_texts()

    End Sub

    '*********************************************************************************************************************************************
    'Timer1的Tick事件处理：（1）画井身结构图；（2）关闭计时器
    '*********************************************************************************************************************************************
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '绘制井身结构图,含管柱
        Call drawWellStruction(Picture1, 2)
        Timer1.Stop()
    End Sub
    '*********************************************************************************************************************************************
    '点击“绘图”按钮事件
    '*********************************************************************************************************************************************
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '绘制井身结构图,含管柱
        Call drawWellStruction(Picture1, 2)
    End Sub
    '*********************************************************************************************************************************************
    '点击"退出[&X]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击"计算[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim hasCjdata As Boolean = True
        If Val(TextBox4.Text) <= 0 Or Not IsNumeric(TextBox4.Text) Or Val(TextBox4.Text) > Val(TextBox5.Text) Then
            msg_prompt = "请输入合适的起始深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(TextBox5.Text) <= 0 Or Not IsNumeric(TextBox5.Text) Or Val(TextBox5.Text) < Val(TextBox4.Text) Then
            msg_prompt = "请输入合适的终止深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(TextBox1.Text) <= 0 Or Not IsNumeric(TextBox1.Text) Then
            msg_prompt = "请输入合适的地层压力系数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select  * from 测井数据表 where 井号='" & well_name & "' order by [井  深(m)]"
            Using oleDbcmd As New OleDbCommand(SQL_command, cn_userdb)
                Using dreader As OleDbDataReader = oleDbcmd.ExecuteReader()
                    If Not dreader.Read Then
                        msg_prompt = "没有测井数据，将按照直井计算地层中部垂深与压力。"
                        msg_buttons = 0 + 48
                        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        hasCjdata = False
                    End If
                End Using
            End Using
        End Using
        '目的层中部井深m
        midMDCDepth = ((Val(TextBox4.Text) + Val(TextBox5.Text)) / 2).ToString
        '目的层中部垂深m
        Dim depth As Double
        If Not hasCjdata Then
            depth = midMDCDepth
        Else
            Dim jxj, fwj As Double
            Call M_cal_jx_fw_cs.cal_jx_fw_cs(midMDCDepth, jxj, fwj, depth)
            TextBox6.Text = Math.Round(depth, 4).ToString()
        End If
        '目的层压力MPa=
        Text10.Text = Math.Round(Val(TextBox1.Text) * depth * 9.81 * 10 ^ (-3), 4).ToString()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中目的层列表中的某行。
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Call set_mdc_texts()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的射孔段列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub set_mdc_texts()
        If (Not IsNothing(Me.BindingSource1.Current)) AndAlso (Not IsDBNull(Me.BindingSource1.Current("起始深度m"))) Then
            Text7.Text = If(BindingSource1.Current("目的层名称") Is DBNull.Value, "", BindingSource1.Current("目的层名称").ToString)
            TextBox4.Text = If(BindingSource1.Current("起始深度m") Is DBNull.Value, "0.0", BindingSource1.Current("起始深度m").ToString)
            TextBox5.Text = If(BindingSource1.Current("终止深度m") Is DBNull.Value, "0.0", BindingSource1.Current("终止深度m").ToString)
            Text9.Text = If(BindingSource1.Current("地层温度℃") Is DBNull.Value, "0.0", BindingSource1.Current("地层温度℃").ToString)
            Text10.Text = If(BindingSource1.Current("地层压力MPa") Is DBNull.Value, "0.0", BindingSource1.Current("地层压力MPa").ToString)
            TextBox6.Text = If(BindingSource1.Current("目的层中部垂深m") Is DBNull.Value, "0.0", BindingSource1.Current("目的层中部垂深m").ToString)
            TextBox9.Text = If(BindingSource1.Current("地层系") Is DBNull.Value, "", BindingSource1.Current("地层系").ToString)
            TextBox10.Text = If(BindingSource1.Current("地层组") Is DBNull.Value, "", BindingSource1.Current("地层组").ToString)
            TextBox11.Text = If(BindingSource1.Current("地层段") Is DBNull.Value, "", BindingSource1.Current("地层段").ToString)
            TextBox1.Text = If(BindingSource1.Current("地层压力系数") Is DBNull.Value, "0.0", BindingSource1.Current("地层压力系数").ToString)
            TextBox2.Text = If(BindingSource1.Current("地破压力MPa") Is DBNull.Value, "0.0", BindingSource1.Current("地破压力MPa").ToString)
            Text13.Text = If(BindingSource1.Current("岩石岩性") Is DBNull.Value, "", BindingSource1.Current("岩石岩性").ToString)
            Text17.Text = If(BindingSource1.Current("弹性模量MPa") Is DBNull.Value, "0.0", BindingSource1.Current("弹性模量MPa").ToString)
            Text110.Text = If(BindingSource1.Current("岩石强度MPa") Is DBNull.Value, "0.0", BindingSource1.Current("岩石强度MPa").ToString)
            Text21.Text = If(BindingSource1.Current("变形系数") Is DBNull.Value, "0.0", BindingSource1.Current("变形系数").ToString)
            Text20.Text = If(BindingSource1.Current("压力敏感系数") Is DBNull.Value, "0.0", BindingSource1.Current("压力敏感系数").ToString)
            Text18.Text = If(BindingSource1.Current("泊松比") Is DBNull.Value, "0.0", BindingSource1.Current("泊松比").ToString)
            TextBox3.Text = If(BindingSource1.Current("孔隙度%") Is DBNull.Value, "0.0", BindingSource1.Current("孔隙度%").ToString)
            TextBox7.Text = If(BindingSource1.Current("渗透率um2") Is DBNull.Value, "0.0", BindingSource1.Current("渗透率um2").ToString)
            midMDCDepth = If(BindingSource1.Current("目的层中部井深m") Is DBNull.Value, "0.0", BindingSource1.Current("目的层中部井深m").ToString)
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用岩石"按钮
    '*********************************************************************************************************************************************
    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
        Call set_ys_texts()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView2单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        Call set_ys_texts()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的岩石列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub set_ys_texts()
        If IsNothing(Me.BindingSource2.Current) Then
            msg_prompt = "请从列表中选择岩石数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource2.Current("岩石岩性")) Then
                msg_prompt = "请从列表中选择岩石。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text13.Text = If(BindingSource2.Current("岩石岩性") Is DBNull.Value, "", BindingSource2.Current("岩石岩性").ToString)
                Text17.Text = If(BindingSource2.Current("弹性模量MPa") Is DBNull.Value, "0.0", BindingSource2.Current("弹性模量MPa").ToString)
                Text110.Text = If(BindingSource2.Current("岩石强度MPa") Is DBNull.Value, "0.0", BindingSource2.Current("岩石强度MPa").ToString)
                Text21.Text = If(BindingSource2.Current("变形系数") Is DBNull.Value, "0.0", BindingSource2.Current("变形系数").ToString)
                Text20.Text = If(BindingSource2.Current("压力敏感系数") Is DBNull.Value, "0.0", BindingSource2.Current("压力敏感系数").ToString)
                Text18.Text = If(BindingSource2.Current("泊松比") Is DBNull.Value, "0.0", BindingSource2.Current("泊松比").ToString)
            End If
        End If
    End Sub

    '*********************************************************************************************************************************************
    '点击"帮助[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除[&D]"按钮
    '删除datagridview1中选定的目的层数据。重点：删除时关联删除目的层上的所有作业
    '*********************************************************************************************************************************************
    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Dim msg As String = ""
        Dim isCurZuoyeDeleted = False
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select * from 目的层参数表 where 井号='" & well_name & "' and  目的层名称='" & Trim(Text7.Text) & "'"
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If RECreader.Read Then
                            msg_prompt = "是否确定要删除所选的目的层数据？目的层关联的作业将全部删除！"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            '找到目的层关联的作业.删除所有和目的层关联的作业
                            SQL_command = "select * from 作业地层参数表  where 井号='" & well_name & "' and  目的层名称='" & Trim(Text7.Text) & "'"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                    Dim del_zyName As String
                                    While RECreader2.Read
                                        del_zyName = RECreader2.Item("作业名称")
                                        Dim table_count As Short = delete_zy(del_zyName)
                                        msg = msg & Chr(13) & Chr(10) & del_zyName & "作业共删除" & table_count & "个表的记录；"
                                        If del_zyName = zuoye_name Then
                                            isCurZuoyeDeleted = True
                                        End If
                                    End While
                                End Using
                            End Using
                            '删除目的层表中的当前目的层
                            SQL_command = "delete * from 目的层参数表 where 井号='" & well_name & "' and  目的层名称='" & Trim(Text7.Text) & "'"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            Call fill_mdcgrid()
                            Call drawWellStruction(Picture1, 2)
                            msg_prompt = "目的层数据删除完成。" & msg
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        Else
                            msg_prompt = "未找到目的层名称对应的数据，请选好要删除的数据！"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        End If
                    End Using
                End Using
            End Using
            Call fill_mdcgrid()
            If isCurZuoyeDeleted Then
                zuoye_name = ""
                'GlobalVariables.zuoye_name = zuoye_name
            End If
        Catch ex As OleDbException
            msg_prompt = "目的层及其关联数据删除时出错。" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
    End Sub
    '*********************************************************************************************************************************************
    '"删除指定作业的所有数据
    '返回删除的表的个数
    ' 20251112,秦彦斌将delete_zy函数移至data_str.vb中
    '*********************************************************************************************************************************************
    'Public Shared Function delete_zy(ByVal zyname As String) As Integer
    '    Dim table_name As String
    '    Dim col_name As String
    '    Dim i As Short = 0
    '    Dim hasZYField As Boolean
    '    Dim hasJHField As Boolean
    '    Dim SQL_command As String
    '    '***********************************************************************************
    '    ' 遍历库中所有表，删除 井号=well_name、作业名称=zyname 的所有记录
    '    '***********************************************************************************
    '    i = 0
    '    Try
    '        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
    '            cn_userdb.Open()
    '            Using dbSchema As DataTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
    '                If dbSchema.Rows.Count > 0 Then
    '                    For Each dbschema_row As DataRow In dbSchema.Rows
    '                        hasZYField = False
    '                        hasJHField = False
    '                        table_name = dbschema_row.Item("TABLE_NAME").ToString
    '                        Using columnTable As DataTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
    '                            For Each columnTab_row As DataRow In columnTable.Rows
    '                                col_name = columnTab_row.Item("COLUMN_NAME").ToString
    '                                If col_name = "作业名称" Then
    '                                    hasZYField = True
    '                                End If
    '                                If col_name = "井号" Then
    '                                    hasJHField = True
    '                                End If
    '                            Next
    '                            If hasZYField And hasJHField Then
    '                                '如果表有“作业名称”和“井号”字段，且存有“井号=well_name、作业名称=zyname”的记录，删除记录
    '                                SQL_command = "select * from " & table_name & " where 井号='" & well_name & "' And 作业名称='" & zyname & "'"
    '                                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
    '                                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
    '                                        If RECreader.Read Then
    '                                            SQL_command = "delete * from " & table_name & " where 井号='" & well_name & "' And 作业名称='" & zyname & "'"
    '                                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
    '                                                EXECOleDbCommand2.ExecuteNonQuery()
    '                                            End Using
    '                                            i = i + 1
    '                                        End If
    '                                    End Using
    '                                End Using
    '                            End If 'hasZYField And hasJHField 判断结束
    '                        End Using
    '                    Next
    '                End If
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        msg_prompt = "删除作业‘" & zyname & "’时出错，请检查数据表内容！"
    '        msg_buttons = 0 + 48
    '        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    '    Finally
    '        delete_zy = i
    '    End Try
    'End Function
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        '目的层名称不为空检查
        If Not CheckWigetData(Text7, "目的层名称", False) Then Exit Sub
        If Val(TextBox4.Text) < 0 Or Not IsNumeric(TextBox4.Text) Or Val(TextBox4.Text) > Val(TextBox5.Text) Then
            msg_prompt = "请输入合适的起始深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(TextBox5.Text) <= 0 Or Not IsNumeric(TextBox5.Text) Or Val(TextBox5.Text) < Val(TextBox4.Text) Then
            msg_prompt = "请输入合适的终止深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '压力系数为数字、不为空、>0检查
        If Not CheckWigetData(TextBox1, "地层压力系数", True, False) Then Exit Sub
        If Not CheckWigetData(Text9, "地层温度", True, False) Then Exit Sub
        If Not CheckWigetData(Text10, "地层压力", True, False) Then Exit Sub
        If Not CheckWigetData(TextBox6, "目的层中部垂深", True, False) Then Exit Sub
        If Not CheckWigetData(TextBox3, "孔隙度", True, False) Then Exit Sub
        If Not CheckWigetData(TextBox3, "渗透率", True, False) Then Exit Sub
        '地破压力为数字检查
        If Not CheckWigetData(TextBox2, "地破压力") Then Exit Sub
        '目的层岩石岩性不为空检查
        If Not CheckWigetData(Text13, "目的层岩石岩性", False) Then Exit Sub
        '为数字、不为空、>0检查
        If Not CheckWigetData(Text110, "岩石强度", True, False) Then Exit Sub
        If Not CheckWigetData(Text17, "弹性模量", True, False) Then Exit Sub
        If Not CheckWigetData(Text21, "变形系数率", True, False) Then Exit Sub
        If Not CheckWigetData(Text20, "压力敏感系数", True, False) Then Exit Sub
        If Not CheckWigetData(Text18, "泊松比", True, False) Then Exit Sub
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select * from 目的层参数表 where 井号='" & well_name & "' and  目的层名称='" & CStr(Text7.Text) & "'"
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If Not RECreader.Read Then
                            msg_prompt = "是否要建立新的目的层数据？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            SQL_command = "insert into 目的层参数表(目的层名称,起始深度m,终止深度m,地层温度℃,地层压力MPa,岩石岩性, 岩石强度MPa, 弹性模量MPa, " _
                                          & "变形系数,[孔隙度%],渗透率um2,泊松比,压力敏感系数,地层系,地层组,地层段,最大地应力MPa,最大地应力方向°,井号,目的层中部井深m," _
                                          & "目的层中部垂深m,地层压力系数,地破压力MPa,备注) values (" _
                                          & "'" & Trim(Text7.Text) & "'," & TextBox4.Text & "," & TextBox5.Text & "," & Text9.Text & "," & Text10.Text & ",'" _
                                          & Text13.Text & "'," & Text110.Text & "," & Text17.Text & "," & Text21.Text & "," & TextBox3.Text & "," & TextBox7.Text & "," & Text18.Text _
                                          & "," & Text20.Text & ",'" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "',0,0," _
                                          & "'" & well_name & "'," & midMDCDepth.ToString & "," & TextBox6.Text & "," & TextBox1.Text & "," & TextBox2.Text & ",'" & TextBox8.Text & "')"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            msg_prompt = "新的目的层数据保存完成。"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        Else
                            msg_prompt = "是否要保存对所选目的层数据的修改？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            SQL_command = "update 目的层参数表 set 起始深度m=" & TextBox4.Text & ",终止深度m=" & TextBox5.Text & ",地层温度℃=" & Text9.Text & ",地层压力MPa=" & Text10.Text _
                                          & ",岩石岩性='" & Text13.Text & "',岩石强度MPa=" & Text110.Text _
                                          & ",弹性模量MPa=" & Text17.Text & ",变形系数=" & Text21.Text & ",[孔隙度%]=" & TextBox3.Text & ",渗透率um2=" & TextBox7.Text & ",泊松比=" & Text18.Text _
                                          & ",压力敏感系数=" & Text20.Text & ",地层系='" & TextBox9.Text & "',地层组='" & TextBox10.Text & "',地层段='" & TextBox11.Text _
                                          & "',最大地应力MPa=0,最大地应力方向°=0,目的层中部井深m=" & midMDCDepth.ToString() & ",目的层中部垂深m=" & TextBox6.Text & ",地层压力系数=" & TextBox1.Text _
                                          & ",地破压力MPa=" & TextBox2.Text & ",备注='" & TextBox8.Text & "' where 井号='" & well_name & "' and  目的层名称='" & Trim(Text7.Text) & "'"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            msg_prompt = "目的层数据保存完成。"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        End If
                    End Using
                End Using
            End Using
            Call save_case()
            Call fill_mdcgrid()
        Catch ex As OleDbException
            msg_prompt = "保存数据出错,请输入正确合理的数据！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
        Call drawWellStruction(Picture1, 2)
    End Sub
    Private Sub save_case()
        Dim SQL_command As String
        Dim bzh As String
        Dim tishi As String
        Dim rec_value As String
        Dim basedb_chg As Boolean

        bzh = (Today) & "输入" & well_name & "井数据时加入。"
        basedb_chg = False

        Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
            cn_basedb.Open()
            ' 保存岩石数据
            SQL_command = "select * from 岩石数据表 " & " where 岩石岩性='" & Trim(Text13.Text) & "'"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_basedb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    If Not RECreader.Read Then
                        msg_prompt = "是否将此岩石加入到岩石数据库中？"
                        msg_buttons = 4 + 32
                        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        If msg_return = 6 Then
                            SQL_command = "insert into 岩石数据表(岩石岩性,岩石强度Mpa,变形系数,压力敏感系数,弹性模量Mpa,泊松比,备注) values (" _
                            & "'" & Trim(Text13.Text) & "'," & Text110.Text & "," & Text21.Text & "," & Text20.Text & "," & Text17.Text & "," _
                            & Text18.Text & ",'" & Trim(bzh) & "')"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_basedb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                                basedb_chg = True
                            End Using
                        End If
                    Else
                        '岩石数据表及界面中的某一数据均不为空或0，但不相等，应提示用户是否用界面数据更新油管库数据。
                        tishi = ""
                        rec_value = RECreader.Item("岩石强度Mpa").ToString
                        If rec_value <> "" And Trim(Text110.Text) <> "" And rec_value <> Text110.Text Then
                            tishi = "岩石强度：数据库中为‘" & rec_value & "’，界面中为‘" & Trim(Text110.Text) & "’"
                        End If
                        rec_value = RECreader.Item("弹性模量Mpa").ToString
                        If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text17.Text) <> 0.0# And rec_value <> Text17.Text Then
                            If tishi = "" Then
                                tishi = "弹性模量：数据库中=" & rec_value & "，界面中=" & CStr(Text17.Text)
                            Else
                                tishi = tishi & Chr(13) & Chr(10) & "弹性模量：数据库中=" & rec_value & "，界面中=" & CStr(Text17.Text)
                            End If
                        End If
                        rec_value = RECreader.Item("泊松比").ToString
                        If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text18.Text) <> 0.0# And rec_value <> Text18.Text Then
                            If tishi = "" Then
                                tishi = "泊松比：数据库中=" & rec_value & "，界面中=" & CStr(Text21.Text)
                            Else
                                tishi = tishi & Chr(13) & Chr(10) & "泊松比：数据库中=" & rec_value & "，界面中=" & CStr(Text18.Text)
                            End If
                        End If
                        rec_value = RECreader.Item("变形系数").ToString
                        If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text21.Text) <> 0.0# And rec_value <> Text21.Text Then
                            If tishi = "" Then
                                tishi = "变形系数：数据库中=" & rec_value & "，界面中=" & CStr(Text21.Text)
                            Else
                                tishi = tishi & Chr(13) & Chr(10) & "变形系数：数据库中=" & rec_value & "，界面中=" & CStr(Text21.Text)
                            End If
                        End If
                        rec_value = RECreader.Item("压力敏感系数").ToString
                        If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text20.Text) <> 0.0# And rec_value <> Text20.Text Then
                            If tishi = "" Then
                                tishi = "压力敏感系数：数据库中=" & rec_value & "，界面中=" & CStr(Text20.Text)
                            Else
                                tishi = tishi & Chr(13) & Chr(10) & "压力敏感系数：数据库中=" & rec_value & "，界面中=" & CStr(Text20.Text)
                            End If
                        End If

                        If tishi <> "" Then
                            msg_prompt = "岩石数据库中下列数据与界面中的值不同：" & Chr(13) & Chr(10) & tishi & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "是否用界面中的数据更新岩石数据库？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return = 6 Then
                                bzh = (Today) & "输入" & well_name & "井数据时最后更新。"
                                SQL_command = "update 岩石数据表 set [备注]='" & bzh & "',[岩石强度Mpa]=" & Trim(Text110.Text) & "," & " [变形系数]=" & Text21.Text _
                                    & ",[压力敏感系数]=" & Text20.Text & "," & " [弹性模量Mpa]=" & Text17.Text & ",[泊松比]=" & Text18.Text _
                                    & " where [岩石岩性]='" & Trim(Text13.Text)
                                Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_basedb)
                                    EXECOleDbCommand2.ExecuteNonQuery()
                                    basedb_chg = True
                                End Using

                            End If
                        End If
                    End If
                End Using
            End Using
        End Using
        If basedb_chg = True Then
            Call fill_ysyxgrid()
        End If
    End Sub
End Class
'为简化控件输入验证过程而设计的公共函数
Public Module ControlValidate
    '''<summary>
    '''验证控件输入空、数字、不能小于等于0的有效性
    '''</summary>
    '''<param name="ctl">控件</param>
    '''<param name="name">控件名称字符串</param>
    '''<param name="isNum">是否要求为数字</param>
    '''<param name="可选参数canSmallZero">是否可以小于0，默认不可以</param>
    '''<returns>Boolean</returns>
    Public Function CheckWigetData(ByRef ctl As System.Windows.Forms.Control, ByVal name As String, ByVal isNum As Boolean, Optional ByVal canSmallZero As Boolean = False) As Boolean
        If IsTextNotEmpty(ctl.Text.ToString()) = False Then
            MessageBox.Show(name & "不能为空")
            Return False
        End If
        If isNum And IsNumeric(ctl.Text.ToString()) = False Then
            MessageBox.Show(name & "请填数值")
            Return False
        End If
        If isNum AndAlso canSmallZero = False Then
            Dim sg = Convert.ToSingle(ctl.Text.ToString())
            If sg <= 0 Then
                MessageBox.Show(name & "数值不能小于等于0")
                Return False
            End If
        End If
        Return True
    End Function
    '''<summary>
    '''只验证控件输入是否为数字
    '''</summary>
    '''<param name="ctl">控件</param>
    '''<param name="name">控件名称字符串</param>
    '''<returns>Boolean</returns>
    Public Function CheckWigetData(ByRef ctl As System.Windows.Forms.Control, ByVal name As String) As Boolean
        If IsNumeric(ctl.Text.ToString()) = False Then
            MessageBox.Show(name & "请填数值")
            Return False
        End If
        Return True
    End Function
    '判断字符串是否不为空
    Private Function IsTextNotEmpty(ByRef str As String) As Boolean
        Return str <> Nothing And Len(str) <> 0
    End Function
End Module
