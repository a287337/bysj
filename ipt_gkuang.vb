Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class ipt_gkuang
    '*********************************************************************************************************************************************
    '                                                   关于工况数据输入与管理界面的说明
    ' 程序升级记事：
    '                                                                                              秦彦斌 2021年12月3日最后整理更新
    '  （1）代码实现让SSTab各页标题宽度一样，不要挤在一起。
    '  （2）综合运用访问数据库的各种方法：
    '       <1> 读到内存DataSet的Table中，用行访问。用到OleDbDataAdapter
    '       <2> 仅读一遍，多条记录用 RECreader。
    '       <3> 执行无返回参数命令用EXECOleDbCommand.ExecuteNonQuery()
    '   (3) 界面重新排列，可最发化。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub ipt_gkuang_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim fgq_newRow As DataRow
        Dim mdgj_newRow As DataRow
        Dim kggj_newRow As DataRow
        Label(9).Text = well_name & "井" & zuoye_name & "作业工况参数"
        Me.Text = well_name & "井" & zuoye_name & "作业工况参数数据输入与修改"
        '***************************************************************************************************************
        '初始化各变量
        '***************************************************************************************************************
        Text53.Text = CStr(0.0#)
        Text54.Text = CStr(20.0#)
        Text7.Text = CStr(20.0#)
        Text10.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text56.Text = CStr(0.0#)
        Text57.Text = CStr(0.0#)
        Text58.Text = CStr(0.0#)
        Text59.Text = CStr(0.0#)
        Text60.Text = CStr(0.0#)
        Text61.Text = ""
        '***************************************************************************************************************
        '根据机械设计手册：
        '    钢-钢     无润滑  静摩擦系数  0.15，   有润滑  静摩擦系数  0.1-0.12
        '    钢-钢     无润滑  动摩擦系数  0.1，    有润滑  动摩擦系数  0.05-0.1
        ' 故这里取值为  0.08
        '***************************************************************************************************************
        Text63.Text = CStr(0.08)
        Text64.Text = CStr(0.0#)
        Text65.Text = CStr(0.0#)
        Text66.Text = CStr(0.0#)
        Text67.Text = CStr(0.0#)
        Text68.Text = CStr(0.0#)
        'Combo1.text = "无"
        Combo2.Text = "不流动"
        Combo3.Text = "不流动"
        Text8.Text = CStr(0)
        Text18.Text = CStr(0)
        Combo5.Text = "计算"
        Combo6.Text = "计算"
        Combo8.Text = "输入"
        Combo9.Text = "输入"
        ComboBox1.Text = "无"
        ComboBox2.Text = "无"
        Text6.Text = CStr(0.0#)
        Text12.Text = CStr(0.0#)
        Text15.Text = CStr(0.0#)
        Text16.Text = CStr(0.0#)
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text13.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text19.Text = CStr(0.35)
        Text20.Text = CStr(0.35)
        Text21.Text = CStr(0.0#)
        Text22.Text = CStr(0.0#)
        Text23.Text = CStr(0.0#)
        TextBox2.Text = CStr(0.0#)
        Call opt3_chg()
        Call opt4_chg()
        Combo7.Text = ""
        '***************************************************************************************************************
        '设计井深m、完钻井深m 只读文本框赋值
        '***************************************************************************************************************
        Text4.Text = ""
        Text5.Text = ""
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位 from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.Read Then
            Text4.Text = RECreader.Item("设计井深m").ToString
            Text5.Text = RECreader.Item("完钻井深m").ToString
        End If
        Text9.Text = ""
        Text1.Text = ""
        TextBox1.Text = ""
        RECreader.Close()
        SQL_command = "select sum(元件长度m) as 管柱总长 from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.Read Then
            Text9.Text = RECreader.Item("管柱总长").ToString()
            Text1.Text = RECreader.Item("管柱总长").ToString()
            TextBox1.Text = Text1.Text
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 工况名称 order by 排序"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        ListBox1.Items.Clear()
        While RECreader.Read
            ListBox1.Items.Add(RECreader.Item("工况名称").ToString)
        End While
        SQL_command = "select * from 封隔器定位方式"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        Combo4.Items.Clear()
        Combo1.Items.Clear()
        While RECreader.Read
            Combo4.Items.Add(RECreader.Item("定位方式").ToString)
            Combo1.Items.Add(RECreader.Item("定位方式").ToString)
        End While
        SQL_command = "select * from 开关元件状态"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        Combo7.Items.Clear()
        While RECreader.Read
            Combo7.Items.Add(RECreader.Item("开关元件状态").ToString)
        End While
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        ''*********************************************************************************************************************************************
        ''（1）让SSTab2各页标题宽度一样，不要挤在一起。
        ''*********************************************************************************************************************************************
        ''这个Fixed设置是必须的
        'SSTab2.SizeMode = TabSizeMode.Fixed
        ''设置标签宽度
        'totalWidth = SSTab2.Width
        'pageCount = SSTab2.TabPages.Count
        ''最后-1 因为tabcontrol有留margin，得空出margin的空间
        'pageWidth = totalWidth / pageCount - 1
        ''ItemSize在SizeMode = TabSizeMode.Fixed才生效
        'SSTab2.ItemSize = New Size(pageWidth, SSTab2.ItemSize.Height)
        ''*********************************************************************************************************************************************
        '工况序号及名称列表DataGridView1设置并填充函数fill_gk_grid1()，原来用DataGrid8和Adodc13
        '*********************************************************************************************************************************************
        Call fill_gk_grid1()
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '***************************************************************************************************************
        '初始化封隔定位元件状态显示：DataGridView2，原来用lsb_gk_fgq、Adodc14、DataGrid1
        '***************************************************************************************************************
        SQL_command = "select 元件序号,元件名称,封隔器状态,定位方式 from 工况_封隔定位元件 where 0>1"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_fgq_Table.Clear()
        ad.Fill(gk_fgq_Table)
        ad.Dispose()
        SQL_command = "select 元件序号,元件名称 from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            fgq_newRow = dst.Tables("gk_fgq_Table").NewRow()
            fgq_newRow.Item("元件序号") = RECreader.Item("元件序号").ToString
            fgq_newRow.Item("元件名称") = RECreader.Item("元件名称").ToString
            fgq_newRow.Item("封隔器状态") = "未坐封"
            fgq_newRow.Item("定位方式") = "未(无)锚定"
            dst.Tables("gk_fgq_Table").Rows.Add(fgq_newRow)
        End While
        RECreader.Close()
        BindingSource2.DataSource = gk_fgq_Table
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
        DataGridView2.Columns(0).Width = (DataGridView2.Width - 26) * 0.2
        DataGridView2.Columns(1).Width = (DataGridView2.Width - 26) * 0.4
        DataGridView2.Columns(2).Width = (DataGridView2.Width - 26) * 0.2
        DataGridView2.Columns(3).Width = (DataGridView2.Width - 26) * 0.2
        DataGridView2.Refresh()
        DataGridView2.Show()
        '***************************************************************************************************************
        '初始化锚定工具工作状态显示：DataGridView3，原来用lsb_gk_mdgj、Adodc2、DataGrid2
        '***************************************************************************************************************
        SQL_command = "select 元件序号,元件名称,定位方式  from 工况_锚定元件 where 0>1"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_mdgj_Table.Clear()
        ad.Fill(gk_mdgj_Table)
        ad.Dispose()
        SQL_command = "select 元件序号,元件名称 from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            mdgj_newRow = dst.Tables("gk_mdgj_Table").NewRow()
            mdgj_newRow.Item("元件序号") = RECreader.Item("元件序号").ToString
            mdgj_newRow.Item("元件名称") = RECreader.Item("元件名称").ToString
            mdgj_newRow.Item("定位方式") = "未(无)锚定"
            dst.Tables("gk_mdgj_Table").Rows.Add(mdgj_newRow)
        End While
        RECreader.Close()
        BindingSource3.DataSource = gk_mdgj_Table
        DataGridView3.ClearSelection()
        DataGridView3.DataSource = BindingSource3
        DataGridView3.ResetBindings()
        DataGridView3.AutoGenerateColumns = True
        DataGridView3.AllowUserToAddRows = False
        DataGridView3.AllowUserToDeleteRows = False
        DataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView3.MultiSelect = False
        DataGridView3.RowHeadersWidth = 24
        DataGridView3.ReadOnly = True
        DataGridView3.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView3.Columns(0).Width = (DataGridView3.Width - 26) * 0.15
        DataGridView3.Columns(1).Width = (DataGridView3.Width - 26) * 0.6
        DataGridView3.Columns(2).Width = (DataGridView3.Width - 26) * 0.25
        DataGridView3.Refresh()
        DataGridView3.Show()
        '***************************************************************************************************************
        '初始化开关工具工作状态显示：DataGridView4，原来用lsb_gk_kgyj、Adodc15、DataGrid10
        '***************************************************************************************************************
        SQL_command = "select 元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 工况_开关元件 where 0>1"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_kggj_Table.Clear()
        ad.Fill(gk_kggj_Table)
        ad.Dispose()
        SQL_command = "select 元件序号,元件名称 from 管柱_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            kggj_newRow = dst.Tables("gk_kggj_Table").NewRow()
            kggj_newRow.Item("元件序号") = RECreader.Item("元件序号").ToString
            kggj_newRow.Item("元件名称") = RECreader.Item("元件名称").ToString
            kggj_newRow.Item("开关状态") = ""
            kggj_newRow.Item("管内嘴损压差MPa") = 0.0
            kggj_newRow.Item("油套嘴损压差MPa") = 0.0
            dst.Tables("gk_kggj_Table").Rows.Add(kggj_newRow)
        End While
        RECreader.Close()
        BindingSource4.DataSource = gk_kggj_Table
        DataGridView4.ClearSelection()
        DataGridView4.DataSource = BindingSource4
        DataGridView4.ResetBindings()
        DataGridView4.AutoGenerateColumns = True
        DataGridView4.AllowUserToAddRows = False
        DataGridView4.AllowUserToDeleteRows = False
        DataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView4.MultiSelect = False
        DataGridView4.RowHeadersWidth = 24
        DataGridView4.ReadOnly = True
        DataGridView4.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView4.Columns(0).Width = (DataGridView4.Width - 26) * 0.12
        DataGridView4.Columns(1).Width = (DataGridView4.Width - 26) * 0.43
        DataGridView4.Columns(2).Width = (DataGridView4.Width - 26) * 0.15
        DataGridView4.Columns(3).Width = (DataGridView4.Width - 26) * 0.15
        DataGridView4.Columns(4).Width = (DataGridView4.Width - 26) * 0.15
        DataGridView4.Refresh()
        DataGridView4.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()

        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
    End Sub
    '*********************************************************************************************************************************************
    '工况序号及名称列表DataGridView1设置并填充函数fill_gk_grid1()，原来用DataGrid8和Adodc13
    '*********************************************************************************************************************************************
    Private Sub fill_gk_grid1()
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 工况序号,工况名称 from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_xhmc_Table.Clear()
        ad.Fill(gk_xhmc_Table)
        ad.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
        BindingSource1.DataSource = gk_xhmc_Table
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
        DataGridView1.Columns(0).Width = 55
        DataGridView1.Columns(1).Width = 135
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    'Timer1的Tick事件处理：（1）画井身结构图；（2）关闭计时器
    '*********************************************************************************************************************************************
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '绘制井身结构图,不含管柱
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
    '点击"退出[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击"<-选用[&C]]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command19_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command19.Click
        Text61.Text = ListBox1.Text
    End Sub
    '*********************************************************************************************************************************************
    '双击待选工况名称列表中选项事件，赋值
    '*********************************************************************************************************************************************
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        Text61.Text = ListBox1.Text
    End Sub
    '*********************************************************************************************************************************************
    '点击"帮助[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中工况序号、名称列表中的某行。原来用DataGrid8和Adodc13
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Call myrefresh()
        Call gk_fgdwshow(CShort(Text53.Text))
        Call gk_mdyjshow(CShort(Text53.Text))
        Call gk_kgyjshow(CShort(Text53.Text))
        Call gk_fgdwflash()
        Call mdgj_flash()
        Call kggj_flash()
    End Sub
    '*********************************************************************************************************************************************
    '填写所选工况的界面中各文本框内容
    '*********************************************************************************************************************************************
    Private Sub myrefresh()
        On Error GoTo errhandler
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
            Text53.Text = IIf(BindingSource1.Current("工况序号").ToString = "", "0", BindingSource1.Current("工况序号").ToString)
            Text61.Text = BindingSource1.Current("工况名称").ToString
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            EXECOleDbCommand.Dispose()
            If RECreader.Read Then
                Text54.Text = RECreader.Item("井口温度℃").ToString
                Text7.Text = RECreader.Item("井底温度℃").ToString
                Text9.Text = RECreader.Item("管压井底深度m").ToString
                Text10.Text = RECreader.Item("井底套压MPa").ToString
                Text11.Text = RECreader.Item("井底管压MPa").ToString
                Text56.Text = RECreader.Item("井口环压MPa").ToString
                Text57.Text = RECreader.Item("环液深度m").ToString
                Text58.Text = RECreader.Item("环液密度g╱cm3").ToString
                Text60.Text = RECreader.Item("环液粘度mPaS").ToString
                Text59.Text = RECreader.Item("环流流量m3╱m").ToString
                Text64.Text = RECreader.Item("井口管压MPa").ToString
                Text65.Text = RECreader.Item("管液深度m").ToString
                Text66.Text = RECreader.Item("管液密度g╱cm3").ToString
                Text68.Text = RECreader.Item("管流粘度mPaS").ToString
                Text67.Text = RECreader.Item("管流流量m3╱m").ToString
                Text63.Text = RECreader.Item("库摩系数").ToString
                Combo2.Text = RECreader.Item("环空流体流向").ToString
                Combo3.Text = RECreader.Item("管内流体流向").ToString
                Combo5.Text = RECreader.Item("井底套压开关").ToString
                Combo6.Text = RECreader.Item("井底管压开关").ToString
                Combo9.Text = RECreader.Item("井口套压开关").ToString
                Combo8.Text = RECreader.Item("井口管压开关").ToString
                Text1.Text = RECreader.Item("套压井底深度m").ToString
                ComboBox1.Text = RECreader.Item("环空流阻模型").ToString
                ComboBox2.Text = RECreader.Item("管内流阻模型").ToString
                Text2.Text = RECreader.Item("环空稠剂浓度").ToString
                Text3.Text = RECreader.Item("环空撑剂浓度").ToString
                Text13.Text = RECreader.Item("环空流变指数n").ToString
                Text14.Text = RECreader.Item("环空稠度系数K").ToString
                Text6.Text = RECreader.Item("管内稠剂浓度").ToString
                Text12.Text = RECreader.Item("管内撑剂浓度").ToString
                Text15.Text = RECreader.Item("管内流变指数n").ToString
                Text16.Text = RECreader.Item("管内稠度系数K").ToString
                Text19.Text = RECreader.Item("环牛模折减系数").ToString
                Text20.Text = RECreader.Item("管牛模折减系数").ToString
                Text21.Text = IIf(RECreader.Item("管流时间h").ToString = "", "0.0", RECreader.Item("管流时间h").ToString)
                Text8.Text = IIf(RECreader.Item("井口加力kN").ToString = "", "0.0", RECreader.Item("井口加力kN").ToString)
                Text18.Text = IIf(RECreader.Item("井口加扭Nm").ToString = "", "0.0", RECreader.Item("井口加扭Nm").ToString)
            End If
            RECreader.Close()
            cn_userdb.Close()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "读写工况数据时出错，可能是因工况数据结构升级所致，请手动删除工况数据或修改数据结构！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '填写所选工况的封隔器状况
    '*********************************************************************************************************************************************
    Private Sub gk_fgdwshow(ByVal xuhao As Short)
        Dim foundRows() As DataRow
        Dim rowfilter As String
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 元件序号,元件名称,封隔器状态,定位方式 from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 工况序号=" & CStr(xuhao) & "  order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            rowfilter = "[元件序号] = " & RECreader.Item("元件序号").ToString
            foundRows = dst.Tables("gk_fgq_Table").Select(rowfilter)
            If foundRows.Length <> 0 Then
                foundRows(0).Item("封隔器状态") = RECreader.Item("封隔器状态").ToString
                foundRows(0).Item("定位方式") = RECreader.Item("定位方式").ToString
            End If
        End While
        RECreader.Close()
        cn_userdb.Close()
        DataGridView2.ResetBindings()
        DataGridView2.Refresh()
        DataGridView2.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView2，在单元格的任何部分被单击时事件-选中封隔器状态表中的某行。原来用DataGrid1和Adodc14
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        Call gk_fgdwflash()
    End Sub
    Private Sub gk_fgdwflash()
        If Not IsNothing(Me.BindingSource2.Current) Then
            Combo4.Text = BindingSource2.Current("定位方式").ToString
            If BindingSource2.Current("封隔器状态").ToString = "坐封" Then
                Option1.Checked = True
                Option2.Checked = False
            Else
                Option1.Checked = False
                Option2.Checked = True
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '坐封情况选择单选钮选择， Option1值发生变化
    '*********************************************************************************************************************************************
    Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
        If eventSender.Checked Then
            Call opt12_chg()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '坐封情况选择单选钮选择，Option2值发生变化
    '*********************************************************************************************************************************************
    Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
        If eventSender.Checked Then
            Call opt12_chg()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '坐封情况选择单选钮选择，处理Option1、Option2发生变化
    '*********************************************************************************************************************************************
    Private Sub opt12_chg()
        If Option1.Checked Then
            If Not IsNothing(Me.BindingSource2.Current) Then
                BindingSource2.Current("封隔器状态") = "坐封"
                DataGridView2.Refresh()
            End If
        Else
            If Not IsNothing(Me.BindingSource2.Current) Then
                BindingSource2.Current("封隔器状态") = "未坐封"
                DataGridView2.Refresh()
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '封隔器定位方式选择列表Combo4选择处理
    '*********************************************************************************************************************************************
    Private Sub Combo4_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo4.SelectedIndexChanged
        If Not IsNothing(Me.BindingSource2.Current) Then
            BindingSource2.Current("定位方式") = Combo4.Text
            DataGridView2.Refresh()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '填写所选工况的锚定工具状况
    '*********************************************************************************************************************************************
    Private Sub gk_mdyjshow(ByVal xuhao As Short)
        Dim foundRows() As DataRow
        Dim rowfilter As String
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 元件序号,元件名称,定位方式 from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 工况序号=" & CStr(xuhao) & "  order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            rowfilter = "[元件序号] = " & RECreader.Item("元件序号").ToString
            foundRows = dst.Tables("gk_mdgj_Table").Select(rowfilter)
            If foundRows.Length <> 0 Then
                foundRows(0).Item("定位方式") = RECreader.Item("定位方式").ToString
            End If
        End While
        RECreader.Close()
        cn_userdb.Close()
        DataGridView3.ResetBindings()
        DataGridView3.Refresh()
        DataGridView3.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView3，在单元格的任何部分被单击时事件-选中锚定工具状态表中的某行。原来用DataGrid2和Adodc2
    '*********************************************************************************************************************************************
    Private Sub DataGridView3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.Click
        Call mdgj_flash()
    End Sub
    Private Sub mdgj_flash()
        If Not IsNothing(Me.BindingSource3.Current) Then
            Combo1.Text = BindingSource3.Current("定位方式").ToString
        End If
    End Sub

    '*********************************************************************************************************************************************
    '锚定工具定位方式选择列表Combo1选择处理
    '*********************************************************************************************************************************************
    Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
        If Not IsNothing(Me.BindingSource3.Current) Then
            BindingSource3.Current("定位方式") = Combo1.Text
            DataGridView3.Refresh()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '填写所选工况的开关工具工作状况
    '*********************************************************************************************************************************************
    Private Sub gk_kgyjshow(ByVal xuhao As Short)
        Dim foundRows() As DataRow
        Dim rowfilter As String
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 工况_开关元件 where 井号='" & well_name _
            & "' and 作业名称='" & zuoye_name & "' and 工况序号=" & CStr(xuhao) & "  order by 元件序号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            rowfilter = "[元件序号] = " & RECreader.Item("元件序号").ToString
            foundRows = dst.Tables("gk_kggj_Table").Select(rowfilter)
            If foundRows.Length <> 0 Then
                foundRows(0).Item("开关状态") = RECreader.Item("开关状态").ToString
                foundRows(0).Item("管内嘴损压差MPa") = RECreader.Item("管内嘴损压差MPa")
                foundRows(0).Item("油套嘴损压差MPa") = RECreader.Item("油套嘴损压差MPa")
            End If
        End While
        RECreader.Close()
        cn_userdb.Close()
        DataGridView4.ResetBindings()
        DataGridView4.Refresh()
        DataGridView4.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView4，在单元格的任何部分被单击时事件-选中开关工具状态表中的某行。原来用DataGrid10和Adodc15
    '*********************************************************************************************************************************************
    Private Sub DataGridView4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView4.Click
        Call kggj_flash()
    End Sub
    Private Sub kggj_flash()
        If Not IsNothing(Me.BindingSource4.Current) Then
            Combo7.Text = BindingSource4.Current("开关状态").ToString
            Text22.Text = BindingSource4.Current("管内嘴损压差MPa").ToString
            Text23.Text = BindingSource4.Current("油套嘴损压差MPa").ToString
        End If
    End Sub
    '*********************************************************************************************************************************************
    '开关工具状态选择列表Combo7选择处理
    '*********************************************************************************************************************************************
    Private Sub Combo7_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo7.SelectedIndexChanged
        If Not IsNothing(Me.BindingSource4.Current) Then
            BindingSource4.Current("开关状态") = Combo7.Text
            DataGridView4.Refresh()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '开关工具状态“"管内嘴损压差MPa”值改变时填写列表
    '*********************************************************************************************************************************************
    Private Sub Text22_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text22.Validated
        If Not IsNothing(Me.BindingSource4.Current) Then
            BindingSource4.Current("管内嘴损压差MPa") = Val(Text22.Text)
            DataGridView4.Refresh()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '开关工具状态“"油套嘴损压差MPa”值改变时填写列表
    '*********************************************************************************************************************************************
    Private Sub Text23_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text23.Validated
        If Not IsNothing(Me.BindingSource4.Current) Then
            BindingSource4.Current("油套嘴损压差MPa") = Val(Text23.Text)
            DataGridView4.Refresh()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '环空流体摩阻计算模型选择
    '*********************************************************************************************************************************************
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call opt3_chg()
    End Sub
    '*********************************************************************************************************************************************
    '环空流体摩阻计算模型选择处理
    '*********************************************************************************************************************************************
    Private Sub opt3_chg()
        Select Case ComboBox1.Text
            Case "无"
                Text59.Enabled = True
                Text60.Enabled = True
                Text2.Enabled = False
                Text3.Enabled = False
                Text13.Enabled = False
                Text14.Enabled = False
                Text19.Enabled = False
            Case "牛顿流体模型"
                Text59.Enabled = True
                Text60.Enabled = True
                Text2.Enabled = False
                Text3.Enabled = True
                Text13.Enabled = False
                Text19.Enabled = True
            Case "降阻比模型"
                Text59.Enabled = True
                Text60.Enabled = True
                Text2.Enabled = True
                Text3.Enabled = True
                Text13.Enabled = False
                Text14.Enabled = False
                Text19.Enabled = False
            Case "幂律流体模型"
                Text59.Enabled = True
                Text60.Enabled = True
                Text2.Enabled = False
                Text3.Enabled = False
                Text13.Enabled = True
                Text14.Enabled = True
                Text19.Enabled = False
        End Select
    End Sub
    '*********************************************************************************************************************************************
    '管内流体摩阻计算模型选择
    '*********************************************************************************************************************************************
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call opt4_chg()
    End Sub

    '*********************************************************************************************************************************************
    '管内流体摩阻计算模型选择处理
    '*********************************************************************************************************************************************
    Private Sub opt4_chg()
        Select Case ComboBox2.Text
            Case "无"
                Text67.Enabled = True
                Text68.Enabled = True
                Text6.Enabled = False
                Text12.Enabled = False
                Text15.Enabled = False
                Text16.Enabled = False
                Text20.Enabled = False
            Case "牛顿流体模型"
                Text67.Enabled = True
                Text68.Enabled = True
                Text6.Enabled = False
                Text12.Enabled = True
                Text15.Enabled = False
                Text16.Enabled = False
                Text20.Enabled = True
            Case "降阻比模型"
                Text67.Enabled = True
                Text68.Enabled = True
                Text6.Enabled = True
                Text12.Enabled = True
                Text15.Enabled = False
                Text16.Enabled = False
                Text20.Enabled = False
            Case "幂律流体模型"
                Text67.Enabled = True
                Text68.Enabled = True
                Text6.Enabled = False
                Text12.Enabled = False
                Text15.Enabled = True
                Text16.Enabled = True
                Text20.Enabled = False
        End Select
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除工况[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选序号的工况参数数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn_userdb.Close()
                Exit Sub
            End If
            SQL_command = "DELETE * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            '顺便删除该作业计算数据
            SQL_command = "DELETE * from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            RECreader.Close()
            cn_userdb.Close()
            Call fill_gk_grid1()
            msg_prompt = "工况参数数据删除完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            RECreader.Close()
            cn_userdb.Close()
            msg_prompt = "未找到工况序号对应的数据，请选好要删除的数据！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "删除工况数据时异常出错，请联系软件编者解决！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存工况[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command21_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command21.Click
        Dim tmp_i As Integer
        Dim tmp_j As Integer
        Dim tmp_k As Integer
        Dim tmp_l As Integer '临时整形变量
        Dim tmp_m As Integer '临时整形变量
        Dim tmp_n As Integer '临时整形变量
        Dim SQL_command As String
        Dim mode_hklt As String
        Dim mode_gnlt As String
        Dim mode_jdysh As String
        Dim row As DataRow
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        mode_hklt = ComboBox1.Text
        mode_gnlt = ComboBox2.Text
        If Text61.Text = "" Then
            msg_prompt = "请选择或输入工况名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text53.Text) = 0 Then
            msg_prompt = "请输入工况序号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo6.Text = "输入" And Val(Text9.Text) = 0 Then
            msg_prompt = "请输入井底管内压力所对应的井底深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo5.Text = "输入" And Val(Text1.Text) = 0 Then
            msg_prompt = "请输入井底环空压力所对应的井底深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text63.Text) <= 0 Then
            msg_prompt = "请输入正确的库伦摩擦系数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo5.Text = "计算" And Combo9.Text = "计算" Then
            msg_prompt = "井口、井底套压总得输入一个。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo5.Text = "输入" And Val(Text10.Text) = 0 Then
            msg_prompt = "请输入正确的井底环空压力。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo5.Text = "计算" And (Val(Text58.Text) = 0 Or Val(Text60.Text) = 0) Then
            msg_prompt = "为计算井底环空压力，请输入正确的环空流体密度和粘度参数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo6.Text = "计算" And Combo8.Text = "计算" Then
            msg_prompt = "井口、井底管内压力总得输入一个。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo6.Text = "输入" And Val(Text11.Text) = 0 Then
            msg_prompt = "请输入正确的井底管内压力。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo6.Text = "计算" And (Val(Text66.Text) = 0 Or Val(Text68.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体参数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text = "不流动" And (ComboBox1.Text <> "无") Then
            msg_prompt = "不流动的环空流体不用计算流动摩阻，请正确选择流体摩阻计算模型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            ComboBox1.Text = "无"
            Exit Sub
        End If
        If Combo2.Text <> "不流动" And (Val(Text59.Text) = 0) Then
            msg_prompt = "环空流体流动，流量不应为0，请输入正确的环空流体流量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text <> "不流动" And mode_hklt = "牛顿流体模型" And (Val(Text60.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体沾度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text <> "不流动" And mode_hklt = "降阻比模型" And (Val(Text2.Text) = 0) Then
            msg_prompt = "请输入正确的环空稠化剂浓度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text <> "不流动" And mode_hklt = "幂律流体模型" And (Val(Text13.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体流变指数n。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text <> "不流动" And mode_hklt = "幂律流体模型" And (Val(Text14.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体稠度系数K。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text = "不流动" And (ComboBox2.Text <> "无") Then
            msg_prompt = "不流动的管内流体不用计算流动摩阻。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            ComboBox2.Text = "无"
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And (Val(Text67.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体流量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And mode_gnlt = "牛顿流体模型" And (Val(Text68.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体沾度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And mode_gnlt = "降阻比模型" And (Val(Text6.Text) = 0) Then
            msg_prompt = "请输入正确的管内稠化剂浓度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And mode_gnlt = "幂律流体模型" And (Val(Text15.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体流变指数n。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And mode_gnlt = "幂律流体模型" And (Val(Text16.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体稠度系数K。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If gk_kggj_Table.Rows.Count > 0 Then
            For Each row In gk_kggj_Table.Rows
                If row.Item("开关状态").ToString = "" Then
                    msg_prompt = "请设定" & row.Item("元件序号").ToString & "号" & row.Item("元件名称").ToString & "的开关状态。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Next
        End If
        If gk_fgq_Table.Rows.Count > 0 Then
            For Each row In gk_fgq_Table.Rows
                If row.Item("封隔器状态").ToString = "未坐封" And row.Item("定位方式").ToString <> "未(无)锚定" Then
                    msg_prompt = row.Item("元件序号").ToString & "号" & row.Item("元件名称").ToString & "封隔器未坐封，定位方式应该是：未(无)锚定。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Next
            For Each row In gk_fgq_Table.Rows
                If row.Item("封隔器状态").ToString = "坐封" And row.Item("定位方式").ToString = "未(无)锚定" Then
                    msg_prompt = row.Item("元件序号").ToString & "号" & row.Item("元件名称").ToString & "封隔器设为坐封，定位方式不应该是：未(无)锚定。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Next
        End If
        '******************************************************************************************************************************************************
        '            查所有的封隔器是否坐卡（锚定）。对于多封隔器，包括水力扩张式的，应选“双向固定”。将来利用轴力差判断是否滑（蠕）动。
        '
        'tmp_k---当前工况封隔器的个数，                         在“工况_封隔定位元件”表中找当前工况，若找到大于0，若没找到等于0。--反映有封隔器
        'tmp_l---当前工况"未坐封"的封隔器个数                   反映是否全部坐封。
        'tmp_n---当前工况定位方式为“双向固定”的封隔器个数封隔器个数 
        '     没有封隔器时：tmp_k = 0  and tmp_l = 0
        '       全部坐封时：tmp_k > 0 and tmp_l = 0 
        '       部分坐封时：tmp_l > 0 And tmp_l < tmp_k
        '     全部未坐封时：tmp_l > 0 And tmp_l = tmp_k
        '******************************************************************************************************************************************************
        tmp_k = gk_fgq_Table.Rows.Count
        tmp_l = 0
        tmp_n = 0
        If tmp_k > 0 Then
            For Each row In gk_fgq_Table.Rows
                If row.Item("定位方式").ToString = "未(无)锚定" Or row.Item("封隔器状态").ToString = "未坐封" Then
                    tmp_l = tmp_l + 1
                End If
                If row.Item("定位方式").ToString = "双向固定" Then
                    tmp_n = tmp_n + 1
                End If
            Next
        End If
        '******************************************************************************************************************************************************
        '查所有的锚定工具是否坐卡
        '
        'tmp_i---当前工况锚定工具的个数，                        tmp_i=0，在“工况_锚定元件”表中找当前工况锚定工具的个数，若找到大于0，若没找到等于0。--反映有锚定工具  
        'tmp_j---当前工况"未(无)锚定"的锚定工具个数              反映是否全部锚定。
        'tmp_m---当前工况定位方式为“双向固定”的锚定工具个数
        '     没有锚定工具时：tmp_i = 0 and tmp_j = 0
        '         全部锚定时：tmp_i > 0 and tmp_j = 0  
        '         部分锚定时：tmp_j > 0 And tmp_j < tmp_i
        '       全部未锚定时：tmp_j > 0 And tmp_j = tmp_i
        '******************************************************************************************************************************************************
        tmp_i = gk_mdgj_Table.Rows.Count
        tmp_j = 0
        tmp_m = 0
        If tmp_i > 0 Then
            For Each row In gk_mdgj_Table.Rows
                If row.Item("定位方式").ToString = "未(无)锚定" Then
                    tmp_j = tmp_j + 1
                End If
                If row.Item("定位方式").ToString = "双向固定" Then
                    tmp_m = tmp_m + 1
                End If
            Next
        End If
        '******************************************************************************************************************************************************
        '                             坐卡（锚定）工况判断,顺便发现坐封(座卡)工况的前面工况锚定工具及封隔器设置不合理的情况

        '1.可计算的“正确”坐封工况。需给变量zf_gkxh赋值，记录坐封工况序号。
        '   可算的坐封(座卡)情况1：(1)有锚定工具，全部锚定；(2)有封隔器，全部坐封锚定。                tmp_i > 0 and tmp_j = 0 and tmp_k > 0 and tmp_l = 0
        '   可算的坐封(座卡)情况2：(1)无锚定工具；(2)有封隔器，全部坐封锚定。                          tmp_i = 0 and tmp_j = 0 and tmp_k > 0 and tmp_l = 0
        '   可算的坐封(座卡)情况3：(1)有锚定工具，全部锚定；(2)无封隔器。                              tmp_i > 0 and tmp_j = 0 and tmp_k = 0  and tmp_l = 0
        '                     
        '
        '2.矛盾的、让计算混乱的“不正确”坐封工况，需要终止程序，提醒用户。
        '   矛盾的、让计算混乱的情况1：（1）有锚定工具，部分锚定；（2）有无封隔器无所谓。              tmp_j > 0 And tmp_j < tmp_i
        '   矛盾的、让计算混乱的情况2：（1）有无锚定工具无所谓；（2）有封隔器，部分坐封锚定。          tmp_l > 0 And tmp_l < tmp_k
        '   矛盾的、让计算混乱的情况3：（1）有锚定工具，全部未锚定；(2)有封隔器，全部坐封锚定。        tmp_j > 0 And tmp_j = tmp_i and tmp_k > 0 and tmp_l = 0
        '   矛盾的、让计算混乱的情况4：（1）有锚定工具，全部锚定；(2)有封隔器，全部未坐封锚定。        tmp_i > 0 and tmp_j = 0 and tmp_l > 0 And tmp_l = tmp_k
        '
        '3.可计算的“未”坐封(座卡)工况。变量zf_gkxh=0，正常计算，不提示。
        '   可计算的“未”坐封(座卡)工况情况1：(1)无锚定工具；(2)无封隔器。
        '   可计算的“未”坐封(座卡)工况情况2：(1)无锚定工具；(2)有封隔器，全部未坐封锚定。
        '   可计算的“未”坐封(座卡)工况情况3：(1)有锚定工具，全部未锚定；(2)无封隔器。
        '
        '******************************************************************************************************************************************************
        If (tmp_j > 0 And tmp_j < tmp_i) Then
            msg_prompt = "锚定工具定位方式设置不合理，锚定工具要么全锚定，要么全未(无)锚定。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (tmp_l > 0 And tmp_l < tmp_k) Then
            msg_prompt = "封隔器坐封及定位方式设置不合理，封隔器要么全坐封全锚定，要么全未坐封全未(无)锚定）。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (tmp_j > 0 And tmp_j = tmp_i And tmp_k > 0 And tmp_l = 0) Then
            msg_prompt = "封隔器已坐封锚定，但锚定工具未全锚定，如此设置不合理（锚定工具应全锚定，封隔器应全坐封全锚定）。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (tmp_i > 0 And tmp_j = 0 And tmp_l > 0 And tmp_l = tmp_k) Then
            msg_prompt = "锚定工具已锚定，但封隔器未全坐封锚定，如此设置不合理（锚定工具应全锚定，封隔器应全坐封全锚定）。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '******************************************************************************************************************************************************
        '                                          有封隔定位工况，封隔器（锚定工具）定位方式的进一步检验
        '                                                                                                                 2022年2月25日秦彦斌最后修订整理
        '     由于每一封隔器（锚定工具）理论上有四种定位方式，对于多封隔器（锚定工具）管柱，多个封隔器（锚定工具）的定位就会有多种组合。每种组合变形协调计算
        ' 的区间会因定位方式、管柱伸长与缩短而不同，穷举各种迭代情况并算法实现非常困难。
        '     为简化算法并程序实现，对于单封隔器（锚定工具）管柱，允许那一个封隔器（锚定工具）可有四种定位方式；对于多封隔器（锚定工具）管柱，所有的定位方式
        ' 只能是“双向固定”。所以对于多封隔器，包括水力扩张式的，应选“双向固定”。将来利用轴力差判断是否滑（蠕）动，利用变形效应判断滑（蠕）动距离。
        '     在工况输入保存及管柱力学分析计算前进行此检验
        '
        '     封隔器（锚定工具）定位方式的进一步检验算法：
        '       统计前工况锚定工具的个数tmp_i，当前工况定位方式为“双向固定”的锚定工具个数tmp_m
        '       统计前工况封隔器的个数tmp_k，当前工况定位方式为“双向固定”的封隔器个数tmp_n
        '       if （当前工况锚定工具的个数tmp_i+当前工况封隔器的个数tmp_k）>1，表示该管柱是多封隔器（锚定工具）管柱
        '           if 定位方式为“双向固定”的锚定工具个数tmp_m<>当前工况锚定工具的个数tmp_i
        '              给提示，退出
        '           endif
        '           if  定位方式为“双向固定”的封隔器个数tmp_n<>当前工况锚定工具的个数tmp_k
        '              给提示，退出
        '           endif
        '      endif
        '******************************************************************************************************************************************************
        If (tmp_i + tmp_k) > 1 Then
            If tmp_j = 0 And (tmp_i <> tmp_m) Then
                msg_prompt = "多封隔器（锚定工具）管柱的锚定工具定位方式应全设置为【双向固定】。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
            If tmp_l = 0 And (tmp_k <> tmp_n) Then
                msg_prompt = "多封隔器（锚定工具）管柱的封隔器定位方式应全设置为【双向固定】。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        basedb_chg = False
        SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.Read Then
            msg_prompt = "是否要建立新的工况数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                Exit Sub
            End If
            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            '            SQL_command = "CREATE TABLE 工况参数表 " _
            ''            & "(作业名称     TEXT(50),工况序号      INTEGER,工况名称    TEXT(50),环空流体流向 TEXT(10),管内流体流向 TEXT(10), " _
            ''            & " 井口温度℃      FLOAT,井底温度℃      FLOAT,管压井底深度m  FLOAT,井口环压MPa     FLOAT,环液深度m       FLOAT, " _
            ''            & " 环液密度g╱cm3  FLOAT,环液粘度mPaS    FLOAT,环流流量m3╱m  FLOAT,井口管压MPa     FLOAT,管液深度m       FLOAT, " _
            ''            & " 管液密度g╱cm3  FLOAT,管流粘度mPaS    FLOAT,管流流量m3╱m  FLOAT,库摩系数        FLOAT,井口加力kN      FLOAT, " _
            ''            & " 井口加扭Nm      FLOAT,井底套压MPa     FLOAT,井底管压MPa    FLOAT,井底套压开关 TEXT(20),井底管压开关 TEXT(20), " _
            ''            & " 井口套压开关 TEXT(20),井口管压开关 TEXT(20),套压井底深度m  FLOAT,环空流阻模型 TEXT(20),管内流阻模型 TEXT(20), " _
            ''            & " 环空稠剂浓度    FLOAT,管内稠剂浓度    FLOAT,环空撑剂浓度   FLOAT,管内撑剂浓度    FLOAT,环空流变指数n   FLOAT, " _
            ''            & " 管内流变指数n   FLOAT,环空稠度系数K   FLOAT,管内稠度系数K  FLOAT,井底约束情况 TEXT(20),井底约束位置m   FLOAT, " _
            ''            & " 约束定位方式 TEXT(20),管牛模折减系数  FLOAT,环牛模折减系数 FLOAT,管流时间h       FLOAT,井号          TEXT(50))"
            SQL_command = "insert into  工况参数表 (作业名称,工况序号,工况名称,井口温度℃,井口环压MPa,环液深度m,环液密度g╱cm3,环液粘度mPaS,环流流量m3╱m,井口管压MPa,管液深度m," _
                & "管液密度g╱cm3,管流粘度mPaS,管流流量m3╱m,库摩系数,井口加力kN,井口加扭Nm,井号,环空流体流向,管内流体流向,井底温度℃," _
                & "管压井底深度m,井底套压MPa,井底管压MPa,井底套压开关,井底管压开关,井口套压开关,井口管压开关,套压井底深度m,环空流阻模型,管内流阻模型," _
                & "环空稠剂浓度,管内稠剂浓度,环空撑剂浓度,管内撑剂浓度,环空流变指数n,管内流变指数n,环空稠度系数K,管内稠度系数K," _
                & "管牛模折减系数,环牛模折减系数,管流时间h) values (" _
                & "'" & zuoye_name & "'," & CStr(Text53.Text) & ",'" & Text61.Text & "'," & CStr(Text54.Text) & "," & CStr(Text56.Text) & "," _
                & CStr(Text57.Text) & "," & CStr(Text58.Text) & "," & CStr(Text60.Text) & "," & CStr(Text59.Text) & "," & CStr(Text64.Text) & "," _
                & CStr(Text65.Text) & "," & CStr(Text66.Text) & "," & CStr(Text68.Text) & "," & CStr(Text67.Text) & "," & CStr(Text63.Text) & "," _
                & CStr(Text8.Text) & "," & CStr(Text18.Text) & ",'" & well_name & "','" & Combo2.Text & "','" & Combo3.Text & "'," _
                & CStr(Text7.Text) & "," & CStr(Text9.Text) & "," & CStr(Text10.Text) & "," & CStr(Text11.Text) & ",'" & Combo5.Text & "','" _
                & Combo6.Text & "','" & Combo9.Text & "','" & Combo8.Text & "'," & CStr(Text1.Text) & ",'" & mode_hklt & "','" _
                & mode_gnlt & "'," & CStr(Text2.Text) & "," & CStr(Text6.Text) & "," & CStr(Text3.Text) & "," & CStr(Text12.Text) & "," _
                & CStr(Text13.Text) & "," & CStr(Text15.Text) & "," & CStr(Text14.Text) & "," & CStr(Text16.Text) & "," & CStr(Text20.Text) & "," _
                & CStr(Text19.Text) & "," & CStr(Text21.Text) & ")"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            '顺便删除该作业计算数据
            SQL_command = "DELETE * from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "select * from 工况名称 where 工况名称='" & Text61.Text & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            EXECOleDbCommand.Dispose()
            If Not RECreader.Read Then
                msg_prompt = "是否将此工况名称加入到工况名称库中？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return = 6 Then
                    ListBox1.Items.Add(Text61.Text)
                    SQL_command = "insert into  工况名称 (工况名称) values ('" & Text61.Text & "')"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                End If
            End If
        Else
            msg_prompt = "是否要保存对所选工况数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                Exit Sub
            End If
            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "delete * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            SQL_command = "update 工况参数表 set " & "工况名称='" & Text61.Text & "',井口温度℃=" & CStr(Text54.Text) & "," & "井底温度℃=" & CStr(Text7.Text) _
                & ",管压井底深度m=" & CStr(Text9.Text) & "," & "井口环压MPa=" & CStr(Text56.Text) & ",环液深度m=" & CStr(Text57.Text) & "," & "环液密度g╱cm3=" & CStr(Text58.Text) _
                & ",环液粘度mPaS=" & CStr(Text60.Text) & "," & "环流流量m3╱m=" & CStr(Text59.Text) & ",井口管压MPa=" & CStr(Text64.Text) & "," & "管液深度m=" & CStr(Text65.Text) _
                & ",管液密度g╱cm3=" & CStr(Text66.Text) & "," & "管流粘度mPaS=" & CStr(Text68.Text) & ",管流流量m3╱m=" & CStr(Text67.Text) & "," & "库摩系数=" & CStr(Text63.Text) _
                & ",井口加力kN=" & CStr(Text8.Text) & "," & "环空流体流向='" & Combo2.Text & "',管内流体流向='" & Combo3.Text & "'," & "井口加扭Nm=" & CStr(Text18.Text) _
                & ",井底套压MPa=" & CStr(Text10.Text) & "," & "井底管压MPa=" & CStr(Text11.Text) & ",井底套压开关='" & Combo5.Text & "'," & "井底管压开关='" & Combo6.Text _
                & "',井口套压开关='" & Combo9.Text & "'," & "井口管压开关='" & Combo8.Text & "',套压井底深度m=" & CStr(Text1.Text) & "," & "环空流阻模型='" & mode_hklt _
                & "',管内流阻模型='" & mode_gnlt & "'," & "环空稠剂浓度=" & CStr(Text2.Text) & ",管内稠剂浓度=" & CStr(Text6.Text) & "," & "环空撑剂浓度=" & CStr(Text3.Text) _
                & ",管内撑剂浓度=" & CStr(Text12.Text) & "," & "环空流变指数n=" & CStr(Text13.Text) & ",管内流变指数n=" & CStr(Text15.Text) & "," & "环空稠度系数K=" & CStr(Text14.Text) _
                & ",管内稠度系数K=" & CStr(Text16.Text) & "," & "管牛模折减系数=" & CStr(Text20.Text) & ",环牛模折减系数=" & CStr(Text19.Text) & "," & "管流时间h=" & CStr(Text21.Text) _
                & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        If gk_fgq_Table.Rows.Count > 0 Then
            For Each row In gk_fgq_Table.Rows
                SQL_command = "insert into  工况_封隔定位元件 (作业名称,工况序号,工况名称,元件序号,元件名称,封隔器状态,定位方式,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(Text53.Text) & "," & "'" & Text61.Text & "'," & row.Item("元件序号").ToString & "," & "'" & row.Item("元件名称").ToString & "'," _
                    & "'" & row.Item("封隔器状态").ToString & "'," & "'" & row.Item("定位方式").ToString & "'," & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            Next
        End If
        If gk_mdgj_Table.Rows.Count > 0 Then
            For Each row In gk_mdgj_Table.Rows
                SQL_command = "insert into  工况_锚定元件 (作业名称,工况序号,工况名称,元件序号,元件名称,定位方式,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(Text53.Text) & "," & "'" & Text61.Text & "'," & row.Item("元件序号").ToString & "," _
                    & "'" & row.Item("元件名称").ToString & "'," & "'" & row.Item("定位方式").ToString & "'," & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            Next
        End If
        If gk_kggj_Table.Rows.Count > 0 Then
            For Each row In gk_kggj_Table.Rows
                SQL_command = "insert into  工况_开关元件 (作业名称,工况序号,工况名称,元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(Text53.Text) & "," & "'" & Text61.Text & "'," & row.Item("元件序号").ToString & "," & "'" & row.Item("元件名称").ToString & "'," _
                    & "'" & row.Item("开关状态").ToString & "'," & row.Item("管内嘴损压差MPa").ToString & "," & row.Item("油套嘴损压差MPa").ToString & "," & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            Next
        End If
        cn_basedb.Close()
        cn_userdb.Close()
        Call fill_gk_grid1()
        msg_prompt = "工况参数数据保存完成。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub

    '*********************************************************************************************************************************************
    '界面Closed，为了清理老版本软件中的无用临时表，升级并保留此部分代码。
    '*********************************************************************************************************************************************
    Private Sub ipt_gkuang_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim rowfilter As String

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='lsb_gk_fgq'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE lsb_gk_fgq"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        rowfilter = "TABLE_NAME='lsb_gk_mdgj'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE lsb_gk_mdgj"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        rowfilter = "TABLE_NAME='lsb_gk_kgyj'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE lsb_gk_kgyj"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        cn_userdb.Close()

        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
End Class