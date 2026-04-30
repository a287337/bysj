Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Imports System.Collections.Generic

Friend Class ipt_sk_gk
    '*********************************************************************************************************************************************
    '                                                   关于射孔工况数据输入与管理界面的说明
    '
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private SQL_command As String
    Private gk_fgq_table As New DataTable    '管柱_封隔器表,对应grid2
    Private sk_gk_table As New DataTable     '射孔工况参数表,对应grid1
    Private skq_wz_table As New DataTable    '射孔枪位置表,对应grid5
    Private rgjd As Single     '人工井底m   
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub ipt_sk_gk_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        GroupBox2.Text = well_name & "井" & zuoye_name & "作业工况参数"
        Me.Text = well_name & "井" & zuoye_name & "作业射孔工况参数数据输入与修改"
        '***************************************************************************************************************
        '初始化各变量
        '***************************************************************************************************************
        Textg11.Text = "0"  '工况序号
        Text2.Text = ""     '射孔施工方式
        Textg7.Text = ""   '工况名称
        
        Textg2.Text = CStr(0.0#)
        Textg3.Text = CStr(0.0#)
        Textg4.Text = CStr(0.0#)
        Textg5.Text = CStr(0.0#)
        Textg6.Text = CStr(0.0#)

        Textg8.Text = CStr(3.0#)
        Textg15.Text = CStr(0.0#)
        Textg23.Text = "0.0"

        '***************************************************************************************************************
        '根据机械设计手册：
        '    轻冲击、低频冲击  动载系数  1.0-1.2，   中等惯性、中等冲击力  动载系数  1.2-1.8
        '    强冲击、高频冲击  动载系数  1.8-3.0
        ' 故这里取值为 1.5
        '***************************************************************************************************************
        Textg9.Text = CStr(1.5)
        Text22.Text = CStr(0.0#)  '筛管距离
        '***************************************************************************************************************
        '设计井深m、完钻井深m 只读文本框赋值
        '***************************************************************************************************************
        Text4.Text = ""
        Text5.Text = ""
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位 from 油气井表 where 井号='" & well_name & "'"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text4.Text = RECreader.Item("设计井深m").ToString
                        Text5.Text = RECreader.Item("完钻井深m").ToString
                    End If
                    TextBox1.Text = ""
                End Using
            End Using
            SQL_command = "select sum(元件长度m) as 管柱总长 from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        TextBox1.Text = RECreader.Item("管柱总长").ToString()
                    End If
                End Using
            End Using
            '获得作业地层关联的目的层输入的人工井底，2025年10月23日修改为直接从作业地层表中获取
            SQL_command = "select 目的层名称,人工井底m,目的层温度℃  from 作业地层参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            'SQL_command = "select 作业地层参数表.目的层名称,人工井底m,地层温度℃  from 作业地层参数表,目的层参数表" _
            '& " where 作业地层参数表.目的层名称=目的层参数表.目的层名称 and 作业地层参数表.井号='" & well_name & "' and 作业地层参数表.作业名称='" & zuoye_name & " '"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        rgjd = If(IsDBNull(RECreader.Item("人工井底m")), 0, RECreader.Item("人工井底m"))
                        Textg10.Text = rgjd.ToString()
                        Textg3.Text = If(IsDBNull(RECreader.Item("目的层温度℃")), "0", RECreader.Item("目的层温度℃").ToString())
                    End If
                End Using
            End Using
        End Using

        Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
            cn_basedb.Open()
            SQL_command = "select * from 射孔施工方式表 order by 施工方式"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_basedb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    ListBox1.Items.Clear()
                    While RECreader.Read
                        ListBox1.Items.Add(RECreader.Item("施工方式").ToString)
                    End While
                End Using
            End Using
        End Using
        '*********************************************************************************************************************************************
        '用管柱数据表计算填(1) gz_yj_table，
        '*********************************************************************************************************************************************
        '元件重量kg是整个长度的重量，不是单位长重量
        Call fill_gzyj_table()
        ''*********************************************************************************************************************************************
        '工况序号及名称列表填sk_gk_table， DataGridView1
        '*********************************************************************************************************************************************
        Call fill_gk_grid()
        '*********************************************************************************************************************************************
        '用管柱数据填射孔器位置表skq_wz_table，用这个表填充 DataGridView5，筛管距射孔段顶端表datagridview4
        '*********************************************************************************************************************************************
        Call fill_skq_wz_grid()
        '*********************************************************************************************************************************************
        '用frm_pjskd_l.vb中定义的函数（工况_射孔夹层表填treeview控件）
        '*********************************************************************************************************************************************
        'Call fill_gkjctree()
        Call fill_gkxztree()
        'TreeView1.CheckBoxes = True
        '*********************************************************************************************************************************************
        '用工况_封隔器表填gk_fgq_Table，DataGridView2
        '*********************************************************************************************************************************************
        Call fill_fgq_grid()
        '*********************************************************************************************************************************************
        '填写筛管距离，射孔段长度等文本框
        '*********************************************************************************************************************************************
        Opt1.Checked = True
        Call opt3_chg()

        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
    End Sub
    '*********************************************************************************************************************************************
    '用工况_封隔器表填gk_fgq_Table，DataGridView2
    '若工况_封隔定位元件表有当前作业的所有封隔器信息（坐封或未坐封，针对不同工况），理论上应优先用这个表填写grid，但是会出现封隔器重复的问题
    '所以：（1）用管柱_封隔定位元件中的封隔器数据，先设置为全部为“未坐封”状态
    '      （2）选中某一工况后，根据工况_封隔定位元件表中当前工况的封隔器状态更新gk_fgq_Table的显示内容
    '*********************************************************************************************************************************************
    Private Sub fill_fgq_grid()
        Dim fgq_newRow As DataRow
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            '建立gk_fgq_Table表结构
            SQL_command = "select 元件序号,元件名称,封隔器状态 from 工况_封隔定位元件 where 0>1"
            Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                gk_fgq_Table.Clear()
                ad.Fill(gk_fgq_Table)
            End Using
            '筛出管柱中所有封隔器，暂时都设置为“未坐封”
            SQL_command = "select 元件序号,元件名称 from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    While RECreader.Read
                        fgq_newRow = gk_fgq_Table.NewRow()
                        fgq_newRow.Item("元件序号") = RECreader.Item("元件序号").ToString
                        fgq_newRow.Item("元件名称") = RECreader.Item("元件名称").ToString
                        fgq_newRow.Item("封隔器状态") = "未坐封"
                        gk_fgq_Table.Rows.Add(fgq_newRow)
                    End While
                End Using
            End Using
        End Using
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
        DataGridView2.Columns(0).Width = (DataGridView2.Width - 26) * 0.25
        DataGridView2.Columns(1).Width = (DataGridView2.Width - 26) * 0.5
        DataGridView2.Columns(2).Width = (DataGridView2.Width - 26) * 0.25
        DataGridView2.Refresh()
        DataGridView2.Show()
    End Sub
    '*********************************************************************************************************************************************
    '射孔工况参数表填sk_gk_table，DataGridView1
    '*********************************************************************************************************************************************
    Private Sub fill_gk_grid()
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select 工况序号,工况名称,井口温度℃,射孔段处井底温度℃,井底初始压力MPa,井口加压MPa,射孔液密度g╱cm3,射孔施工方式," _
                        & "筛管距射孔顶端距离m,人工井底m,动载系数,[射孔液粘度mPa·s],射孔段起始深度m,射孔夹层总厚度m,作业名称,井号  from 射孔工况参数表 where 井号='" _
                        & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号"
            Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                sk_gk_table.Clear()
                ad.Fill(sk_gk_table)
            End Using
        End Using
        BindingSource1.DataSource = sk_gk_table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Width = DataGridView1.Width / 13
        Next
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(7).Width = 200
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    
    '*********************************************************************************************************************************************
    '用管柱数据表计算填(1) gz_yj_table，
    '                 （2）sg_jl_table，DataGridView4；
    '                  (3) gk_wz_table，DataGridView3
    '*********************************************************************************************************************************************
    Private Sub fill_skq_wz_grid()
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            '建立skq_wz_table结构
            SQL_command = "select 射孔枪序号,射孔枪名称,起始深度m, 终止深度m,射孔枪状态, 起始深度m as 到筛管的距离m, 起始深度m as 射孔段起始深度m from 工况_射孔夹层表 where 0>1"
            Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                skq_wz_table.Clear()
                ad.Fill(skq_wz_table)
            End Using
        End Using
        Dim filter As String = "元件性质='射孔枪'"
        Dim sort As String = "元件序号 asc"
        Dim skq_rows As DataRow() = gz_yj_table.Select(filter, sort)
        Dim skq_wz_row As DataRow
        For Each row As DataRow In skq_rows
            If row.Item("元件性质").ToString = "射孔枪" Then
                skq_wz_row = skq_wz_table.NewRow()
                skq_wz_row.Item("射孔枪序号") = row.Item("元件序号").ToString()
                skq_wz_row.Item("射孔枪名称") = row.Item("元件名称").ToString()
                skq_wz_row.Item("起始深度m") = row.Item("起始深度m").ToString()
                skq_wz_row.Item("终止深度m") = row.Item("终止深度m").ToString()
                skq_wz_row.Item("射孔枪状态") = "工作"
                skq_wz_row.Item("射孔段起始深度m") = row.Item("起始深度m").ToString()
                '找上方最近的筛管
                filter = "元件性质='筛管' and 元件序号<" & row.Item("元件序号").ToString()
                sort = "元件序号  desc"
                Dim sg_rows() As DataRow = gz_yj_table.Select(filter, sort)
                If sg_rows.Length > 0 Then
                    skq_wz_row.Item("到筛管的距离m") = CSng(row.Item("起始深度m")) - CSng(sg_rows(0).Item("起始深度m"))
                Else
                    skq_wz_row.Item("到筛管的距离m") = 0
                End If
                skq_wz_table.Rows.Add(skq_wz_row)
            End If
        Next

        BindingSource5.DataSource = skq_wz_table
        DataGridView5.ClearSelection()
        DataGridView5.DataSource = BindingSource5
        DataGridView5.ResetBindings()
        DataGridView5.AutoGenerateColumns = True
        DataGridView5.AllowUserToAddRows = False
        DataGridView5.AllowUserToDeleteRows = False
        DataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView5.MultiSelect = False
        DataGridView5.RowHeadersWidth = 24
        DataGridView5.ReadOnly = True
        DataGridView5.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView5.Columns(0).Width = (DataGridView5.Width - 26) * 0.1
        DataGridView5.Columns(1).Width = (DataGridView5.Width - 26) * 0.2
        DataGridView5.Columns(2).Width = (DataGridView5.Width - 26) * 0.2
        DataGridView5.Columns(3).Width = (DataGridView5.Width - 26) * 0.2
        DataGridView5.Columns(4).Width = (DataGridView5.Width - 26) * 0.1
        DataGridView5.Columns(5).Width = (DataGridView5.Width - 26) * 0.2
        DataGridView5.Refresh()
        DataGridView5.Show()

    End Sub
    '*********************************************************************************************************************************************
    '当前作业的所有工况的夹层情况,填jc_hd_table
    '*********************************************************************************************************************************************
    Private Function fill_gkxztree() As Integer
        Dim SQL_command As String
        Dim jc_hd_table As New DataTable
        Dim zhd_table As New DataTable
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                
                SQL_command = "SELECT  工况序号,起始深度m,终止深度m,终止深度m - 起始深度m AS 夹层厚度m FROM 工况_射孔夹层表" _
                                  & " WHERE 井号 = '" & well_name & "' and 作业名称='" & zuoye_name & "' AND 射孔枪状态 = '工作'" _
                                  & " ORDER BY  工况序号 ASC,射孔枪序号 ASC"
                Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                    jc_hd_table.Clear()
                    ad.Fill(jc_hd_table)
                End Using
                SQL_command = "SELECT  工况序号,射孔夹层总厚度m FROM 射孔工况参数表" _
                                  & " WHERE 井号 = '" & well_name & "' and 作业名称='" & zuoye_name & "'" _
                                  & " ORDER BY  工况序号 ASC"
                Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                    zhd_table.Clear()
                    ad.Fill(zhd_table)
                End Using
            End Using
            fill_gkxztree = jc_hd_table.Rows.Count
        Catch ex As OleDbException
            fill_gkxztree = -1
            msg_prompt = "读取工况_射孔夹层数据时出错！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try

        TreeView1.Nodes.Clear()
        Dim rootNode As New TreeNode(zuoye_name & "射孔层")
        '根节点name=root
        rootNode.Name = "root"
        treeview1.Nodes.Add(rootNode)
        '获得唯一的工况序号和射孔总厚度表
        For Each key As DataRow In zhd_table.Rows
            '根节点的第一层子节点，显示“工况k”，name=“工况k”
            Dim gkNode As New TreeNode("工况" & key.Item("工况序号").ToString() & ",射孔总厚度" & key.Item("射孔夹层总厚度m").ToString & "m")
            gkNode.Name = "工况" & key.Item("工况序号").ToString()
            '构建“工况k”节点的子节点gkNode——总厚度；起始-终止m，共m；……。name中都包含“工况k-XXX”
            'Dim zhdNode As New TreeNode("射孔总厚度：" & key.Item("射孔夹层总厚度m").ToString & "m")
            ''gkNode的第一个子节点
            ''Name = "工况k-厚度"
            'zhdNode.Name = "工况" & key.Item("工况序号").ToString() & "-厚度"
            'gkNode.Nodes.Add(zhdNode)
            Dim childNodeText As String = ""
            Dim jcCount As Integer = 0
            'gkNode的其余子节点
            For Each jc As DataRow In jc_hd_table.Rows
                If jc("工况序号").ToString = key.Item("工况序号").ToString Then
                    jcCount = jcCount + 1
                    childNodeText = jc("起始深度m").ToString() & "-" & jc("终止深度m").ToString() & "m，" & jc("夹层厚度m").ToString() & "m"
                    Dim childNode As New TreeNode(childNodeText)
                    '其余子节点的name=“工况k-j”
                    childNode.Name = "工况" & key.Item("工况序号").ToString() & "-" & jcCount.ToString()
                    gkNode.Nodes.Add(childNode)
                End If
            Next
            rootNode.Nodes.Add(gkNode)
        Next
        treeview1.ExpandAll()
        treeview1.Refresh()
    End Function
    '***************************************************************************************************************************
    '树控件设定只有工况k节点能被选中,
    '选中后各界面显示当前工况的内容
    '***************************************************************************************************************************
    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        '找到节点代表的工况序号
        If e.Node Is Nothing Then Exit Sub
        Dim gkxh As Integer = Convert.ToInt32(e.Node.Name.Substring(2))
        '选中sk_gk_table及其grid的gkxh行
        If Not IsNothing(BindingSource1) AndAlso BindingSource1.Count > 0 Then
            Dim position As Integer
            position = BindingSource1.Find("工况序号", gkxh)
            If Not position < 0 Then
                DataGridView1.CurrentCell = DataGridView1.Rows(position).Cells(0)
                DataGridView1.Rows(position).Selected = True
            End If
        End If
        Call set_skgk_texts("工况表")
        '显示当前工况下，gk_fgq_Table中各封隔器的状态
        Call gk_fgdwshow(Textg11.Text)
        '设定封隔器坐封与否的Option
        Call gk_fgdwflash()
        '显示当前工况下，skq_wz_table中各射孔枪的状态
        Call gk_skq_zt(Textg11.Text)
        '设定射孔器工作与否的Option
        Call gk_skq_ztOpt()
    End Sub
    '设定树控件的非工况节点不能被选中
    Private Sub TreeView1_BeforeSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeSelect
        If e.Node.Name.Contains("-") Or e.Node.Name = "root" Then
            e.Cancel = True
        End If
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
    '点击"<-选用[&C]]"--射孔施工方式---按钮
    '*********************************************************************************************************************************************
    Private Sub Command19_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command19.Click
        Text2.Text = ListBox1.Text
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除[&F]]"--射孔施工方式---按钮
    '*********************************************************************************************************************************************
    Private Sub Button3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Button3.Click
        If ListBox1.Items.Count > 0 AndAlso Trim(ListBox1.Text) <> "" Then
            msg_prompt = "是否确认从基础基础数据库表中删除备选施工方式:" & Chr(13) & Chr(10) & ListBox1.Text & "?"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                Exit Sub
            End If
            Try
                Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                    cn_basedb.Open()
                    SQL_command = "delete * from  射孔施工方式表 where 施工方式='" & Trim(ListBox1.Text) & "'"
                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand2.ExecuteNonQuery()
                    End Using
                    SQL_command = "select * from 射孔施工方式表 order by 施工方式"
                    Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_basedb)
                        Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                            ListBox1.Items.Clear()
                            While RECreader.Read
                                ListBox1.Items.Add(RECreader.Item("施工方式").ToString)
                            End While
                        End Using
                    End Using
                End Using
            Catch ex As OleDbException
                msg_prompt = "基础数据库访问异常！"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End Try
        End If
    End Sub
    '*********************************************************************************************************************************************
    '双击待选工况名称列表中选项事件，赋值
    '*********************************************************************************************************************************************
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        Text2.Text = ListBox1.Text
    End Sub
    '*********************************************************************************************************************************************
    '点击"帮助[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中射孔工况列表中的某行。
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Call set_skgk_texts("工况表")
        '显示当前工况下，gk_fgq_Table中各封隔器的状态
        Call gk_fgdwshow(Textg11.Text)
        '设定封隔器坐封与否的Option
        Call gk_fgdwflash()
        '显示当前工况下，skq_wz_table中各射孔枪的状态
        Call gk_skq_zt(Textg11.Text)
        '设定射孔枪工作与否的Option
        Call gk_skq_ztOpt()
    End Sub
    '*********************************************************************************************************************************************
    '显示当前选定工况的封隔器坐封状况
    '*********************************************************************************************************************************************
    Private Sub gk_fgdwshow(ByVal xuhao As String)
        Dim foundRows() As DataRow
        Dim rowfilter As String
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select 元件序号,元件名称,封隔器状态 from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 工况序号=" & CStr(xuhao) & "  order by 元件序号"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    While RECreader.Read
                        rowfilter = "[元件序号] = " & RECreader.Item("元件序号").ToString
                        foundRows = gk_fgq_Table.Select(rowfilter)
                        '修改gk_fgq_Table的封隔器状态
                        If foundRows.Length <> 0 Then
                            foundRows(0).Item("封隔器状态") = RECreader.Item("封隔器状态").ToString
                        End If
                    End While
                End Using
            End Using
        End Using

        DataGridView2.ResetBindings()
        DataGridView2.Refresh()
        DataGridView2.Show()

    End Sub

    '*********************************************************************************************************************************************
    '单击DataGridView2，在单元格的任何部分被单击时事件-选中封隔器状态表中的某行。设定封隔器状态Option
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        Call gk_fgdwflash()
    End Sub
    '*********************************************************************************************************************************************
    '设定封隔器状态Option
    '*********************************************************************************************************************************************
    Private Sub gk_fgdwflash()
        If Not IsNothing(Me.BindingSource2.Current) Then
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
    Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged, Opt1.CheckedChanged
        If eventSender.Checked Then
            Call opt12_chg()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '坐封情况选择单选钮选择，Option2值发生变化
    '*********************************************************************************************************************************************
    Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged, Opt2.CheckedChanged
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
    '显示当前选定工况的射孔枪工作状况
    '即：用当前工况的各射孔枪工作状态修改grid，并显示
    '*********************************************************************************************************************************************
    Private Sub gk_skq_zt(ByVal xuhao As String)
        Dim foundRows() As DataRow
        Dim rowfilter As String
        Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select 射孔枪序号,射孔枪名称,射孔枪状态 from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 工况序号=" & xuhao & "  order by 射孔枪序号"
            Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                    While RECreader.Read
                        rowfilter = "[射孔枪序号] = " & RECreader.Item("射孔枪序号").ToString
                        foundRows = skq_wz_table.Select(rowfilter)
                        '修改skq_wz_table的射孔枪状态
                        If foundRows.Length <> 0 Then
                            foundRows(0).Item("射孔枪状态") = RECreader.Item("射孔枪状态").ToString
                        End If

                    End While
                End Using
            End Using
        End Using

        DataGridView5.ResetBindings()
        DataGridView5.Refresh()
        DataGridView5.Show()

    End Sub

    '*********************************************************************************************************************************************
    '单击DataGridView5，在单元格的任何部分被单击时事件-选中射孔枪状态表中的某行。设定射孔枪状态Option
    '*********************************************************************************************************************************************
    Private Sub DataGridView5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView5.Click
        Call gk_skq_ztOpt()
    End Sub
    '*********************************************************************************************************************************************
    '设定射孔枪状态Option
    '*********************************************************************************************************************************************
    Private Sub gk_skq_ztOpt()
        If Not IsNothing(Me.BindingSource5.Current) Then
            If BindingSource5.Current("射孔枪状态").ToString = "工作" Then
                Opt1.Checked = True
                Opt2.Checked = False
            Else
                Opt1.Checked = False
                Opt2.Checked = True
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '射孔枪工作状态选择单选钮选择， Opt1值发生变化
    '*********************************************************************************************************************************************
    Private Sub Opt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Opt1.CheckedChanged
        If eventSender.Checked Then
            Call opt3_chg()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '射孔枪工作状态选择单选钮选择， Opt2值发生变化
    '*********************************************************************************************************************************************
    Private Sub Opt2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Opt2.CheckedChanged
        If eventSender.Checked Then
            Call opt3_chg()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '射孔枪工作状态选择单选钮选择，处理Opt1、Opt2发生变化
    '用opt1和opt2的值改变grid，设定为待保存工况的射孔枪状态
    '根据射孔枪状态改编设定text22（射孔段到筛管的距离），计算射孔夹层总厚度，填textg15
    '*********************************************************************************************************************************************
    Private Sub opt3_chg()
        Dim find_row() As DataRow
        Dim filter As String
        If Opt1.Checked Then
            If Not IsNothing(Me.BindingSource5.Current) Then
                '直接修改绑定表skq_wz_table的值，避免延迟保存，
                filter = "射孔枪序号=" & BindingSource5.Current("射孔枪序号").ToString
                find_row = skq_wz_table.Select(filter)
                If find_row.Length > 0 Then
                    find_row(0).Item("射孔枪状态") = "工作"
                End If
                DataGridView5.ResetBindings()
                DataGridView5.Refresh()
                DataGridView5.Show()
                'BindingSource5.Current("射孔枪状态") = "工作"
                'DataGridView5.Refresh()
            End If
        Else
            If Not IsNothing(Me.BindingSource5.Current) Then
                filter = "射孔枪序号=" & BindingSource5.Current("射孔枪序号").ToString
                find_row = skq_wz_table.Select(filter)
                If find_row.Length > 0 Then
                    find_row(0).Item("射孔枪状态") = "不工作"
                End If
                DataGridView5.ResetBindings()
                DataGridView5.Refresh()
                DataGridView5.Show()
                'BindingSource5.Current("射孔枪状态") = "不工作"
                'DataGridView5.Refresh()
            End If
        End If
        ' 复制原始数据到新的DataTable，在副本上筛选排序
        Dim dtCopy As DataTable = DirectCast(BindingSource5.DataSource, DataTable).Copy()
        Dim row() As DataRow = dtCopy.Select("射孔枪状态='工作'", "射孔枪序号 ASC")
        Dim jczhd As Single = 0
        If row.Length > 0 Then
            Text22.Text = row(0).Item("到筛管的距离m").ToString()
            Textg23.Text = row(0).Item("射孔段起始深度m").ToString()
        Else
            Text22.Text = "0"
            Textg23.Text = "0"
        End If
        For Each jc As DataRow In row
            jczhd = jczhd + (CSng(jc.Item("终止深度m")) - CSng(jc.Item("起始深度m")))
        Next
        Textg15.Text = jczhd.ToString


    End Sub
    '*********************************************************************************************************************************************
    '填写所选工况的界面中各文本框内容
    '*********************************************************************************************************************************************
    Private Sub set_skgk_texts(ByVal dataFrom As String)
        If dataFrom = "工况表" Then
            If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
                Textg11.Text = If(IsDBNull(BindingSource1.Current("工况序号")), "0", BindingSource1.Current("工况序号").ToString)
                Text2.Text = If(IsDBNull(BindingSource1.Current("射孔施工方式")), "", BindingSource1.Current("射孔施工方式").ToString)
                Textg2.Text = If(IsDBNull(BindingSource1.Current("井口温度℃")), "0.0", BindingSource1.Current("井口温度℃").ToString)
                Textg3.Text = If(IsDBNull(BindingSource1.Current("射孔段处井底温度℃")), "0.0", BindingSource1.Current("射孔段处井底温度℃").ToString)
                Textg4.Text = If(IsDBNull(BindingSource1.Current("井底初始压力MPa")), "0.0", BindingSource1.Current("井底初始压力MPa").ToString)
                Textg5.Text = If(IsDBNull(BindingSource1.Current("井口加压MPa")), "0.0", BindingSource1.Current("井口加压MPa").ToString)
                Textg6.Text = If(IsDBNull(BindingSource1.Current("射孔液密度g╱cm3")), "0.0", BindingSource1.Current("射孔液密度g╱cm3").ToString)
                Textg7.Text = If(IsDBNull(BindingSource1.Current("工况名称")), "", BindingSource1.Current("工况名称").ToString)
                Textg8.Text = If(IsDBNull(BindingSource1.Current("射孔液粘度mPa·s")), "0.0", BindingSource1.Current("射孔液粘度mPa·s").ToString)
                Textg9.Text = If(IsDBNull(BindingSource1.Current("动载系数")), "0.0", BindingSource1.Current("动载系数").ToString)
                Text22.Text = If(IsDBNull(BindingSource1.Current("筛管距射孔顶端距离m")), "0.0", BindingSource1.Current("筛管距射孔顶端距离m").ToString)
                Textg10.Text = If(IsDBNull(BindingSource1.Current("人工井底m")), "0.0", BindingSource1.Current("人工井底m").ToString)
                Textg15.Text = If(IsDBNull(BindingSource1.Current("射孔夹层总厚度m")), "0.0", BindingSource1.Current("射孔夹层总厚度m").ToString)
                Textg23.Text = If(IsDBNull(BindingSource1.Current("射孔段起始深度m")), "0.0", BindingSource1.Current("射孔段起始深度m").ToString)
            End If
        Else
            Textg11.Text = "0"      '"工况序号"
            Text2.Text = ""         '"射孔施工方式"
            Textg2.Text = "0.0"     '"井口温度℃"
            Textg3.Text = "0.0"     '"射孔段处井底温度℃"
            Textg4.Text = "0.0"     '井底初始压力MPa"
            Textg5.Text = "0.0"     '井口加压MPa"
            Textg6.Text = "3.0"     '射孔液密度g╱cm3"，默认为清水
            Textg7.Text = ""        '工况名称"
            Textg8.Text = "0.0"     '[射孔液粘度mPa·s]
            Textg9.Text = "1.5"     '动载系数"
            Text22.Text = "0.0"     '筛管距射孔顶端距离m
            Textg10.Text = rgjd.ToString     '人工井底m
            Textg15.Text = "0.0"     '射孔夹层总厚度
            Textg23.Text = "0.0"     '射孔段起始深度
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除工况[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click
        Try

            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If RECreader.Read Then
                            msg_prompt = "是否确定要删除所选序号的工况参数数据？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            SQL_command = "DELETE * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 射孔段爆轰计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 射孔段封隔器计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 射孔油管应力计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            msg_prompt = "工况参数数据删除完成。"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            Call fill_gk_grid()
                            'Call fill_gkjctree()
                            Call fill_gkxztree()
                            Call fill_skq_wz_grid()
                            Call fill_fgq_grid()
                        Else
                            msg_prompt = "未找到工况序号对应的数据，请选好要删除的数据！"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "删除工况数据时数据库访问异，请联系软件编者解决！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Catch ex As Exception
            msg_prompt = "删除工况程序逻辑有误！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
        
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存工况[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command21_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command21.Click
        Dim SQL_command As String

        If Not IsNumeric(Textg11.Text) Or Val(Textg11.Text) <= 0 Then
            msg_prompt = "请输入合适的工况序号，需要为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text2.Text = "" Then
            msg_prompt = "请选择或输入射孔施工方式。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg2.Text) Then
            msg_prompt = "请输入合适的井口温度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg3.Text) Or Val(Textg3.Text) <= 0 Then
            msg_prompt = "请输入合适的井底温度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg4.Text) Or Val(Textg4.Text) <= 0 Then
            msg_prompt = "请输入合适的井底初始压力。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg6.Text) Or Val(Textg6.Text) <= 0 Then
            msg_prompt = "请输入合适的射孔液密度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg5.Text) Or Val(Textg5.Text) <= 0 Then
            msg_prompt = "请输入合适的井口加压。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Textg7.Text) = "" Then
            msg_prompt = "工况名称不能为空。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Textg8.Text) Or Val(Textg8.Text) < 0 Then
            msg_prompt = "射孔液粘度应为数字且大于0。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text22.Text) Or Val(Text22.Text) <= 0 Then
            msg_prompt = "请输入或计算筛管与射孔枪顶端的距离。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Dim worked_skq() As DataRow = skq_wz_table.Select("射孔枪状态='工作'", "射孔枪序号 asc")
        Dim first_worked_xh As Integer = 0
        Dim last_worked_xh As Integer = 0
        If worked_skq.length = 0 Then
            msg_prompt = "当前工况没有设定射孔夹层。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        Else
            first_worked_xh = worked_skq(0).Item("射孔枪序号")
            last_worked_xh = worked_skq(worked_skq.Length - 1).Item("射孔枪序号")
        End If
        Dim unworked_skq() As DataRow = skq_wz_table.Select("射孔枪状态='不工作'", "射孔枪序号 asc")
        Dim notLx As Boolean = False
        If worked_skq.Length > 0 And unworked_skq.Length > 0 Then
            For Each row As DataRow In unworked_skq
                If row.Item("射孔枪序号") > first_worked_xh And row.Item("射孔枪序号") < last_worked_xh Then
                    notLx = True
                    Exit For
                End If
            Next
        End If
        If notLx Then
            msg_prompt = "当前设定的夹层不连续，这不合理。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Try

            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Textg11.Text)
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If Not RECreader.Read Then
                            msg_prompt = "是否要建立新的射孔工况数据？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "insert into  射孔工况参数表 (作业名称,工况序号,井口温度℃,射孔段起始深度m,射孔夹层总厚度m,射孔段处井底温度℃,井底初始压力MPa,井口加压MPa,射孔液密度g╱cm3,工况名称,射孔施工方式," _
                                               & "[射孔液粘度mPa·s],射孔枪下井时间h,装弹数,计算总装药量方法,动载系数,筛管距射孔顶端距离m,人工井底m,井号) values (" _
                                & "'" & zuoye_name & "'," & Textg11.Text & "," & Textg2.Text & "," & Textg23.Text & "," & Textg15.Text & "," _
                                & Textg3.Text & "," & Textg4.Text & "," & Textg5.Text & "," & Textg6.Text & ",'" & Textg7.Text & "','" _
                                & Text2.Text & "'," & Textg8.Text & ",0,0,'夹层厚度'," & Textg9.Text & "," & Text22.Text & "," & Textg10.Text & ",'" & well_name & "')"
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                        Else
                            msg_prompt = "是否要保存对所选工况数据的修改？"
                            msg_buttons = 4 + 32
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                            If msg_return <> 6 Then
                                Exit Sub
                            End If
                            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using

                            SQL_command = "update 射孔工况参数表 set " _
                                & " 井口温度℃=" & Textg2.Text & "," & "射孔段处井底温度℃=" & Textg3.Text & ",井底初始压力MPa=" & Textg4.Text & "," & "井口加压MPa=" & Textg5.Text _
                                & ",射孔液密度g╱cm3=" & Textg6.Text & "," & "工况名称='" & Textg7.Text & "',射孔施工方式='" & Text2.Text & "'," & "[射孔液粘度mPa·s]=" & Textg8.Text _
                                & ",射孔枪下井时间h=0,装弹数=0,计算总装药量方法='夹层厚度',动载系数=" & Textg9.Text & ",人工井底m=" & Textg10.Text & ",射孔夹层总厚度m=" & Textg15.Text _
                                & ",射孔段起始深度m=" & Textg23.Text & ",筛管距射孔顶端距离m=" & Text22.Text & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            ''顺便删除该作业、工况的计算数据
                            SQL_command = "delete * from 射孔段爆轰计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 射孔段封隔器计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                            SQL_command = "delete * from 射孔油管应力计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & Textg11.Text
                            Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                EXECOleDbCommand2.ExecuteNonQuery()
                            End Using
                        End If
                        If gk_fgq_table.Rows.Count > 0 Then
                            For Each row As DataRow In gk_fgq_table.Rows
                                SQL_command = "insert into  工况_封隔定位元件 (作业名称,工况序号,工况名称,元件序号,元件名称,封隔器状态,定位方式,井号) values (" _
                                    & "'" & zuoye_name & "'," & CStr(Textg11.Text) & ",'" & Textg7.Text & "'," & row.Item("元件序号").ToString & "," & "'" & row.Item("元件名称").ToString & "'," _
                                    & "'" & row.Item("封隔器状态").ToString & "'," & "'未(无)锚定'," & "'" & well_name & "')"
                                Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                    EXECOleDbCommand2.ExecuteNonQuery()
                                End Using
                            Next
                        End If
                        If skq_wz_table.Rows.Count > 0 Then
                            For Each row As DataRow In skq_wz_table.Rows
                                SQL_command = "insert into  工况_射孔夹层表 (作业名称,工况序号,射孔枪序号,射孔枪名称,起始深度m,终止深度m,射孔枪状态,工况名称,井号) values (" _
                                & "'" & zuoye_name & "'," & Textg11.Text & "," & row.Item("射孔枪序号").ToString & ",'" & row.Item("射孔枪名称").ToString & "'," _
                                & row.Item("起始深度m").ToString & "," & row.Item("终止深度m").ToString & ",'" & row.Item("射孔枪状态").ToString & "','" & Textg7.Text & "','" & well_name & "')"
                                Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                    EXECOleDbCommand2.ExecuteNonQuery()
                                End Using
                            Next
                        End If
                    End Using
                End Using
            End Using
            Call fill_gk_grid()
            'Call fill_gkjctree()
            Call fill_gkxztree()
            Call fill_skq_wz_grid()
            Call fill_fgq_grid()
            msg_prompt = "工况参数数据保存完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If Not ListBox1.Items.Contains(Text2.Text) Then
                msg_prompt = "是否要将当前的射孔施工方式：" & Text2.Text & "加入的基础数据库的备选射孔施工方式数据表中？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return <> 6 Then
                    Exit Sub
                End If
                Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                    cn_basedb.Open()
                    SQL_command = "insert into  射孔施工方式表 (施工方式) values ('" & Text2.Text & "')"
                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand2.ExecuteNonQuery()
                    End Using
                    SQL_command = "select * from 射孔施工方式表 order by 施工方式"
                    Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_basedb)
                        Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                            ListBox1.Items.Clear()
                            While RECreader.Read
                                ListBox1.Items.Add(RECreader.Item("施工方式").ToString)
                            End While
                        End Using
                    End Using
                End Using
            End If
        Catch ex As OleDbException
            msg_prompt = "保存数据过程数据库访问异常！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Catch ex As Exception
            msg_prompt = "保存数据过程逻辑有误！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try

    End Sub
    '*********************************************************************************************************************************************
    '界面Closed。
    '*********************************************************************************************************************************************
    Private Sub ipt_gkuang_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '单击“计算”按钮，计算井底初始压力MPa=静液柱压力
    '*********************************************************************************************************************************************
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Val(Textg6.Text) <= 0 And Not IsNumeric(Textg6.Text) Then
            msg_prompt = "请先输入射孔液密度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Textg10.Text) <= 0 And Not IsNumeric(Textg10.Text) Then
            msg_prompt = "请先输入有效的人工井底。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Dim fwj, jxj, jinShen As Single
        Dim jdcsyl As Single
        If skq_wz_table.Rows.Count > 0 Then
            Call cal_jx_fw_cs(CSng(Textg10.Text), jxj, fwj, jinShen)
            jdcsyl = CSng(Textg6.Text) * 9.81 * jinShen / 1000
            Textg4.Text = Math.Round(jdcsyl, 4)
        Else
            msg_prompt = "射孔枪位置表没有填充。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub

End Class

