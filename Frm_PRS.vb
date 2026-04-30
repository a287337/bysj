Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class Frm_PRS
    '*********************************************************************************************************************************************
    '                                                   关于管柱力学节点分析、显示界面窗体的说明
    '                                                                                                           秦彦斌 2022年11月5日最后整理更新
    ' 说明：
    '     一个界面，通过不同侧面观察、处理计算及计算结果。
    '
    ' 程序升级记事：
    '   （1）修正了在保存工况参数时，若修改了工况名称，待选工况名称没有修改的BUG
    '   （2）用iPlotX绘制曲线时，该控件默认最多自动给出7种颜色，而且有白色及黄色太浅，分辨度不够。为此：
    '             在基础数据库中增加("绘图颜色表")，自动加入经反复对比找出的26种分辨度高颜色。详见文档《RGB颜色对照表.docx》
    '             提供颜色查取函数，传入颜色序号，1开头，返回Uint32值
    '             iPlotX使用颜色时，次序时B.G.R，故库里次序也是B.G.R
    '    (3) 设计算结果的最大、最小值计算参数极限值表MaxMinCSTable，Load事件或重算后填充，其它操作时读取，可减少访问数据库表次数。 
    '    (4) 经反复试验，DataGridView对应的表结构不得变化，否则绑定、显示无法实现。在“按节点显示-按类型分工况”页中，原想选定按节及数据类型后，
    '        页面表中仅显示该节点该类型不同工况的值，但编程中发现应临时表结构会因类型而变，DataGridView12无法随着变化，故采用一个BindingSource9
    '        绑定3个DataGridView。
    '    (5) 在减少访问数据库名的算法上还有工作可做。
    '    (6) 20211117 重新布置界面控件排列，实现可缩放。
    '    (7) 20220308 在按数据类型显示页增加“内外压差(MPa)”单选钮。实现选真实轴力时绘制管体抗拉强度和接头加抗拉强度线；选内外压差时绘制管体/
    '        接头抗内压强度和抗外挤强度度线。
    '    (8) 20220308-0320 舍弃原“按节点显示”页的图形，代以管柱单轴双轴三轴应力安全性判断包络图并绘制所选管柱段/点工作状态线/点。
    '        实现了考虑管柱等效应力强度安全系数、管柱抗外挤强度安全系数、管体抗内压强度安全系数、管体抗拉强度安全系数、接头抗内压强度安全系数、
    '        接头抗拉强度安全系数的综合校核。
    '    (9) 20220906-0907 舍弃两个TablelayoutPanel控件，代以用GroupBox画的表格，提高了界面加载速度。
    '    (10) 20221101-1105 （1）在“按工况显示-计算结果数值”页中，将选定工况后，列表显示该工况下伸缩管状态改为：增加工况序号、工况名称列，一股
    '         脑显示所有工况下伸缩管状态，以方便比较且可减少访问数据库的次数。（2）在“按工况显示-工况参数页”中，改选定工况后列出封隔器、锚定工
    '         具、开关工具在此工况下的工作状态为一股脑显示所有工况下各工具的工作状态，以方便比较且可减少访问数据库的次数。（3）在“按工况显示-计
    '         算结果数值”页的选定工况节点计算参数列表中增加“工况序号”列，以便确定所列数据就是所选工况的。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '计算参数极限值表赋值
    '*********************************************************************************************************************************************
    Private Sub Fill_MaxMinCSTable()
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select max(节点下深m) as max_xs," _
            & " min(节点温度℃) as min_jd_wendu, " & " max(节点温度℃) as max_jd_wendu, " & " min(温度变形m) as min_bx_wd, " & " max(温度变形m) as max_bx_wd, " _
            & " min(轴力变形m) as min_bx_zl, " & " max(轴力变形m) as max_bx_zl, " & " min(鼓胀变形m) as min_bx_gz, " & " max(鼓胀变形m) as max_bx_gz, " _
            & " min(螺旋变形m) as min_bx_lx, " & " max(螺旋变形m) as max_bx_lx, " & " min(综合变形m) as min_bx_zh, " & " max(综合变形m) as max_bx_zh, " _
            & " min(温度效应m) as min_xy_wd, " & " max(温度效应m) as max_xy_wd, " & " min(轴力效应m) as min_xy_zl, " & " max(轴力效应m) as max_xy_zl, " _
            & " min(鼓胀效应m) as min_xy_gz, " & " max(鼓胀效应m) as max_xy_gz, " & " min(螺旋效应m) as min_xy_lx, " & " max(螺旋效应m) as max_xy_lx, " _
            & " min(综合效应m) as min_xy_zh, " & " max(综合效应m) as max_xy_zh, " & " min(合成应力MPa) as min_xgm4, " & " max(合成应力MPa) as max_xgm4, " _
            & " min(安全系数) as min_aqxs, " & " max(安全系数) as max_aqxs, " & " min(接触力N) as min_jchl, " & " max(接触力N) as max_jchl, " _
            & " min(合弯矩Nm) as min_hwj, " & " max(合弯矩Nm) as max_hwj, " & " min(真实轴力N) as min_shz, " & " max(真实轴力N) as max_shz, " _
            & " min(等效轴力N) as min_xz, " & " max(等效轴力N) as max_xz, " & " min(管内压力MPa) as min_gnyl, " & " max(管内压力MPa) as max_gnyl, " _
            & " min(等效悬持力N) as min_dxxcl, " & " max(等效悬持力N) as max_dxxcl, " & " min(管外压力MPa) as min_gwyl, " & " max(管外压力MPa) as max_gwyl " _
            & " from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        MaxMinCSTable.Clear()
        ad.Fill(MaxMinCSTable)
        ad.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Sub
    '*********************************************************************************************************************************************
    '界面Closing
    '*********************************************************************************************************************************************
    Private Sub Frm_TSM_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call write_chanshu("管柱力学分析-抗外挤强度安全系数", CStr(TextBox7.Text))
        Call write_chanshu("管柱力学分析-抗内压强度安全系数", CStr(TextBox8.Text))
        Call write_chanshu("管柱力学分析-抗拉强度安全系数", CStr(TextBox9.Text))
        Call write_chanshu("管柱力学分析-三轴强度安全系数", CStr(Text20.Text))
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub Frm_TSM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim pageCount As Integer
        Dim totalWidth As Integer
        Dim pageWidth As Integer
        Dim i As Short
        Dim c_canshu As String
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim jdqk_row As DataRow
        Dim rowfilter As String
        Dim qzh_length As Double
        Label2.Text = well_name & "井" & zuoye_name & "作业管柱力学分析"
        'Me.Text = "管柱力学分析TSM-" & well_name & "井" & zuoye_name & "作业"
        c_canshu = get_canshu("管柱力学分析-节点间距")
        If c_canshu <> "" Then
            Text1.Text = CStr(Val(c_canshu))
        Else
            Text1.Text = CStr(25.0#)
        End If

        For i = 0 To 23
            Text2(i).Text = ""
        Next i
        _Text2_24.Text = ""
        _Text2_25.Text = ""
        _Text2_26.Text = ""
        _Text2_27.Text = ""
        _Text2_28.Text = ""
        _Text2_29.Text = ""
        _Text2_30.Text = ""
        _Text2_31.Text = ""
        _Text2_32.Text = ""
        _Text2_33.Text = ""
        _Text2_34.Text = ""
        _Text2_35.Text = ""
        _Text2_36.Text = ""
        _Text2_37.Text = ""
        _Text2_38.Text = ""
        _Text2_39.Text = ""
        _Text2_40.Text = ""
        _Text2_41.Text = ""
        _Text2_42.Text = ""
        _Text2_43.Text = ""
        _Text2_44.Text = ""
        _Text2_45.Text = ""
        _Text2_46.Text = ""
        _Text2_47.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        '抗外挤强度安全系数
        c_canshu = get_canshu("管柱力学分析-抗外挤强度安全系数")
        If c_canshu <> "" Then
            TextBox7.Text = CStr(Val(c_canshu))
        Else
            TextBox7.Text = CStr(1.4)
        End If
        '抗内压强度安全系数
        c_canshu = get_canshu("管柱力学分析-抗内压强度安全系数")
        If c_canshu <> "" Then
            TextBox8.Text = CStr(Val(c_canshu))
        Else
            TextBox8.Text = CStr(1.25)     '抗内压强度安全系数
        End If
        '抗拉强度安全系数
        c_canshu = get_canshu("管柱力学分析-抗拉强度安全系数")
        If c_canshu <> "" Then
            TextBox9.Text = CStr(Val(c_canshu))
        Else
            TextBox9.Text = CStr(1.6)     '抗拉强度安全系数
        End If
        '三轴强度安全系数
        c_canshu = get_canshu("管柱力学分析-三轴强度安全系数")
        If c_canshu <> "" Then
            Text20.Text = CStr(Val(c_canshu))
        Else
            Text20.Text = CStr(1.5)       '三轴强度安全系数
        End If
        '*********************************************************************************************************************************************
        '让主界面中SSTab1各页标题宽度一样，不要挤在一起。
        '*********************************************************************************************************************************************
        '这个Fixed设置是必须的
        SSTab1.SizeMode = TabSizeMode.Fixed
        '设置标签宽度
        totalWidth = SSTab1.Width
        pageCount = SSTab1.TabPages.Count
        '最后-1 因为tabcontrol有留margin，得空出margin的空间
        pageWidth = totalWidth / pageCount - 2
        'ItemSize在SizeMode = TabSizeMode.Fixed才生效
        SSTab1.ItemSize = New Size(pageWidth, SSTab1.ItemSize.Height)
        '*********************************************************************************************************************************************
        '让第1页“按工况显示页”中SSTab2各页标题宽度一样，不要挤在一起。
        '*********************************************************************************************************************************************
        '这个Fixed设置是必须的
        SSTab2.SizeMode = TabSizeMode.Fixed
        '设置标签宽度
        totalWidth = SSTab2.Width
        pageCount = SSTab2.TabPages.Count
        '最后-1 因为tabcontrol有留margin，得空出margin的空间
        pageWidth = totalWidth / pageCount - 2
        'ItemSize在SizeMode = TabSizeMode.Fixed才生效
        SSTab2.ItemSize = New Size(pageWidth, SSTab2.ItemSize.Height)
        '*********************************************************************************************************************************************
        '让第2页“按节点显示页”中SSTab3各页标题宽度一样，不要挤在一起。
        '*********************************************************************************************************************************************
        '这个Fixed设置是必须的
        'SSTab3.SizeMode = TabSizeMode.Fixed
        ''设置标签宽度
        'totalWidth = SSTab3.Width
        'pageCount = SSTab3.TabPages.Count
        ''最后-1 因为tabcontrol有留margin，得空出margin的空间
        'pageWidth = totalWidth / pageCount - 1
        ''ItemSize在SizeMode = TabSizeMode.Fixed才生效
        'SSTab3.ItemSize = New Size(pageWidth, SSTab3.ItemSize.Height)
        '*********************************************************************************************************************************************
        '计算参数极限值表赋值
        '*********************************************************************************************************************************************
        Call Fill_MaxMinCSTable()

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '*********************************************************************************************************************************************
        '第1页“按工况显示页”中工况列表赋初值，原来用Adodc1， DataGrid1
        '*********************************************************************************************************************************************
        SQL_command = "select 工况序号,工况名称 from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_xhmc_Table.Clear()
        ad.Fill(gk_xhmc_Table)
        ad.Dispose()
        BindingSource1.DataSource = gk_xhmc_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        'DataGridView1.RowHeadersWidth = 24
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Columns(0).Width = 30
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Refresh()
        DataGridView1.Show()
        '封隔器状态列表
        SQL_command = "select 工况序号,元件序号,元件名称,封隔器状态,定位方式 from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号,元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_fgq_Table.Clear()
        ad.Fill(gk_fgq_Table)
        ad.Dispose()
        BindingSource4.DataSource = gk_fgq_Table
        DataGridView5.ClearSelection()
        DataGridView5.DataSource = BindingSource4
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
        DataGridView5.Columns(0).Width = DataGridView5.Width / 6 - 5
        DataGridView5.Columns(1).Width = DataGridView5.Width / 6 - 5
        DataGridView5.Columns(2).Width = DataGridView5.Width / 6 * 2
        DataGridView5.Columns(3).Width = DataGridView5.Width / 6
        DataGridView5.Columns(4).Width = DataGridView5.Width / 6
        DataGridView5.Refresh()
        DataGridView5.Show()
        '开关工具状态列表
        SQL_command = "select 工况序号,元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号,元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_kggj_Table.Clear()
        ad.Fill(gk_kggj_Table)
        ad.Dispose()
        BindingSource5.DataSource = gk_kggj_Table
        DataGridView6.ClearSelection()
        DataGridView6.DataSource = BindingSource5
        DataGridView6.ResetBindings()
        DataGridView6.AutoGenerateColumns = True
        DataGridView6.AllowUserToAddRows = False
        DataGridView6.AllowUserToDeleteRows = False
        DataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView6.MultiSelect = False
        DataGridView6.RowHeadersWidth = 24
        DataGridView6.ReadOnly = True
        DataGridView6.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView6.Columns(0).Width = DataGridView6.Width / 7 - 5
        DataGridView6.Columns(1).Width = DataGridView6.Width / 7 - 5
        DataGridView6.Columns(2).Width = DataGridView6.Width / 7 * 2
        DataGridView6.Columns(3).Width = DataGridView6.Width / 7
        DataGridView6.Columns(4).Width = DataGridView6.Width / 7
        DataGridView6.Columns(5).Width = DataGridView6.Width / 7
        DataGridView6.Refresh()
        DataGridView6.Show()
        '锚定工具状态列表()
        SQL_command = "select 工况序号,元件序号,元件名称,定位方式 " & " from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by  工况序号,元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_mdgj_Table.Clear()
        ad.Fill(gk_mdgj_Table)
        ad.Dispose()
        BindingSource6.DataSource = gk_mdgj_Table
        DataGridView7.ClearSelection()
        DataGridView7.DataSource = BindingSource6
        DataGridView7.ResetBindings()
        DataGridView7.AutoGenerateColumns = True
        DataGridView7.AllowUserToAddRows = False
        DataGridView7.AllowUserToDeleteRows = False
        DataGridView7.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView7.MultiSelect = False
        DataGridView7.RowHeadersWidth = 24
        DataGridView7.ReadOnly = True
        DataGridView7.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView7.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView7.Columns(0).Width = DataGridView7.Width / 6 - 5
        DataGridView7.Columns(1).Width = DataGridView7.Width / 6 - 5
        DataGridView7.Columns(2).Width = DataGridView7.Width / 6 * 3
        DataGridView7.Columns(3).Width = DataGridView7.Width / 6
        DataGridView7.Refresh()
        DataGridView7.Show()
        '*********************************************************************************************************************************************
        '第1-2页即按工况显示-工况参数页中管柱组成简表赋初值,原来用Adodc3， DataGrid3
        '*********************************************************************************************************************************************
        SQL_command = "select 元件序号 as 序号,元件性质 as 类型,元件名称 as 规格,元件外径mm as 外径（mm）,元件内径mm as 内径（mm）,元件长度m as 长度（m） from 管柱数据表 where 井号='" _
            & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        sp_gzh_Table.Clear()
        ad.Fill(sp_gzh_Table)
        ad.Dispose()
        '求管柱总长
        qzh_length = 0.0
        For Each jdqk_row In sp_gzh_Table.Rows
            qzh_length = qzh_length + Val(jdqk_row.Item("长度（m）").ToString)
        Next
        Label3.Text = "总长" & Trim(CStr(qzh_length)) & "m"
        BindingSource7.DataSource = sp_gzh_Table
        DataGridView8.ClearSelection()
        DataGridView8.DataSource = BindingSource7
        DataGridView8.ResetBindings()
        DataGridView8.AutoGenerateColumns = True
        DataGridView8.AllowUserToAddRows = False
        DataGridView8.AllowUserToDeleteRows = False
        DataGridView8.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView8.MultiSelect = False
        DataGridView8.RowHeadersWidth = 24
        DataGridView8.ReadOnly = True
        DataGridView8.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView8.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView8.Columns(0).Width = 60
        DataGridView8.Columns(1).Width = 110
        DataGridView8.Columns(2).Width = 300
        DataGridView8.Columns(3).Width = 120
        DataGridView8.Columns(4).Width = 120
        DataGridView8.Columns(5).Width = 120
        DataGridView8.Refresh()
        DataGridView8.Show()
        '*********************************************************************************************************************************************
        '第2页"按节点显示"页拟报告输出节点列表            原来第2页"按节点显示"页拟报告输出节点列表用Adodc9和DataGrid11,数据表用lsb_repoutlist
        '*********************************************************************************************************************************************
        SQL_command = "select 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m from 节点情况表,节点计算参数表 where  0>1"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        pg2_prtjdlb_Table.Clear()
        ad.Fill(pg2_prtjdlb_Table)
        ad.Dispose()
        BindingSource10.DataSource = pg2_prtjdlb_Table
        DataGridView13.ClearSelection()
        DataGridView13.DataSource = BindingSource10
        DataGridView13.ResetBindings()
        DataGridView13.AutoGenerateColumns = True
        DataGridView13.AllowUserToAddRows = False
        DataGridView13.AllowUserToDeleteRows = False
        DataGridView13.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView13.MultiSelect = False
        DataGridView13.RowHeadersWidth = 24
        DataGridView13.ReadOnly = True
        DataGridView13.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView13.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView13.Columns(0).Width = 35
        DataGridView13.Columns(1).Width = 55
        DataGridView13.Columns(2).Width = 70
        DataGridView13.Columns(3).Width = 95
        DataGridView13.Refresh()
        DataGridView13.Show()
        '*********************************************************************************************************************************************
        '建立第2页拟打印节点临时表
        '*********************************************************************************************************************************************
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='lsb_repoutlist'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE lsb_repoutlist"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        dbSchema.Clear()
        dbSchema.Dispose()
        SQL_command = "select 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m into lsb_repoutlist from 节点情况表,节点计算参数表 where  0>1"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
        '第1页"按工况显示"页中各DataGridView填充
        Call page12_jdlb_fill()
        '第1-1页图形初始化程序，前提是需要工况列表不为空
        Call draw_init()
        '按工况显示-计算结果数值页，显示所有工况下伸缩管状态列表填充
        Call page12_ssglb_fill()
        '第1-3页即按工况显示-工况参数页赋值与填写
        Call gkxs_gkchange()
        '第2页节点列表赋值
        Call page2_dxjdlb_fill()
        '第2页"按节点显示"页中各DataGridView填充
        Call page2_DataGridView_flash()
        '第2-1页图形初始化
        Call draw_init3()
        '第3页图形初始化程序，前提是需要工况列表不为空
        Call draw_init2()
        '第3页节点列表赋值
        Call page3_jdlb_fill()
        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
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
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub Frm_TSM_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim rowfilter As String

        '清理无用临时表
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='lsb_repoutlist'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE lsb_repoutlist"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        dbSchema.Clear()
        dbSchema.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()

        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '点击第1页即”按工况显示“页工况列表，刷新各DataGridView。
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Call page12_jdlb_fill()
        Call gkxs_gkchange()
        Call draw_init()
    End Sub
    ''*********************************************************************************************************************************************
    ''第1页即”按工况显示“页工况列表选择变化时，刷新各DataGridView。
    ''*********************************************************************************************************************************************
    'Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
    '    Call page12_jdlb_fill()
    '    Call gkxs_gkchange()
    '    Call draw_init()
    'End Sub

    Private Sub page12_ssglb_fill()
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '各工况伸缩管状态列表
        SQL_command = "select 工况_伸缩管状态.工况序号,工况参数表.工况名称,工况_伸缩管状态.元件序号,工况_伸缩管状态.元件名称,剪销状态,伸缩长度m," _
                & " 伸缩状态,伸缩行程m,全缩短长度m,管柱数据表.元件长度m as 下入长度m,本工况初始长度m,伸缩动作载荷kN," _
                & " 伸缩动作压力MPa,剪切力kN,不伸缩等效轴力kN,等效轴力kN,压差剪切力kN,内压MPa,外压MPa " _
                & " from 工况_伸缩管状态,管柱数据表,工况参数表 where" _
                & " 工况_伸缩管状态.元件序号=管柱数据表.元件序号 and 工况_伸缩管状态.井号=管柱数据表.井号 and 工况_伸缩管状态.作业名称=管柱数据表.作业名称 " _
                & " and 工况_伸缩管状态.井号=工况参数表.井号 and 工况_伸缩管状态.作业名称=工况参数表.作业名称 and 工况_伸缩管状态.工况序号=工况参数表.工况序号 " _
                & " and 工况_伸缩管状态.井号='" & well_name & "' and 工况_伸缩管状态.作业名称='" & zuoye_name & "'" _
                & " order by  工况_伸缩管状态.工况序号,工况_伸缩管状态.元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_ssgzt_Table.Clear()
        ad.Fill(gk_ssgzt_Table)
        ad.Dispose()
        BindingSource3.DataSource = gk_ssgzt_Table
        DataGridView4.ClearSelection()
        DataGridView4.DataSource = BindingSource3
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
        DataGridView4.Columns(0).Width = DataGridView4.Width / 34 - 5
        DataGridView4.Columns(1).Width = DataGridView4.Width / 34 * 3
        DataGridView4.Columns(2).Width = DataGridView4.Width / 34 - 5
        DataGridView4.Columns(3).Width = DataGridView4.Width / 34 * 3
        DataGridView4.Columns(4).Width = DataGridView4.Width / 34 * 1.6
        DataGridView4.Columns(5).Width = DataGridView4.Width / 34 * 1.5
        DataGridView4.Columns(6).Width = DataGridView4.Width / 34 * 1.6
        DataGridView4.Columns(7).Width = DataGridView4.Width / 34 * 1.3
        DataGridView4.Columns(8).Width = DataGridView4.Width / 34 * 1.3
        DataGridView4.Columns(9).Width = DataGridView4.Width / 34 * 1.3
        DataGridView4.Columns(10).Width = DataGridView4.Width / 34 * 1.3
        DataGridView4.Columns(11).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(12).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(13).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(14).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(15).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(16).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(17).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Columns(18).Width = DataGridView4.Width / 34 * 2
        DataGridView4.Refresh()
        DataGridView4.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取各工况下伸缩管状态数据时出错，可能是软件升级数据结构发生变化，点击【计算】按钮可消除此错误。若仍未解决，请联系软件开发人员解决。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '    第1-2页"按工况显示"-"计算结果数值"页,所有节点check1变化时
    '*********************************************************************************************************************************************
    Private Sub Check1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Check1.CheckedChanged
        Call page12_jdlb_fill()
    End Sub
    '*********************************************************************************************************************************************
    '    第1-2页"按工况显示"-"计算结果数值"页,填充以下内容：
    '(1)选定工况待选节点列表                                      原来选定工况待选节点列表用Adodc4和DataGrid4
    '(2)选定工况节点计算参数列表                                  原来选定工况节点计算参数列表用Adodc2和DataGrid2
    '*********************************************************************************************************************************************
    Private Sub page12_jdlb_fill()
        Dim gk_xh_str As String
        Dim i As Integer
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        gk_xh_str = "0"
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
            gk_xh_str = IIf(BindingSource1.Current("工况序号").ToString = "", "0", BindingSource1.Current("工况序号").ToString)
        End If
        If Check1.CheckState = 0 Then
            SQL_command = "select 节点计算参数表.工况序号,节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m,套管内径mm,油管外径mm, 油管内径mm,管内压力MPa," _
                & " 管外压力MPa,真实轴力N,等效轴力N ,接触力N, 合弯矩Nm, 节点温度℃, 温度变形m, 轴力变形m, 鼓胀变形m, 螺旋变形m, 综合变形m," _
                & " 温度效应m, 轴力效应m, 鼓胀效应m, 螺旋效应m, 综合效应m,合成应力MPa, 安全系数,等效悬持力N,屈服强度MPa,油管钢级,线重kg╱m," _
                & " Fecrh_N,曲率rad╱m,油套摩擦系数, 环液密度g╱cm3,管液密度g╱cm3,扭矩Nm,节点计算参数表.抗挤强度MPa" _
                & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                & " and 节点性质<>'计算点' and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                & " and 节点计算参数表.工况序号=" & gk_xh_str & " order by 节点计算参数表.节点编号"
        Else
            SQL_command = "select 节点计算参数表.工况序号,节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m,套管内径mm,油管外径mm, 油管内径mm,管内压力MPa," _
                & " 管外压力MPa,真实轴力N,等效轴力N ,接触力N, 合弯矩Nm, 节点温度℃, 温度变形m, 轴力变形m, 鼓胀变形m, 螺旋变形m, 综合变形m," _
                & " 温度效应m, 轴力效应m, 鼓胀效应m, 螺旋效应m, 综合效应m,合成应力MPa, 安全系数,等效悬持力N,屈服强度MPa,油管钢级,线重kg╱m," _
                & " Fecrh_N,曲率rad╱m,油套摩擦系数, 环液密度g╱cm3,管液密度g╱cm3,扭矩Nm,节点计算参数表.抗挤强度MPa" _
                & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                & " and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                & " and 节点计算参数表.工况序号=" & gk_xh_str & " order by 节点计算参数表.节点编号"
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_jd_jscs_Table.Clear()
        ad.Fill(gk_jd_jscs_Table)
        ad.Dispose()
        BindingSource2.DataSource = gk_jd_jscs_Table
        '选定工况节点计算参数列表
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
        DataGridView2.Refresh()
        DataGridView2.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取选定工况的节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '    第1-3页即按工况显示-工况参数页赋值与填写
    '*********************************************************************************************************************************************
    Private Sub gkxs_gkchange()
        Dim gk_xh_str As String
        On Error GoTo errhandler
        gk_xh_str = "0"
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
            gk_xh_str = IIf(BindingSource1.Current("工况序号").ToString = "", "0", BindingSource1.Current("工况序号").ToString)
        End If
        Text53.Text = gk_xh_str
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.Read Then
            Text61.Text = RECreader.Item("工况名称").ToString()
            Text54.Text = RECreader.Item("井口温度℃").ToString()
            Text7.Text = RECreader.Item("井底温度℃").ToString()
            Text9.Text = RECreader.Item("管压井底深度m").ToString()
            Text56.Text = RECreader.Item("井口环压MPa").ToString()
            Text57.Text = RECreader.Item("环液深度m").ToString()
            Text58.Text = RECreader.Item("环液密度g╱cm3").ToString()
            Text60.Text = RECreader.Item("环液粘度mPaS").ToString()
            Text59.Text = RECreader.Item("环流流量m3╱m").ToString()
            Text64.Text = RECreader.Item("井口管压MPa").ToString()
            Text65.Text = RECreader.Item("管液深度m").ToString()
            Text66.Text = RECreader.Item("管液密度g╱cm3").ToString()
            Text68.Text = RECreader.Item("管流粘度mPaS").ToString()
            Text67.Text = RECreader.Item("管流流量m3╱m").ToString()
            Text63.Text = RECreader.Item("库摩系数").ToString()
            Text26.Text = RECreader.Item("井口加力kN").ToString()
            Combo3.Text = RECreader.Item("环空流体流向").ToString()
            Combo4.Text = RECreader.Item("管内流体流向").ToString()
            Text8.Text = RECreader.Item("井口加扭Nm").ToString()
            Text3.Text = RECreader.Item("井底套压MPa").ToString()
            Text4.Text = RECreader.Item("井底管压MPa").ToString()
            Combo1.Text = RECreader.Item("井底套压开关").ToString()
            Combo2.Text = RECreader.Item("井底管压开关").ToString()
            Combo6.Text = RECreader.Item("井口套压开关").ToString()
            Combo5.Text = RECreader.Item("井口管压开关").ToString()
            Text5.Text = RECreader.Item("套压井底深度m").ToString()
            Combo7.Text = RECreader.Item("环空流阻模型").ToString()
            Combo8.Text = RECreader.Item("管内流阻模型").ToString()
            Text6.Text = RECreader.Item("环空稠剂浓度").ToString()
            Text10.Text = RECreader.Item("环空撑剂浓度").ToString()
            Text11.Text = RECreader.Item("环空流变指数n").ToString()
            Text12.Text = RECreader.Item("环空稠度系数K").ToString()
            Text13.Text = RECreader.Item("管内稠剂浓度").ToString()
            Text14.Text = RECreader.Item("管内撑剂浓度").ToString()
            Text15.Text = RECreader.Item("管内流变指数n").ToString()
            Text16.Text = RECreader.Item("管内稠度系数K").ToString()
            Text18.Text = RECreader.Item("环牛模折减系数").ToString()
            Text19.Text = RECreader.Item("管牛模折减系数").ToString()
        Else
            Text61.Text = ""
            Text54.Text = CStr(20.0#)
            Text7.Text = CStr(20.0#)
            Text9.Text = CStr(0.0#)
            Text56.Text = CStr(0.0#)
            Text57.Text = CStr(0.0#)
            Text58.Text = CStr(0.0#)
            Text60.Text = CStr(0.0#)
            Text59.Text = CStr(0.0#)
            Text64.Text = CStr(0.0#)
            Text65.Text = CStr(0.0#)
            Text66.Text = CStr(0.0#)
            Text68.Text = CStr(0.0#)
            Text67.Text = CStr(0.0#)
            '***************************************************************************************************************
            '根据机械设计手册：
            '    钢-钢     无润滑  静摩擦系数  0.15，   有润滑  静摩擦系数  0.1-0.12
            '    钢-钢     无润滑  动摩擦系数  0.1，    有润滑  动摩擦系数  0.05-0.1
            ' 故这里取值为  0.08
            '***************************************************************************************************************
            Text63.Text = CStr(0.08#)
            Text26.Text = CStr(0.0#)
            Combo3.Text = "不流动"
            Combo4.Text = "不流动"
            Text8.Text = CStr(0.0#)
            Text3.Text = CStr(0.0#)
            Text4.Text = CStr(0.0#)
            Combo1.Text = ""
            Combo2.Text = ""
            Combo6.Text = ""
            Combo5.Text = ""
            Text5.Text = CStr(0.0#)
            Combo7.Text = "无"
            Combo8.Text = "无"
            Text6.Text = CStr(0.0#)
            Text10.Text = CStr(0.0#)
            Text11.Text = CStr(0.0#)
            Text12.Text = CStr(0.0#)
            Text13.Text = CStr(0.0#)
            Text14.Text = CStr(0.0#)
            Text15.Text = CStr(0.0#)
            Text16.Text = CStr(0.0#)
            Text18.Text = CStr(0.35)
            Text19.Text = CStr(0.35)
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取工况参数数据时出错，请确认工况参数输入的正确性。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 点击第1-2页即按工况显示-计算结果数值页中待选节点列表
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        Call gk_jd_change()
    End Sub
    '*********************************************************************************************************************************************
    ' 第1-2页即按工况显示-计算结果数值页中待选节点列表选择变化时
    '*********************************************************************************************************************************************
    'Private Sub DataGridView2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Validated
    '    Call gk_jd_change()
    'End Sub
    '*********************************************************************************************************************************************
    ' 按工况显示-计算结果数值页中，节点变化时，将计算参数赋值给各文本框以刷新显示
    '*********************************************************************************************************************************************
    Private Sub gk_jd_change()
        On Error GoTo errhandler
        If (Not IsNothing(Me.BindingSource2.Current)) Then
            If (Not IsDBNull(Me.BindingSource2.Current("节点编号"))) Then
                Text2(0).Text = IIf(BindingSource2.Current("套管内径mm").ToString = "", "0", BindingSource2.Current("套管内径mm").ToString)
                Text2(1).Text = BindingSource2.Current("油管钢级").ToString
                Text2(2).Text = IIf(BindingSource2.Current("油管外径mm").ToString = "", "0", BindingSource2.Current("油管外径mm").ToString)
                Text2(3).Text = IIf(BindingSource2.Current("油管内径mm").ToString = "", "0", BindingSource2.Current("油管内径mm").ToString)
                Text2(4).Text = IIf(BindingSource2.Current("屈服强度MPa").ToString = "", "0", BindingSource2.Current("屈服强度MPa").ToString)
                Text2(5).Text = IIf(BindingSource2.Current("节点温度℃").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("节点温度℃").ToString)) / 1000))
                Text2(6).Text = IIf(BindingSource2.Current("管外压力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("管外压力MPa").ToString)) / 1000))
                Text2(7).Text = IIf(BindingSource2.Current("管内压力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("管内压力MPa").ToString)) / 1000))
                Text2(8).Text = IIf(BindingSource2.Current("真实轴力N").ToString = "", "0", CStr(Fix(Val(BindingSource2.Current("真实轴力N").ToString)) / 1000))
                Text2(9).Text = IIf(BindingSource2.Current("接触力N").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("接触力N").ToString)) / 1000))
                Text2(10).Text = IIf(BindingSource2.Current("等效轴力N").ToString = "", "0", CStr(Fix(Val(BindingSource2.Current("等效轴力N").ToString)) / 1000))
                Text2(11).Text = IIf(BindingSource2.Current("合弯矩Nm").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("合弯矩Nm").ToString)) / 1000))
                Text2(12).Text = IIf(BindingSource2.Current("合成应力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("合成应力MPa").ToString)) / 1000))
                Text2(13).Text = IIf(BindingSource2.Current("安全系数").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("安全系数").ToString)) / 1000))
                Text2(14).Text = IIf(BindingSource2.Current("温度变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("温度变形m").ToString)) / 1000))
                Text2(15).Text = IIf(BindingSource2.Current("温度效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("温度效应m").ToString)) / 1000))
                Text2(16).Text = IIf(BindingSource2.Current("轴力变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("轴力变形m").ToString)) / 1000))
                Text2(17).Text = IIf(BindingSource2.Current("轴力效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("轴力效应m").ToString)) / 1000))
                Text2(18).Text = IIf(BindingSource2.Current("鼓胀变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("鼓胀变形m").ToString)) / 1000))
                Text2(19).Text = IIf(BindingSource2.Current("鼓胀效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("鼓胀效应m").ToString)) / 1000))
                Text2(20).Text = IIf(BindingSource2.Current("螺旋变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("螺旋变形m").ToString)) / 1000))
                Text2(21).Text = IIf(BindingSource2.Current("螺旋效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("螺旋效应m").ToString)) / 1000))
                Text2(22).Text = IIf(BindingSource2.Current("综合变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("综合变形m").ToString)) / 1000))
                Text2(23).Text = IIf(BindingSource2.Current("综合效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource2.Current("综合效应m").ToString)) / 1000))
            End If
        Else
            msg_prompt = "    所选工况没有节点数据可显示，请点击[计算]按钮进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取所选工况节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 点击第1-2页即按工况显示-工况参数页中“保存工况参数修改”按钮
    '*********************************************************************************************************************************************
    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
        Dim foundRows() As DataRow
        Dim rowfilter As String
        On Error GoTo errhandler
        If Text61.Text = "" Then
            msg_prompt = "请输入工况名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text = "输入" And (Val(Text9.Text) = 0 Or (Not IsNumeric(Text9.Text)) Or IsDBNull(Text9.Text)) Then
            msg_prompt = "请输入井底管内压力所对应的井底深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "输入" And (Val(Text5.Text) = 0 Or (Not IsNumeric(Text5.Text)) Or IsDBNull(Text5.Text)) Then
            msg_prompt = "请输入井底环空压力所对应的井底深度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '***************************************************************************************************************
        '根据机械设计手册：
        '    钢-钢     无润滑  静摩擦系数  0.15，   有润滑  静摩擦系数  0.1-0.12
        '    钢-钢     无润滑  动摩擦系数  0.1，    有润滑  动摩擦系数  0.05-0.1
        ' 故这里取值为  0.08
        '***************************************************************************************************************
        If Val(Text63.Text) <= 0.01 Or Val(Text63.Text) > 0.3 Then
            msg_prompt = "请输入正确的库伦摩擦系数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "计算" And Combo6.Text = "计算" Then
            msg_prompt = "井口、井底环空压力总得输入一个。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "输入" And Val(Text3.Text) = 0 Then
            msg_prompt = "请输入正确的井底环空压力。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "计算" And (Val(Text58.Text) = 0 Or Val(Text60.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体参数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo5.Text = "计算" And Combo2.Text = "计算" Then
            msg_prompt = "井口、井底管内压力总得输入一个。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text = "输入" And Val(Text4.Text) = 0 Then
            msg_prompt = "请输入正确的井底管内压力。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text = "计算" And (Val(Text66.Text) = 0 Or Val(Text68.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体参数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Combo3.Text = "不流动" And Trim(Combo7.Text) <> "无" Then
            msg_prompt = "不流动的环空流体不用计算流动摩阻，请正确选择流体摩阻计算模型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And (Val(Text59.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体流量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Combo3.Text <> "不流动" And Combo7.Text = "牛顿流体模型" And (Val(Text60.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体沾度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And Combo7.Text = "降阻比模型" And (Val(Text6.Text) = 0) Then
            msg_prompt = "请输入正确的环空稠化剂浓度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        'If Combo3.text <> "不流动" And Combo7.text = "降阻比模型" And (Text10.text = 0) Then
        '     msg_prompt = "请输入正确的环空支撑剂浓度。"
        '     msg_buttons = 0 + 48
        '     msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        '     Exit Sub
        'End If
        If Combo3.Text <> "不流动" And Combo7.Text = "幂律流体模型" And (Val(Text11.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体流变指数n。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And Combo7.Text = "幂律流体模型" And (Val(Text12.Text) = 0) Then
            msg_prompt = "请输入正确的环空流体稠度系数K。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Combo4.Text = "不流动" And Combo8.Text <> "无" Then
            msg_prompt = "不流动的管内流体不用计算流动摩阻，请正确选择流体摩阻计算模型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Combo4.Text <> "不流动" And (Val(Text67.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体流量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo4.Text <> "不流动" And Combo8.Text = "牛顿流体模型" And (Val(Text68.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体沾度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo4.Text <> "不流动" And Combo8.Text = "降阻比模型" And (Val(Text13.Text) = 0) Then
            msg_prompt = "请输入正确的管内稠化剂浓度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        'If Combo4.text <> "不流动" And Combo8.text = "降阻比模型" And (Text14.text = 0) Then
        '     msg_prompt = "请输入正确的管内支撑剂浓度。"
        '     msg_buttons = 0 + 48
        '     msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        '     Exit Sub
        'End If
        If Combo4.Text <> "不流动" And Combo8.Text = "幂律流体模型" And (Val(Text15.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体流变指数n。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo3.Text <> "不流动" And Combo8.Text = "幂律流体模型" And (Val(Text16.Text) = 0) Then
            msg_prompt = "请输入正确的管内流体稠度系数K。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        msg_prompt = "是否要保存对所选工况数据的修改？"
        msg_buttons = 4 + 32
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        If msg_return <> 6 Then
            Exit Sub
        End If
        If Not Text61.Text = BindingSource1.Current("工况名称").ToString Then
            rowfilter = "工况序号 = " & CStr(Text53.Text)
            foundRows = dst.Tables("gk_xhmc_Table").Select(rowfilter)
            If foundRows.Length <> 0 Then
                foundRows(0).Item("工况名称") = Text61.Text
            End If
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "update 工况参数表 set " & "工况名称='" & Text61.Text & "',井口温度℃=" & CStr(Text54.Text) & "," & "井底温度℃=" & CStr(Text7.Text) & ",管压井底深度m=" & CStr(Text9.Text) & "," _
            & "井口环压MPa=" & CStr(Text56.Text) & ",环液深度m=" & CStr(Text57.Text) & "," & "环液密度g╱cm3=" & CStr(Text58.Text) & ",环液粘度mPaS=" & CStr(Text60.Text) & "," _
            & "环流流量m3╱m=" & CStr(Text59.Text) & ",井口管压MPa=" & CStr(Text64.Text) & "," & "管液深度m=" & CStr(Text65.Text) & ",管液密度g╱cm3=" & CStr(Text66.Text) & "," _
            & "管流粘度mPaS=" & CStr(Text68.Text) & ",管流流量m3╱m=" & CStr(Text67.Text) & "," & "库摩系数=" & CStr(Text63.Text) & ",井口加力kN='" & Text26.Text & "'," _
            & "环空流体流向='" & Combo3.Text & "',管内流体流向='" & Combo4.Text & "'," & "井口加扭Nm=" & CStr(Text8.Text) & ",井底套压MPa=" & CStr(Text3.Text) & "," _
            & "井底管压MPa=" & CStr(Text4.Text) & ",井底套压开关='" & Combo1.Text & "'," & "井底管压开关='" & Combo2.Text & "',套压井底深度m=" & CStr(Text5.Text) & "," _
            & "井口管压开关='" & Combo5.Text & "',井口套压开关='" & Combo6.Text & "'," & "环空流阻模型='" & Combo7.Text & "',管内流阻模型='" & Combo8.Text & "'," _
            & "环空稠剂浓度=" & CStr(Text6.Text) & ",管内稠剂浓度=" & CStr(Text13.Text) & "," & "环空撑剂浓度=" & CStr(Text10.Text) & ",管内撑剂浓度=" & CStr(Text14.Text) & "," _
            & "环空流变指数n=" & CStr(Text11.Text) & ",管内流变指数n=" & CStr(Text15.Text) & "," & "环空稠度系数K=" & CStr(Text12.Text) & ",管内稠度系数K=" & CStr(Text16.Text) & "," _
            & "环牛模折减系数=" & CStr(Text18.Text) & ",管牛模折减系数=" & CStr(Text19.Text) _
            & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(Text53.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
        msg_prompt = "    工况参数数据保存完成，请点击[计算]按钮重新进行管柱力学分析计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    工况参数保存出错，请输入正确的工况参数，不需要输入的参数，请不要空，输入0即可！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '    第2页"按节点显示"页填充：
    '(1)第2页"按节点显示"页待选节点列表                                          原来第2"按节点显示"页待选节点列表用Adodc5和DataGrid5
    '*********************************************************************************************************************************************
    Private Sub page2_dxjdlb_fill()
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        If Check2.CheckState = 0 Then
            SQL_command = "select distinct 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m,油管钢级,油管外径mm,油管内径mm,屈服强度MPa,节点情况表.抗挤强度MPa as 纯外压抗挤强度MPa,管体抗内压MPa,接头抗内压MPa,管体抗拉kN,接头抗拉kN " _
                & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                & " and 节点性质<>'计算点' and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                & " order by 节点计算参数表.节点编号"
        Else
            SQL_command = "select distinct 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m,油管钢级,油管外径mm,油管内径mm,屈服强度MPa,节点情况表.抗挤强度MPa as 纯外压抗挤强度MPa,管体抗内压MPa,接头抗内压MPa,管体抗拉kN,接头抗拉kN " _
            & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
            & " and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
            & " order by 节点计算参数表.节点编号"
        End If

        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        pg2_dxjdlb_Table.Clear()
        ad.Fill(pg2_dxjdlb_Table)
        ad.Dispose()
        BindingSource8.DataSource = pg2_dxjdlb_Table
        DataGridView9.ClearSelection()
        DataGridView9.DataSource = BindingSource8
        DataGridView9.ResetBindings()
        DataGridView9.AutoGenerateColumns = True
        DataGridView9.AllowUserToAddRows = False
        DataGridView9.AllowUserToDeleteRows = False
        DataGridView9.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView9.MultiSelect = False
        DataGridView9.RowHeadersWidth = 24
        DataGridView9.ReadOnly = True
        DataGridView9.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView9.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView9.Columns(0).Width = 35
        DataGridView9.Columns(1).Width = 55
        DataGridView9.Columns(2).Width = 70
        DataGridView9.Columns(3).Width = 95
        DataGridView9.Refresh()
        DataGridView9.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    按节点显示页读取待选节点列表数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '    第2页“按节点显示”页待选节点列表全部节点与关键节点复选框切换
    '*********************************************************************************************************************************************
    Private Sub Check2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check2.CheckedChanged
        Call page2_dxjdlb_fill()
    End Sub
    '*********************************************************************************************************************************************
    '    点击第2页“按节点显示”页待选节点列表,刷新本页各DataGridView,绘图
    '*********************************************************************************************************************************************
    Private Sub DataGridView9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView9.Click
        If IsNothing(Me.BindingSource8.Current) Then
            msg_prompt = "没有节点数据可显示，请点击[计算]按钮进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            Call page2_DataGridView_flash()
            If BindingSource8.Current("节点性质").ToString = "油井管" Or BindingSource8.Current("节点性质").ToString = "计算点" Or BindingSource8.Current("节点性质").ToString = "普通钻杆" Then
                Call draw_AFEDP_curve()
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '    第2页“按节点显示”页待选节点列表选择变化时,刷新本页各DataGridView,绘图
    '*********************************************************************************************************************************************
    'Private Sub DataGridView9_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView9.SelectionChanged
    '    If IsNothing(Me.BindingSource8.Current) Then
    '        msg_prompt = "没有节点数据可显示，请点击[计算]按钮进行计算。"
    '        msg_buttons = 0 + 48
    '        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    '    Else
    '        Call page2_DataGridView_flash()
    '        Call draw_AFEDP_curve()
    '    End If
    'End Sub

    '*********************************************************************************************************************************************
    '第2-2页“按节点显示-按工况计算结果”页,选定节点后刷新以下内容：                          原来Private Sub adodc8_flash()的功能
    '(1)待选工况列表                                              原来待选工况列表用Adodc7和DataGrid7
    '(2)2-2页选定节点不同工况计算参数列表                         原来2-2页选定节点不同工况计算参数列表Adodc8和DataGrid8
    '(3)2-1页选定节点不同工况计算参数列表                         原来2-1页选定节点不同工况计算参数列表Adodc6和DataGrid6
    '*********************************************************************************************************************************************
    Private Sub page2_DataGridView_flash()
        Dim jd_bh_str As String
        Dim i As Integer
        Dim seach_str1 As String
        Dim seach_str2 As String
        Dim seach_str3 As String
        Dim seach_str4 As String
        On Error GoTo errhandler
        i = 0
        jd_bh_str = "0"
        seach_str1 = ""
        seach_str2 = ""
        seach_str3 = ""
        If (Not IsNothing(Me.BindingSource8.Current)) Then
            If (Not IsDBNull(Me.BindingSource8.Current("节点编号"))) Then
                jd_bh_str = IIf(BindingSource8.Current("节点编号").ToString = "", "0", BindingSource8.Current("节点编号").ToString)
            End If
        End If
        seach_str1 = "select 节点计算参数表.工况序号,工况名称,"
        seach_str2 = " 管内压力MPa,管外压力MPa,真实轴力N,等效轴力N ,接触力N, 合弯矩Nm, 节点温度℃, 温度变形m, 轴力变形m, 鼓胀变形m, 螺旋变形m, 综合变形m," _
                   & " 温度效应m, 轴力效应m, 鼓胀效应m, 螺旋效应m, 综合效应m,合成应力MPa, 安全系数,等效悬持力N,屈服强度MPa," _
                   & " 套管内径mm,油管钢级,油管外径mm, 油管内径mm,Fecrh_N,扭矩Nm,"
        seach_str3 = " 环空流体流向,管内流体流向,井口温度℃,井底温度℃,管压井底深度m,套压井底深度m,井口环压MPa,环液深度m," _
                   & " 节点计算参数表.环液密度g╱cm3,环液粘度mPaS,环流流量m3╱m,井口管压MPa,管液深度m,节点计算参数表.管液密度g╱cm3,管流粘度mPaS,管流流量m3╱m,库摩系数,井口加力kN," _
                   & " 井口加扭Nm,节点计算参数表.抗挤强度MPa,管体抗内压MPa,接头抗内压MPa,管体抗拉kN,接头抗拉kN"
        seach_str4 = " from 工况参数表,节点情况表,节点计算参数表 where " _
                   & " 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                   & " and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                   & " and 工况参数表.井号=节点计算参数表.井号 and 工况参数表.作业名称=节点计算参数表.作业名称" _
                   & " and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点计算参数表.节点编号=" & jd_bh_str _
                   & " and 节点计算参数表.工况序号=工况参数表.工况序号" & " order by 节点计算参数表.工况序号"
        SQL_command = seach_str1 & seach_str2 & seach_str3 & seach_str4
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        jd_gk_jscs_Table.Clear()
        ad.Fill(jd_gk_jscs_Table)
        ad.Dispose()
        BindingSource9.DataSource = jd_gk_jscs_Table
        '2-2页选定节点不同工况计算参数列表                              原来选定节点不同工况计算参数列表Adodc8和DataGrid8
        DataGridView11.ClearSelection()
        DataGridView11.DataSource = BindingSource9
        DataGridView11.ResetBindings()
        DataGridView11.AutoGenerateColumns = True
        DataGridView11.AllowUserToAddRows = False
        DataGridView11.AllowUserToDeleteRows = False
        DataGridView11.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView11.MultiSelect = False
        DataGridView11.RowHeadersWidth = 24
        DataGridView11.ReadOnly = True
        DataGridView11.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView11.Refresh()
        DataGridView11.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取【按节点显示】页选定节点在不同工况下计算参数时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************
    '第2页图形初始化
    '*********************************************************************************************
    Private Sub draw_init3()
        Dim Channel_count As Integer
        Dim XAxis_count As Integer
        Dim YAxis_count As Integer
        Dim xTitle As String
        Dim yTitle As String
        Dim j As Short
        Dim titletext() As String
        Dim gk_row As DataRow
        Dim i As Short
        Dim num_of_gk As Integer
        On Error GoTo errhandler
        '初始化图形
        Channel_count = 0
        XAxis_count = 0
        YAxis_count = 0
        xTitle = "轴力(kN)"
        'If RadioButton3.Checked = True Then
        '    xTitle = "真实轴力(kN)"
        'End If
        'If RadioButton4.Checked = True Then
        '    xTitle = "等效轴力(kN)"
        'End If
        yTitle = "内外压差(MPa)"
        P2iPlotX1.get_Labels(1).Caption = "安全系数：抗外挤=" & Trim(TextBox7.Text) & "，抗内压=" & Trim(TextBox8.Text) & "，抗拉=" & Trim(TextBox9.Text) & "，三轴=" & Trim(Text20.Text)
        P2iPlotX1.ClearAllData()
        P2iPlotX1.RemoveAllChannels()
        P2iPlotX1.get_ToolBar(0).ShowEditButton = True
        '建立两个强度极限dChannel
        j = P2iPlotX1.AddChannel()
        P2iPlotX1.get_Channel(j).TitleText = "单双轴强度极限"
        P2iPlotX1.get_Channel(j).Visible = True
        P2iPlotX1.get_Channel(j).VisibleInLegend = True
        P2iPlotX1.get_Channel(j).TraceLineWidth = 3
        P2iPlotX1.get_Channel(j).Color = GetUint32color(1)
        j = P2iPlotX1.AddChannel()
        P2iPlotX1.get_Channel(j).TitleText = "三轴强度极限"
        P2iPlotX1.get_Channel(j).Visible = True
        P2iPlotX1.get_Channel(j).VisibleInLegend = True
        P2iPlotX1.get_Channel(j).TraceLineWidth = 3
        P2iPlotX1.get_Channel(j).Color = GetUint32color(1)
        'iPlotX1.TitleVisible = True
        'iPlotX1.TitleText = "请选择管柱节点/段"
        '查有多少工况，建立各工况Channel
        If dst.Tables("gk_xhmc_Table").Rows.Count <> 0 Then
            num_of_gk = dst.Tables("gk_xhmc_Table").Rows.Count
            ReDim titletext(num_of_gk)
            For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                j = P2iPlotX1.AddChannel()
                P2iPlotX1.get_Channel(j).TitleText = gk_row.Item("工况序号").ToString & "-" & gk_row.Item("工况名称").ToString
                P2iPlotX1.get_Channel(j).TraceLineWidth = 3
                P2iPlotX1.get_Channel(j).Color = GetUint32color(j)
            Next
        End If
        '设置X、Y轴
        P2iPlotX1.get_XAxis(XAxis_count).Title = xTitle
        P2iPlotX1.get_YAxis(YAxis_count).Title = yTitle
        P2iPlotX1.get_XAxis(XAxis_count).Min = -4000
        P2iPlotX1.get_XAxis(XAxis_count).Span = 8000
        P2iPlotX1.get_XAxis(XAxis_count).DesiredIncrement = 0.001
        P2iPlotX1.get_YAxis(YAxis_count).Min = -200
        P2iPlotX1.get_YAxis(YAxis_count).Span = 400
        P2iPlotX1.get_YAxis(YAxis_count).DesiredIncrement = 0.001
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    按节点显示页图像初始化读取节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************
    '第2页绘图
    '*********************************************************************************************
    Private Sub draw_AFEDP_curve()
        Dim burst_lim As Double
        Dim tension_lim As Double
        Dim collapse_lim As Double
        Dim compression_lim As Double
        Dim i As Integer
        Dim j As Integer
        Dim ireccount As Integer
        Dim waijing As Double
        Dim neijing As Double
        Dim Fz As Double
        Dim xgm_s As Double          '根据钢级确定的屈服强度，油套管使用手册上没有。
        Dim ped As Double            '在轴向应力和内压作用下的组合加载当量等级、等效屈服强度
        Dim jd_gk_jscs_row As DataRow
        Dim gz_row() As DataRow
        Dim gangji As String
        Dim yjxzh As String

        On Error GoTo errhandler
        i = 0
        If IsNothing(Me.BindingSource9.Current) Then
            msg_prompt = "没有节点数据可显示，请点击[计算]按钮进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        P2iPlotX1.get_Labels(1).Caption = "安全系数：抗外挤=" & Trim(TextBox7.Text) & "，抗内压=" & Trim(TextBox8.Text) & "，抗拉=" & Trim(TextBox9.Text) & "，三轴=" & Trim(Text20.Text)
        '*********************************************************************************************
        ' 1.极限强度赋值
        '*********************************************************************************************
        '抗内压强度取管体抗内压MPa、接头抗内压MPa 中的小值
        burst_lim = Val(BindingSource8.Current("管体抗内压MPa").ToString)
        If burst_lim > Val(BindingSource8.Current("接头抗内压MPa").ToString) Then
            burst_lim = Val(BindingSource8.Current("接头抗内压MPa").ToString)
        End If
        If Val(TextBox8.Text) <> 0 And Val(TextBox8.Text) >= 1 Then
            burst_lim = burst_lim / Val(TextBox8.Text)
        End If
        '抗拉强度取管体抗拉kN、接头抗拉kN 中的小值
        tension_lim = Val(BindingSource8.Current("管体抗拉kN").ToString)
        If tension_lim > Val(BindingSource8.Current("接头抗拉kN").ToString) Then
            tension_lim = Val(BindingSource8.Current("接头抗拉kN").ToString)
        End If
        If Val(TextBox9.Text) <> 0 And Val(TextBox9.Text) >= 1 Then
            tension_lim = tension_lim / Val(TextBox9.Text)
        End If
        '抗挤强度MPa
        collapse_lim = (-1) * Val(BindingSource8.Current("纯外压抗挤强度MPa").ToString)
        If Val(TextBox7.Text) <> 0 And Val(TextBox7.Text) >= 1 Then
            collapse_lim = collapse_lim / Val(TextBox7.Text)
        End If
        '抗压挤强度取(-1) *抗拉强度
        compression_lim = (-1) * tension_lim
        '屈服强度
        xgm_s = Val(BindingSource8.Current("屈服强度MPa").ToString)
        If Val(Text20.Text) <> 0 And Val(Text20.Text) >= 1 Then
            xgm_s = xgm_s / Val(Text20.Text)
        End If
        '油管钢级,
        gangji = BindingSource8.Current("油管钢级").ToString
        waijing = Val(BindingSource8.Current("油管外径mm").ToString)
        neijing = Val(BindingSource8.Current("油管内径mm").ToString)
        '  节点列表中的节点ID，即sp_gzh_Table中的序号（从管柱表中选取时将“元件序号”as 为“序号”，将“元件性质”as 为“类型”）,
        '利用此关系得到管柱表中的元件性质                                  20231210 秦彦斌注
        yjxzh = ""
        gz_row = sp_gzh_Table.Select("序号 =" & BindingSource8.Current("节点ID").ToString)
        If gz_row.Length > 0 Then
            yjxzh = Trim(gz_row(0).Item("类型"))
        End If
        '*********************************************************************************************
        ' 2.画强度极限包络图
        '*********************************************************************************************
        If RadioButton3.Checked = True Then
            Call draw_qdjx_blqx(P2iPlotX1, burst_lim, tension_lim, collapse_lim, compression_lim, xgm_s, waijing, neijing, 0)
            'P2iPlotX1.get_XAxis(0).Title = "真实轴力(kN)"
        End If
        If RadioButton4.Checked = True Then
            Call draw_qdjx_blqx(P2iPlotX1, burst_lim, tension_lim, collapse_lim, compression_lim, xgm_s, waijing, neijing, 1)
            'P2iPlotX1.get_XAxis(0).Title = "等效轴力(kN)"
        End If
        '*********************************************************************************************
        '3.1 当前节点对应管柱在各工况下轴力-压差曲线绘制
        '*********************************************************************************************
        If RadioButton2.Checked = True Then   '选中管柱段画线
            '临时借用，壁厚
            Fz = Fix(0.5 * (waijing - neijing) * 10000 + 0.9) / 10000
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            P2iPlotX1.TitleText = "Φ" & Trim(Str(waijing)) & "×" & Trim(Str(Fz)) & "mm" & Trim(gangji) & yjxzh & "安全性包络图"
            ireccount = 0
            For Each jd_gk_jscs_row In dst.Tables("jd_gk_jscs_Table").Select
                SQL_command = "select 节点情况表.节点编号,节点情况表.节点ID,管内压力MPa,管外压力MPa,等效轴力N,真实轴力N" _
                           & " from 节点情况表,节点计算参数表 where " _
                           & " 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                           & " and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                           & " and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.油管钢级='" & gangji & "' " _
                           & " and 节点情况表.油管外径mm=" & CStr(waijing) & " and 节点情况表.油管内径mm=" & CStr(neijing) _
                           & " and 节点计算参数表.工况序号=" & jd_gk_jscs_row.Item("工况序号").ToString _
                           & " order by 节点情况表.节点编号"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                EXECOleDbCommand.Dispose()
                If RECreader.HasRows Then
                    i = 1 + Val(jd_gk_jscs_row.Item("工况序号").ToString)
                    P2iPlotX1.get_Channel(i).Clear()
                    P2iPlotX1.get_Channel(i).TraceLineWidth = 3
                    j = 0
                    While RECreader.Read()
                        '不连续的节点，增加“断点”连线，避免误会。
                        If j <> 0 Then
                            If (Val(RECreader.Item("节点编号").ToString()) - ireccount) <> 1 Then
                                P2iPlotX1.get_Channel(i).AddXNull(0)
                            End If
                        End If
                        ireccount = Val(RECreader.Item("节点编号").ToString())
                        j = j + 1
                        ped = Val(RECreader.Item("管内压力MPa").ToString()) - Val(RECreader.Item("管外压力MPa").ToString())
                        Fz = Val(RECreader.Item("真实轴力N").ToString) / 1000.0
                        If RadioButton3.Checked = True Then
                            Fz = Val(RECreader.Item("真实轴力N").ToString) / 1000.0
                        End If
                        If RadioButton4.Checked = True Then
                            Fz = Val(RECreader.Item("等效轴力N").ToString) / 1000.0
                        End If
                        P2iPlotX1.get_Channel(i).AddXY(Fz, ped)
                    End While
                End If
                RECreader.Close()
            Next
            cn_userdb.Close()
            cn_userdb.Dispose()
        End If
        '*********************************************************************************************
        ''3.2 当前节点各工况画点
        '*********************************************************************************************
        If RadioButton1.Checked = True Then   '选中管柱节点点点
            '临时借用，壁厚
            Fz = Fix(0.5 * (waijing - neijing) * 10000 + 0.9) / 10000
            P2iPlotX1.TitleText = "Φ" & Trim(Str(waijing)) & "×" & Trim(Str(Fz)) & "mm" & Trim(gangji) & yjxzh & Str(Fix(Val(BindingSource8.Current("节点下深m").ToString) * 1000) / 1000) & "m处安全性包络图"
            For Each jd_gk_jscs_row In jd_gk_jscs_Table.Rows
                i = 1 + Val(jd_gk_jscs_row.Item("工况序号").ToString)
                ped = Val(jd_gk_jscs_row.Item("管内压力MPa").ToString) - Val(jd_gk_jscs_row.Item("管外压力MPa").ToString)
                Fz = Val(jd_gk_jscs_row.Item("真实轴力N").ToString) / 1000.0
                If RadioButton3.Checked = True Then
                    Fz = Val(jd_gk_jscs_row.Item("真实轴力N").ToString) / 1000.0
                End If
                If RadioButton4.Checked = True Then
                    Fz = Val(jd_gk_jscs_row.Item("等效轴力N").ToString) / 1000.0
                End If
                P2iPlotX1.get_Channel(i).Clear()
                Call drawDot_at_ch(P2iPlotX1, i, Fz, ped)
            Next
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    按节点显示页绘制安全性包络图读取节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 点击按节点显示-按工况计算结果页中的选定节点在不同工况下的计算参数列表
    '*********************************************************************************************************************************************
    Private Sub DataGridView11_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView11.Click
        Call jd_gk_change()
    End Sub
    '*********************************************************************************************************************************************
    ' 按节点显示-按工况计算结果页中的选定节点在不同工况下的计算参数列表选择变化时
    '*********************************************************************************************************************************************
    'Private Sub DataGridView11_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView11.SelectionChanged
    '    Call jd_gk_change()
    'End Sub

    '*********************************************************************************************************************************************
    ' 按节点显示-按工况计算结果页中，工况变化时，将计算参数赋值给各文本框以刷新显示
    '*********************************************************************************************************************************************
    Private Sub jd_gk_change()
        On Error GoTo errhandler
        If (Not IsNothing(Me.BindingSource9.Current)) Then
            If (Not IsDBNull(Me.BindingSource9.Current("工况序号"))) Then
                _Text2_47.Text = IIf(BindingSource9.Current("套管内径mm").ToString = "", "0", BindingSource9.Current("套管内径mm").ToString)
                _Text2_46.Text = BindingSource9.Current("油管钢级").ToString
                _Text2_45.Text = IIf(BindingSource9.Current("油管外径mm").ToString = "", "0", BindingSource9.Current("油管外径mm").ToString)
                _Text2_44.Text = IIf(BindingSource9.Current("油管内径mm").ToString = "", "0", BindingSource9.Current("油管内径mm").ToString)
                _Text2_43.Text = IIf(BindingSource9.Current("屈服强度MPa").ToString = "", "0", BindingSource9.Current("屈服强度MPa").ToString)
                _Text2_42.Text = IIf(BindingSource9.Current("节点温度℃").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("节点温度℃").ToString)) / 1000))
                _Text2_41.Text = IIf(BindingSource9.Current("管外压力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("管外压力MPa").ToString)) / 1000))
                _Text2_40.Text = IIf(BindingSource9.Current("管内压力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("管内压力MPa").ToString)) / 1000))
                _Text2_39.Text = IIf(BindingSource9.Current("真实轴力N").ToString = "", "0", CStr(Fix(Val(BindingSource9.Current("真实轴力N").ToString)) / 1000))
                _Text2_38.Text = IIf(BindingSource9.Current("接触力N").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("接触力N").ToString)) / 1000))
                _Text2_37.Text = IIf(BindingSource9.Current("等效轴力N").ToString = "", "0", CStr(Fix(Val(BindingSource9.Current("等效轴力N").ToString)) / 1000))
                _Text2_36.Text = IIf(BindingSource9.Current("合弯矩Nm").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("合弯矩Nm").ToString)) / 1000))
                _Text2_35.Text = IIf(BindingSource9.Current("合成应力MPa").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("合成应力MPa").ToString)) / 1000))
                _Text2_34.Text = IIf(BindingSource9.Current("安全系数").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("安全系数").ToString)) / 1000))
                _Text2_33.Text = IIf(BindingSource9.Current("温度变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("温度变形m").ToString)) / 1000))
                _Text2_32.Text = IIf(BindingSource9.Current("温度效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("温度效应m").ToString)) / 1000))
                _Text2_31.Text = IIf(BindingSource9.Current("轴力变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("轴力变形m").ToString)) / 1000))
                _Text2_30.Text = IIf(BindingSource9.Current("轴力效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("轴力效应m").ToString)) / 1000))
                _Text2_29.Text = IIf(BindingSource9.Current("鼓胀变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("鼓胀变形m").ToString)) / 1000))
                _Text2_28.Text = IIf(BindingSource9.Current("鼓胀效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("鼓胀效应m").ToString)) / 1000))
                _Text2_27.Text = IIf(BindingSource9.Current("螺旋变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("螺旋变形m").ToString)) / 1000))
                _Text2_26.Text = IIf(BindingSource9.Current("螺旋效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("螺旋效应m").ToString)) / 1000))
                _Text2_25.Text = IIf(BindingSource9.Current("综合变形m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("综合变形m").ToString)) / 1000))
                _Text2_24.Text = IIf(BindingSource9.Current("综合效应m").ToString = "", "0", CStr(Fix(1000 * Val(BindingSource9.Current("综合效应m").ToString)) / 1000))
                TextBox1.Text = IIf(BindingSource9.Current("抗挤强度MPa").ToString = "", "0", BindingSource9.Current("抗挤强度MPa").ToString)
                TextBox2.Text = IIf(BindingSource9.Current("管体抗内压MPa").ToString = "", "0", BindingSource9.Current("管体抗内压MPa").ToString)
                TextBox3.Text = IIf(BindingSource9.Current("接头抗内压MPa").ToString = "", "0", BindingSource9.Current("接头抗内压MPa").ToString)
                TextBox4.Text = IIf(BindingSource9.Current("管体抗拉kN").ToString = "", "0", BindingSource9.Current("管体抗拉kN").ToString)
                TextBox5.Text = IIf(BindingSource9.Current("接头抗拉kN").ToString = "", "0", BindingSource9.Current("接头抗拉kN").ToString)
                TextBox6.Text = CStr(Fix(1000000 * (Val(BindingSource9.Current("管内压力MPa").ToString) - Val(BindingSource9.Current("管外压力MPa").ToString))) / 1000000)
                'Call draw_AFEDP_curve()
            End If
        Else
            msg_prompt = "没有结果数据可显示，请点击[计算]按钮进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    读取选定节点在不同工况下计算结果出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 在按节点显示页中，点击“添加”按钮
    '*********************************************************************************************************************************************
    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
        Dim foundRows() As DataRow
        Dim tw_newRow As DataRow
        Dim rowfilter As String
        On Error GoTo errhandler
        If IsNothing(Me.BindingSource8.Current) Then
            msg_prompt = "没有节点数据可添加，请点击[计算]按钮进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            rowfilter = "节点编号 = " & BindingSource8.Current("节点编号").ToString
            foundRows = dst.Tables("pg2_prtjdlb_Table").Select(rowfilter)
            If foundRows.Length <> 0 Then
                msg_prompt = "该节点已添加。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                tw_newRow = dst.Tables("pg2_prtjdlb_Table").NewRow()
                tw_newRow.Item("节点编号") = CType(BindingSource8.Current("节点编号").ToString, Integer)
                tw_newRow.Item("节点ID") = BindingSource8.Current("节点ID").ToString
                tw_newRow.Item("节点性质") = BindingSource8.Current("节点性质").ToString
                tw_newRow.Item("节点下深m") = CType(BindingSource8.Current("节点下深m").ToString, Double)
                dst.Tables("pg2_prtjdlb_Table").Rows.Add(tw_newRow)
                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "insert into  lsb_repoutlist (节点编号,节点ID,节点性质,节点下深m) values (" _
                    & BindingSource8.Current("节点编号").ToString & ",'" _
                    & BindingSource8.Current("节点ID").ToString & "','" _
                    & BindingSource8.Current("节点性质").ToString & "'," _
                    & BindingSource8.Current("节点下深m").ToString & ")"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                cn_userdb.Close()
                cn_userdb.Dispose()
            End If
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "待打印节点添加出错，请确认数据文件未被其它软件打开。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 在按节点显示页中，点击“删除”按钮
    '*********************************************************************************************************************************************
    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
        On Error GoTo errhandler
        If IsNothing(Me.BindingSource10.Current) Then
            msg_prompt = "没有节点数据可删除。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "delete from lsb_repoutlist where 节点编号=" & BindingSource10.Current("节点编号").ToString
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            cn_userdb.Close()
            cn_userdb.Dispose()
            BindingSource10.Current().delete()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "待打印节点删除出错，请确认数据文件未被其它软件打开。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    ' 在按节点显示页中，点击“所选节点生成Word报告[&D]...”按钮
    '*********************************************************************************************************************************************
    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
        '***********************************************************************************************
        '*  变量说明及赋初值
        '***********************************************************************************************
        Dim j As Integer
        Dim i As Integer
        Dim k As Integer
        Dim ireccount As Integer
        Dim gk_count As Short
        Dim FileName As String
        Dim filepath As String
        Dim WordWP As New Microsoft.Office.Interop.Word.Application
        Dim WordBook As New Microsoft.Office.Interop.Word.Document
        Dim mytable As Microsoft.Office.Interop.Word.Table
        Dim gk_row As DataRow
        Dim RECreader2 As OleDbDataReader
        Dim lsstr As String
        On Error GoTo errhandler
        '***********************************************************************************************
        '*  待打印列表为空时给出提示
        '***********************************************************************************************
        If IsNothing(Me.BindingSource10.Current) Then
            msg_prompt = "    没有节点数据可输出，请选择节点。若没有可选节点，请点击[计算]按钮计算后再选择。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '*********************************************************************************************************************************************
        '向用户要文件名
        '*********************************************************************************************************************************************
        filepath = My.Application.Info.DirectoryPath & "\算例"
        SaveFileDialog1.InitialDirectory = filepath & "\"
        SaveFileDialog1.Filter = "Documents(*.DOC)|*.DOC"
        SaveFileDialog1.FileName = well_name & "井" & zuoye_name & "管柱力学分析报告-节点法"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.Title = "保存"
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            FileName = SaveFileDialog1.FileName
        Else
            Exit Sub
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        WordBook = WordWP.Documents.Add
        With WordWP
            With .Selection
                '***********************************************************************************************
                '*  1.写报告标题行
                '***********************************************************************************************
                ' .Font.Size = 9                  '小五号
                ' .Font.Size = 10.5               '五号
                ' .Font.Size = 12                 '小四号
                ' .Font.Size = 14                 '四号
                ' .Font.Size = 15                 '小三号
                ' .Font.Size = 16                 '三号
                ' .Font.Size = 18                 '小二号
                ' .Font.Size = 22                 '二号
                '***********************************************************************************************
                '.PageSetup.Orientation = wdOrientLandscape
                .Font.Size = 16
                .Font.Name = "宋体"
                .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .TypeText(Text:=well_name & "井" & zuoye_name & "管柱力学分析报告")
                .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                .TypeParagraph()
                .TypeParagraph()
                If dst.Tables("gk_xhmc_Table").Rows.Count <> 0 Then
                    gk_count = dst.Tables("gk_xhmc_Table").Rows.Count
                    '***********************************************************************************************
                    '*  2.写工况参数表-（表1）
                    '***********************************************************************************************
                    '***********************************************************************************************
                    '*  2.1 写工况数据标题行
                    '***********************************************************************************************
                    .Font.Size = 12
                    .Font.Name = "宋体"
                    .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .TypeText(Text:="表1  " & well_name & "井" & zuoye_name & "管柱工况参数")
                    .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  2.2 写工况参数表
                    '***********************************************************************************************
                    .Font.Size = 10.5
                    mytable = .Tables.Add(.Range, 13, gk_count + 1)
                    .Tables.Item(1).Select()
                    .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Tables.Item(1).Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter
                    .Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    With mytable.Borders
                        .InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .InsideLineStyle = True
                        '.OutsideLineStyle = wdLineStyleDouble
                        .OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                    End With
                    'mytable.Cell(1, 1).Range.InsertAfter "工况序号"
                    mytable.Cell(1, 1).Range.InsertAfter("工况名称")
                    mytable.Cell(2, 1).Range.InsertAfter("井口温度(℃)")
                    mytable.Cell(3, 1).Range.InsertAfter("井底温度(℃)")
                    mytable.Cell(4, 1).Range.InsertAfter("井口油压(MPa)")
                    mytable.Cell(5, 1).Range.InsertAfter("井口套压(MPa)")
                    mytable.Cell(6, 1).Range.InsertAfter("井口加力kN")
                    mytable.Cell(7, 1).Range.InsertAfter("井口加扭Nm")
                    mytable.Cell(8, 1).Range.InsertAfter("管内流量(m3/m)")
                    mytable.Cell(9, 1).Range.InsertAfter("管内流体密度(g/cm3)")
                    mytable.Cell(10, 1).Range.InsertAfter("管内流体粘度(mPa.s)")
                    mytable.Cell(11, 1).Range.InsertAfter("环空流量(m3/m)")
                    mytable.Cell(12, 1).Range.InsertAfter("环空流体密度(g/cm3)")
                    mytable.Cell(13, 1).Range.InsertAfter("环空流体粘度(mPa.s)")

                    'mytable.Cell(1, 7).Range.InsertAfter ("井底深度(m)")
                    'mytable.Cell(1, 8).Range.InsertAfter ("环空液面深度(m)")
                    'mytable.Cell(1, 9).Range.InsertAfter ("环空流体流向")
                    'mytable.Cell(1, 14).Range.InsertAfter ("管内液面深度(m)")
                    'mytable.Cell(1, 15).Range.InsertAfter ("管内流体流向")
                    'mytable.Cell(1, 20).Range.InsertAfter ("库摩系数")

                    SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    j = 2
                    While RECreader.Read
                        'mytable.Cell(j, 1).Range.InsertAfter RECreader.Item("工况序号"))
                        mytable.Cell(1, j).Range.InsertAfter(RECreader.Item("工况名称").ToString)
                        mytable.Cell(2, j).Range.InsertAfter(RECreader.Item("井口温度℃").ToString)
                        mytable.Cell(3, j).Range.InsertAfter(RECreader.Item("井底温度℃").ToString)
                        mytable.Cell(4, j).Range.InsertAfter(RECreader.Item("井口管压MPa").ToString)
                        mytable.Cell(5, j).Range.InsertAfter(RECreader.Item("井口环压MPa").ToString)
                        mytable.Cell(6, j).Range.InsertAfter(RECreader.Item("井口加力kN").ToString)
                        mytable.Cell(7, j).Range.InsertAfter(RECreader.Item("井口加扭Nm").ToString)
                        mytable.Cell(8, j).Range.InsertAfter(RECreader.Item("管流流量m3╱m").ToString)
                        mytable.Cell(9, j).Range.InsertAfter(RECreader.Item("管液密度g╱cm3").ToString)
                        mytable.Cell(10, j).Range.InsertAfter(RECreader.Item("管流粘度mPaS").ToString)
                        mytable.Cell(11, j).Range.InsertAfter(RECreader.Item("环流流量m3╱m").ToString)
                        mytable.Cell(12, j).Range.InsertAfter(RECreader.Item("环液密度g╱cm3").ToString)
                        mytable.Cell(13, j).Range.InsertAfter(RECreader.Item("环液粘度mPaS").ToString)

                        'mytable.Cell(j, 7).Range.InsertAfter (RECreader.Item("井底深度m").ToString)
                        'mytable.Cell(j, 8).Range.InsertAfter (RECreader.Item("环液深度m").ToString)
                        'mytable.Cell(j, 9).Range.InsertAfter (RECreader.Item("环空流体流向").ToString)
                        'mytable.Cell(j, 14).Range.InsertAfter (RECreader.Item("管液深度m").ToString)
                        'mytable.Cell(j, 15).Range.InsertAfter (RECreader.Item("管内流体流向").ToString)
                        'mytable.Cell(j, 20).Range.InsertAfter (RECreader.Item("库摩系数").ToString)
                        j = j + 1
                    End While
                    .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    .TypeParagraph()
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  3.写管柱重量及拉伸应力安全系数-（表2）
                    '***********************************************************************************************
                    '***********************************************************************************************
                    '*  4.写管柱载荷、应力及安全系数数值表-（表3）
                    '***********************************************************************************************
                    '***********************************************************************************************
                    '*  4.1 写管柱载荷、应力及安全系数数值表-（表3）标题行
                    '***********************************************************************************************
                    .Font.Size = 12
                    .Font.Name = "宋体"
                    .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .TypeText(Text:="表3  " & well_name & "井" & zuoye_name & "管柱载荷、应力及安全系数数值")
                    .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  4.2 设定管柱载荷、应力及安全系数数值表-（表3）行数、列数
                    '***********************************************************************************************
                    If dst.Tables("pg2_prtjdlb_Table").Rows.Count <> 0 Then
                        ireccount = dst.Tables("pg2_prtjdlb_Table").Rows.Count
                        .Font.Size = 10.5
                        mytable = .Tables.Add(.Range, ireccount * 6 + 1, gk_count + 2)
                        .Tables.Item(1).Select()
                        .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Tables.Item(1).Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter
                        .Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With mytable.Borders
                            .InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            .InsideLineStyle = True
                            '.OutsideLineStyle = wdLineStyleDouble
                            .OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End With
                        '***********************************************************************************************
                        '*  4.3 填写管柱载荷、应力及安全系数数值表-（表3）部分表头
                        '***********************************************************************************************
                        'mytable.Cell(1, 1).Range.InsertAfter "节点编号"
                        'mytable.Cell(1, 1).Range.InsertAfter "节点ID"
                        'mytable.Cell(1, 2).Range.InsertAfter "节点性质"
                        mytable.Cell(1, 1).Range.InsertAfter("位置(m)")
                        mytable.Cell(1, 2).Range.InsertAfter("参数名称")
                        '***********************************************************************************************
                        '*  显示进度条
                        '***********************************************************************************************
                        frmprogress.PB1.Maximum = gk_count + 1
                        frmprogress.PB1.Minimum = 0
                        frmprogress.Label1.Text = "正在进行导出..."
                        frmprogress.Show()
                        i = 1
                        For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                            mytable.Cell(1, i + 2).Range.InsertAfter(gk_row.Item("工况名称").ToString)
                            frmprogress.PB1.Value = i
                            '***********************************************************************************************
                            '*  读取节点并循环 .节点编号,节点ID,节点性质,节点情况表.节点下深m
                            '***********************************************************************************************
                            'SQL_command = "select 节点计算参数表.节点编号,节点ID,节点性质,lsb_repoutlist.节点下深m,管内压力MPa," _
                            '    & " 管外压力MPa,节点温度℃,真实轴力N,合成应力MPa, 安全系数" & " from 节点计算参数表 INNER JOIN lsb_repoutlist on  节点计算参数表.节点编号=lsb_repoutlist.节点编号" _
                            '    & " where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                            '    & " and 节点计算参数表.工况序号=" & gk_row.Item("工况序号").ToString & " order by 节点计算参数表.节点编号"
                            SQL_command = "select 节点计算参数表.节点编号,节点ID,节点性质,lsb_repoutlist.节点下深m,管内压力MPa," _
                                & " 管外压力MPa,节点温度℃,真实轴力N,等效轴力N,合成应力MPa, 安全系数" & " from 节点计算参数表 INNER JOIN lsb_repoutlist on  节点计算参数表.节点编号=lsb_repoutlist.节点编号" _
                                & " where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                                & " and 节点计算参数表.工况序号=" & gk_row.Item("工况序号").ToString & " order by 节点计算参数表.节点编号"
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                            RECreader = EXECOleDbCommand.ExecuteReader()
                            EXECOleDbCommand.Dispose()
                            j = 1
                            While RECreader.Read
                                '***********************************************************************************************
                                '*  写各节点计算参数数据
                                '***********************************************************************************************
                                If i = 1 Then
                                    mytable.Cell((j - 1) * 6 + 2, 1).Range.InsertAfter(CStr(Fix(100.0# * Val(RECreader.Item("节点下深m").ToString)) / 100.0#))
                                    mytable.Cell((j - 1) * 6 + 2, 2).Range.InsertAfter("内压(MPa)")
                                    mytable.Cell((j - 1) * 6 + 3, 2).Range.InsertAfter("外压(MPa)")
                                    mytable.Cell((j - 1) * 6 + 4, 2).Range.InsertAfter("等效轴力(kN)")       '20230401改为等效轴力
                                    mytable.Cell((j - 1) * 6 + 5, 2).Range.InsertAfter("真实轴力(kN)")    '20230401改为等效轴力
                                    mytable.Cell((j - 1) * 6 + 6, 2).Range.InsertAfter("σxd4(MPa)")
                                    mytable.Cell((j - 1) * 6 + 7, 2).Range.InsertAfter("安全系数")
                                End If
                                mytable.Cell((j - 1) * 6 + 2, 2 + i).Range.InsertAfter(CStr(Fix(1.0# * Val(RECreader.Item("管内压力MPa").ToString)) / 1.0#))
                                mytable.Cell((j - 1) * 6 + 3, 2 + i).Range.InsertAfter(CStr(Fix(1.0# * Val(RECreader.Item("管外压力MPa").ToString)) / 1.0#))
                                mytable.Cell((j - 1) * 6 + 4, 2 + i).Range.InsertAfter(CStr(Fix(Val(RECreader.Item("等效轴力N").ToString) / 1000) / 1.0#))      '20230401改为等效轴力
                                mytable.Cell((j - 1) * 6 + 5, 2 + i).Range.InsertAfter(CStr(Fix(Val(RECreader.Item("真实轴力N").ToString) / 1000) / 1.0#))       '20230401改为等效轴力
                                mytable.Cell((j - 1) * 6 + 6, 2 + i).Range.InsertAfter(CStr(Fix(1.0# * Val(RECreader.Item("合成应力MPa").ToString)) / 1.0#))
                                If Val(RECreader.Item("安全系数").ToString) = 5 Then
                                    lsstr = ">5"
                                Else
                                    lsstr = CStr(Fix(100.0# * Val(RECreader.Item("安全系数").ToString)) / 100.0#)
                                End If
                                mytable.Cell((j - 1) * 6 + 7, 2 + i).Range.InsertAfter(lsstr)
                                j = j + 1
                            End While
                            RECreader.Close()
                            i = i + 1
                        Next
                        .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    Else
                        msg_prompt = "    未选择报告输出的节点，无法输出管柱载荷、应力及安全系数数值表。"
                        msg_buttons = 0 + 48
                        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    End If
                    .TypeParagraph()
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  5.写管柱轴向变形和效应值（m）-（表4）
                    '***********************************************************************************************
                    '***********************************************************************************************
                    '*  5.1 写管柱轴向变形和效应值（m）-（表4）标题行
                    '***********************************************************************************************
                    .Font.Size = 12
                    .Font.Name = "宋体"
                    .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .TypeText(Text:="表4  " & well_name & "井" & zuoye_name & "管柱轴向变形和""效应""值（m）")
                    .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  5.2 设定管柱轴向变形和效应值（m）-（表4）行数、列数
                    '***********************************************************************************************
                    '找到封隔器前一管柱的节点编号
                    SQL_command = "select top 1 节点编号 from 节点情况表 " & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'" & " and 节点性质='封隔器' order by 节点编号"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    If RECreader.HasRows Then
                        RECreader.Read()
                        ireccount = Val(RECreader.Item("节点编号").ToString) - 1
                        SQL_command = "select distinct 节点类型,节点ID,节点性质,节点编号,节点下深m from 节点情况表 " & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name _
                            & "' and 节点编号= " & CStr(ireccount)
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                        EXECOleDbCommand.Dispose()
                        If RECreader2.HasRows Then
                            .Font.Size = 10.5
                            mytable = .Tables.Add(.Range, 11, gk_count + 2)
                            .Tables.Item(1).Select()
                            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Tables.Item(1).Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter
                            .Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With mytable.Borders
                                .InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                .InsideLineStyle = True
                                '.OutsideLineStyle = wdLineStyleDouble
                                .OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With
                            '***********************************************************************************************
                            '*  5.3 填写管柱轴向变形和效应值（m）-（表4）部分表头
                            '***********************************************************************************************
                            'mytable.Cell(1, 1).Range.InsertAfter "节点编号"
                            'mytable.Cell(1, 1).Range.InsertAfter "节点ID"
                            'mytable.Cell(1, 2).Range.InsertAfter "节点性质"
                            mytable.Cell(1, 1).Range.InsertAfter("位置(m)")
                            mytable.Cell(1, 2).Range.InsertAfter("参数名称")
                            '***********************************************************************************************
                            '*  显示进度条
                            '***********************************************************************************************
                            frmprogress.PB1.Maximum = gk_count + 1
                            frmprogress.PB1.Minimum = 0
                            frmprogress.Label1.Text = "正在进行导出..."
                            frmprogress.Show()
                            i = 1
                            For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                                mytable.Cell(1, i + 2).Range.InsertAfter(gk_row.Item("工况名称").ToString)
                                frmprogress.PB1.Value = i
                                '***********************************************************************************************
                                '*  读取节点并循环
                                '***********************************************************************************************
                                SQL_command = "select 节点编号,节点下深m,温度变形m, 轴力变形m, 鼓胀变形m, 螺旋变形m, 综合变形m,温度效应m, 轴力效应m, 鼓胀效应m, 螺旋效应m, 综合效应m " _
                                    & " from 节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                                    & " and 节点计算参数表.工况序号=" & gk_row.Item("工况序号").ToString & " and 节点计算参数表.节点编号=" & CStr(ireccount)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader2 = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If RECreader2.HasRows Then
                                    j = 1
                                    While RECreader2.Read
                                        '***********************************************************************************************
                                        '*  写各节点计算参数数据
                                        '***********************************************************************************************
                                        If i = 1 Then
                                            'mytable.Cell((j - 1) * 5 + 2, 1).Range.InsertAfter CStr(SQL_rst_jd.Fields("节点编号"))
                                            'mytable.Cell((j - 1) * 5 + 2, 1).Range.InsertAfter CStr(SQL_rst_jd.Fields("节点ID"))
                                            'mytable.Cell((j - 1) * 5 + 2, 2).Range.InsertAfter CStr(SQL_rst_jd.Fields("节点性质"))
                                            mytable.Cell((j - 1) * 10 + 2, 1).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("节点下深m").ToString)) / 1000.0#))
                                            mytable.Cell((j - 1) * 10 + 2, 2).Range.InsertAfter("温度变形")
                                            mytable.Cell((j - 1) * 10 + 3, 2).Range.InsertAfter("温度效应")
                                            mytable.Cell((j - 1) * 10 + 4, 2).Range.InsertAfter("轴力变形")
                                            mytable.Cell((j - 1) * 10 + 5, 2).Range.InsertAfter("轴力效应")
                                            mytable.Cell((j - 1) * 10 + 6, 2).Range.InsertAfter("鼓胀变形")
                                            mytable.Cell((j - 1) * 10 + 7, 2).Range.InsertAfter("鼓胀效应")
                                            mytable.Cell((j - 1) * 10 + 8, 2).Range.InsertAfter("螺旋变形")
                                            mytable.Cell((j - 1) * 10 + 9, 2).Range.InsertAfter("螺旋效应")
                                            mytable.Cell((j - 1) * 10 + 10, 2).Range.InsertAfter("综合变形")
                                            mytable.Cell((j - 1) * 10 + 11, 2).Range.InsertAfter("综合效应")
                                        End If
                                        mytable.Cell((j - 1) * 10 + 2, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("温度变形m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 3, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("温度效应m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 4, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("轴力变形m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 5, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("轴力效应m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 6, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("鼓胀变形m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 7, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("鼓胀效应m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 8, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("螺旋变形m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 9, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("螺旋效应m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 10, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("综合变形m").ToString)) / 1000.0#))
                                        mytable.Cell((j - 1) * 10 + 11, 2 + i).Range.InsertAfter(CStr(Fix(1000.0# * Val(RECreader2.Item("综合效应m").ToString)) / 1000.0#))
                                        j = j + 1
                                    End While
                                End If
                                RECreader2.Close()
                                i = i + 1
                            Next
                            .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        Else
                            msg_prompt = "    在封隔器前找不到节点数据，无法输出管柱轴向变形和效应值表。"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        End If
                    Else
                        msg_prompt = "    找不到封隔器，无法输出管柱轴向变形和效应值表。"
                        msg_buttons = 0 + 48
                        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    End If
                    .TypeParagraph()
                    .TypeParagraph()
                    '***********************************************************************************************
                    '*  6.写各工况下伸缩管状态表-（表5...）
                    '***********************************************************************************************
                    '***********************************************************************************************
                    '*  6.1 确定写几个表
                    '***********************************************************************************************
                    SQL_command = "select DISTINCT 元件序号,元件名称 from 工况_伸缩管状态 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    If RECreader.HasRows Then
                        j = 1
                        While RECreader.Read
                            '***********************************************************************************************
                            '*  6.1 写管柱轴向变形和效应值（m）-（表4）标题行
                            '***********************************************************************************************
                            .Font.Size = 12
                            .Font.Name = "宋体"
                            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .TypeText(Text:="表" & Trim(Str(4 + j)) & well_name & "井" & zuoye_name & Trim(RECreader.Item("元件序号").ToString) _
                                      & "号" & Trim(RECreader.Item("元件名称").ToString) & "状态表")
                            .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                            .TypeParagraph()
                            '***********************************************************************************************
                            '*  6.2 设定伸缩管状态表的行数、列数
                            '***********************************************************************************************
                            .Font.Size = 10.5
                            mytable = .Tables.Add(.Range, 9, gk_count + 1)
                            .Tables.Item(1).Select()
                            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Tables.Item(1).Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter
                            .Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With mytable.Borders
                                .InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                .InsideLineStyle = True
                                '.OutsideLineStyle = wdLineStyleDouble
                                .OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End With
                            '***********************************************************************************************
                            '*  5.3 填写伸缩管状态表表头
                            '***********************************************************************************************
                            mytable.Cell(1, 1).Range.InsertAfter("工况")
                            mytable.Cell(2, 1).Range.InsertAfter("内压（MPa）")
                            mytable.Cell(3, 1).Range.InsertAfter("外压（MPa）")
                            mytable.Cell(4, 1).Range.InsertAfter("不伸缩轴力（kN）")
                            mytable.Cell(5, 1).Range.InsertAfter("剪切力（kN）")
                            mytable.Cell(6, 1).Range.InsertAfter("轴力（kN）")
                            mytable.Cell(7, 1).Range.InsertAfter("剪销状态")
                            mytable.Cell(8, 1).Range.InsertAfter("伸缩状态")
                            mytable.Cell(9, 1).Range.InsertAfter("伸缩长度（m）")
                            i = 1
                            For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                                mytable.Cell(1, i + 1).Range.InsertAfter(gk_row.Item("工况名称").ToString)
                                '***********************************************************************************************
                                '*  读取工况_伸缩管状态并写入
                                '***********************************************************************************************
                                SQL_command = "select * from 工况_伸缩管状态 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'" _
                                    & " and 工况序号=" & gk_row.Item("工况序号").ToString & " and 元件序号=" & RECreader.Item("元件序号").ToString
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader2 = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If RECreader2.HasRows Then
                                    RECreader2.Read()
                                    mytable.Cell(2, 1 + i).Range.InsertAfter(RECreader2.Item("内压MPa").ToString)
                                    mytable.Cell(3, 1 + i).Range.InsertAfter(RECreader2.Item("外压MPa").ToString)
                                    mytable.Cell(4, 1 + i).Range.InsertAfter(RECreader2.Item("不伸缩等效轴力kN").ToString)
                                    mytable.Cell(5, 1 + i).Range.InsertAfter(RECreader2.Item("剪切力kN").ToString)
                                    mytable.Cell(6, 1 + i).Range.InsertAfter(RECreader2.Item("等效轴力kN").ToString)
                                    mytable.Cell(7, 1 + i).Range.InsertAfter(RECreader2.Item("剪销状态").ToString)
                                    mytable.Cell(8, 1 + i).Range.InsertAfter(RECreader2.Item("伸缩状态").ToString)
                                    mytable.Cell(9, 1 + i).Range.InsertAfter(RECreader2.Item("伸缩长度m").ToString)
                                End If
                                RECreader2.Close()
                                i = i + 1
                            Next
                            .EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                            .TypeParagraph()
                            .TypeParagraph()
                        End While
                    End If
                Else
                    msg_prompt = "    没有找到工况数据，无法输出写工况参数表、管柱载荷、应力及安全系数数值表、管柱轴向变形和效应值表。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                End If
            End With
        End With
        cn_userdb.Close()
        cn_userdb.Dispose()
        frmprogress.Hide()
        If SaveFileDialog1.FileName <> "" Then
            FileName = SaveFileDialog1.FileName
            WordBook.SaveAs(FileName)
            MsgBox("文件已经成功的保存到Word文档中！", , "提示")
        End If
        WordBook.Close()
        WordWP.Visible = False
        WordWP.Quit()
        Exit Sub
errhandler:
        frmprogress.Hide()
        msg_prompt = "导出出错，请检查Word安装是否正确，所有Winword进程是否关闭！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '第3页 CheckBox1选择
    '*********************************************************************************************************************************************
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Call page3_jdlb_fill()
    End Sub
    '*********************************************************************************************************************************************
    '第3页节点列表填充
    '*********************************************************************************************************************************************
    Private Sub page3_jdlb_fill()
        Dim i As Short
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        If CheckBox1.CheckState = 0 Then
            SQL_command = "select distinct 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m " _
                & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
                & " and 节点性质<>'计算点' and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
                & " order by 节点计算参数表.节点编号"
        Else
            SQL_command = "select distinct 节点计算参数表.节点编号,节点ID,节点性质,节点情况表.节点下深m " _
            & " from 节点情况表,节点计算参数表 where 节点计算参数表.井号='" & well_name & "' and 节点计算参数表.作业名称='" & zuoye_name & "'" _
            & " and 节点情况表.节点编号=节点计算参数表.节点编号 and 节点情况表.井号=节点计算参数表.井号 and 节点情况表.作业名称=节点计算参数表.作业名称" _
            & " order by 节点计算参数表.节点编号"
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        pg3_jdlb_Table.Clear()
        ad.Fill(pg3_jdlb_Table)
        ad.Dispose()
        BindingSource12.DataSource = pg3_jdlb_Table
        DataGridView3.ClearSelection()
        DataGridView3.DataSource = BindingSource12
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
        DataGridView3.Columns(0).Width = 30
        DataGridView3.Columns(1).Width = 45
        DataGridView3.Columns(2).Width = 65
        DataGridView3.Columns(3).Width = 80
        For i = 4 To DataGridView3.Columns.Count - 1
            DataGridView3.Columns(i).Visible = False
        Next i
        DataGridView3.Refresh()
        DataGridView3.Show()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    按数据类型显示页读取待选节点列表数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub


    '**************************************************************************************************
    '第3页图形初始化程序，前提是需要工况列表不为空
    '**************************************************************************************************
    Private Sub draw_init2()
        Dim num_of_gk As Short
        Dim i As Short
        Dim j As Short
        Dim titletext() As String
        Dim gk_row As DataRow
        Dim ls_V As UInteger
        num_of_gk = 0
        AxiPlotX2.RemoveAllChannels()
        AxiPlotX2.TitleText = "请选择绘图参数类型"
        '加一个Channel，选安全系数时为许用安全系数，选等效轴力时为管体/接头抗拉强度，选内外压差时为抗内压强度
        j = AxiPlotX2.AddChannel()
        AxiPlotX2.get_Channel(j).TitleText = "许用安全系数"
        AxiPlotX2.get_Channel(j).Color = GetUint32color(1)
        AxiPlotX2.get_Channel(j).Visible = False
        AxiPlotX2.get_Channel(j).VisibleInLegend = False
        '加一个Channel，选等效轴力时为接头加抗拉强度，选内外压差时为为抗外挤强度
        j = AxiPlotX2.AddChannel()
        AxiPlotX2.get_Channel(j).TitleText = "接头抗拉强度kN"
        AxiPlotX2.get_Channel(j).Color = GetUint32color(1)
        AxiPlotX2.get_Channel(j).Visible = False
        AxiPlotX2.get_Channel(j).VisibleInLegend = False

        '1.查有多少工况！
        If dst.Tables("gk_xhmc_Table").Rows.Count <> 0 Then
            num_of_gk = dst.Tables("gk_xhmc_Table").Rows.Count
            ReDim titletext(num_of_gk)
            i = 0
            For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                titletext(i) = gk_row.Item("工况序号").ToString & "：" & gk_row.Item("工况名称").ToString
                j = AxiPlotX2.AddChannel()
                AxiPlotX2.get_Channel(j).TitleText = titletext(i)
                AxiPlotX2.get_Channel(j).Color = GetUint32color(j)
                ls_V = AxiPlotX2.get_Channel(j).Color
                i = i + 1
            Next
        End If
        '2 设置x轴
        AxiPlotX2.get_XAxis(0).Title = "井深(m)"
        AxiPlotX2.get_XAxis(0).Min = 0
        AxiPlotX2.get_XAxis(0).Span = 3000
        If MaxMinCSTable.Rows.Count > 0 Then
            If MaxMinCSTable.Rows(0).Item("max_xs").ToString() <> "" Then
                AxiPlotX2.get_XAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xs").ToString())
            End If
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
    End Sub
    '**************************************************************************************************
    '选择第3页中option1后绘制图形,"内外压差"这个是新的，所以与众不同
    '**************************************************************************************************
    Private Sub _Option1_19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _Option1_19.CheckedChanged
        Dim i As Short
        Dim gk_row As DataRow
        Dim temp As Double
        On Error GoTo errhandler
        If dst.Tables("pg2_dxjdlb_Table").Rows.Count > 0 Then
            AxiPlotX2.ClearAllData()
            If MaxMinCSTable.Rows.Count > 0 Then
                If MaxMinCSTable.Rows(0).Item("max_xs").ToString() <> "" Then
                    AxiPlotX2.get_XAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xs").ToString())
                    '*********************************************************************************************************************
                    ' Y轴设置
                    '*********************************************************************************************************************
                    If _Option1_19.Checked = True Then
                        AxiPlotX2.get_Labels(1).Visible = True
                        AxiPlotX2.get_Labels(1).Caption = "安全系数：抗外挤=" & Trim(TextBox7.Text) & "，抗内压=" & Trim(TextBox8.Text)
                        AxiPlotX2.TitleText = "内外压差(MPa)"
                        AxiPlotX2.get_YAxis(0).Min = 1
                        AxiPlotX2.get_YAxis(0).Span = 5
                    End If
                End If
            End If
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            '*********************************************************************************************************************
            ' 选内外压差时显示管体体和接头的抗内压强度、抗挤强度
            '*********************************************************************************************************************
            If _Option1_19.Checked = True Then
                AxiPlotX2.get_Labels(1).Visible = True
                AxiPlotX2.get_Labels(1).Caption = "安全系数：抗外挤=" & Trim(TextBox7.Text) & "，抗内压=" & Trim(TextBox8.Text)
                AxiPlotX2.get_Channel(0).TitleText = "管体/接头抗内压强度MPa"
                AxiPlotX2.get_Channel(0).Visible = True
                AxiPlotX2.get_Channel(0).VisibleInLegend = True
                AxiPlotX2.get_Channel(1).TitleText = "纯外压抗挤强度MPa"
                AxiPlotX2.get_Channel(1).Visible = True
                AxiPlotX2.get_Channel(1).VisibleInLegend = True
                SQL_command = "select DISTINCT 节点编号,节点下深m,抗挤强度MPa,管体抗内压MPa,接头抗内压MPa from 节点情况表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 节点编号"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                EXECOleDbCommand.Dispose()
                While RECreader.Read()
                    temp = Val(RECreader.Item("管体抗内压MPa").ToString())
                    If temp > Val(RECreader.Item("接头抗内压MPa").ToString()) Then
                        temp = Val(RECreader.Item("接头抗内压MPa").ToString())
                    End If
                    If Val(TextBox8.Text) <> 0 And Val(TextBox8.Text) >= 1 Then
                        temp = temp / Val(TextBox8.Text)
                    End If
                    AxiPlotX2.get_Channel(0).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                    temp = -1 * Val(RECreader.Item("抗挤强度MPa").ToString())
                    If Val(TextBox7.Text) <> 0 And Val(TextBox7.Text) >= 1 Then
                        temp = temp / Val(TextBox7.Text)
                    End If
                    AxiPlotX2.get_Channel(1).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                End While
                RECreader.Close()
                AxiPlotX2.get_Channel(0).TraceLineWidth = 2
                AxiPlotX2.get_Channel(1).TraceLineWidth = 2
            Else
                AxiPlotX2.get_Labels(1).Visible = False
                AxiPlotX2.get_Channel(0).Visible = False
                AxiPlotX2.get_Channel(0).VisibleInLegend = False
                AxiPlotX2.get_Channel(1).Visible = False
                AxiPlotX2.get_Channel(1).VisibleInLegend = False
            End If
            '*********************************************************************************************************************
            ' 曲线绘制
            '*********************************************************************************************************************
            If dst.Tables("gk_xhmc_Table").Rows.Count > 0 Then
                i = 2
                For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                    AxiPlotX2.get_Channel(i).TraceLineWidth = 2

                    SQL_command = "select * from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & gk_row.Item("工况序号").ToString & " order by 节点编号"
                    '画抗挤强度
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.HasRows Then
                        While RECreader.Read()
                            If _Option1_19.Checked = True Then
                                temp = -1 * Val(RECreader.Item("抗挤强度MPa").ToString()) / Val(TextBox7.Text)
                                AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                            End If
                        End While
                    End If
                    EXECOleDbCommand.Dispose()
                    RECreader.Close()
                    '断开曲线
                    AxiPlotX2.get_Channel(i).AddXNull(0)
                    AxiPlotX2.get_Channel(i).AddXNull(0)

                    '画曲线
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.HasRows Then
                        While RECreader.Read()
                            If _Option1_19.Checked = True Then
                                temp = Val(RECreader.Item("管内压力MPa").ToString()) - Val(RECreader.Item("管外压力MPa").ToString())
                                AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                            End If
                        End While
                    End If
                    EXECOleDbCommand.Dispose()
                    RECreader.Close()
                    i = i + 1
                Next
            Else
                msg_prompt = "没有结果数据可显示，请点击[计算]按钮进行计算。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            msg_prompt = "    按数据类型显示页读取节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = Option1.GetIndex(eventSender)
            Dim i As Short
            Dim gk_row As DataRow
            Dim temp As Double
            On Error GoTo errhandler
            If dst.Tables("pg2_dxjdlb_Table").Rows.Count > 0 Then
                AxiPlotX2.ClearAllData()
                If MaxMinCSTable.Rows.Count > 0 Then
                    If MaxMinCSTable.Rows(0).Item("max_xs").ToString() <> "" Then
                        AxiPlotX2.get_XAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xs").ToString())
                        '*********************************************************************************************************************
                        ' 各种Y轴设置
                        '*********************************************************************************************************************
                        If Option1(0).Checked = True Then
                            AxiPlotX2.TitleText = "管内压力(MPa)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_gnyl").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_gnyl").ToString())
                        End If
                        If Option1(1).Checked = True Then
                            AxiPlotX2.TitleText = "管外压力(MPa)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_gwyl").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_gwyl").ToString())
                        End If
                        If Option1(2).Checked = True Then
                            AxiPlotX2.TitleText = "真实轴力(kN)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_shz").ToString()) / 1000
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_shz").ToString()) / 1000
                        End If
                        If Option1(3).Checked = True Then
                            AxiPlotX2.TitleText = "等效轴力(kN)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xz").ToString()) / 1000
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xz").ToString()) / 1000
                        End If
                        If Option1(4).Checked = True Then
                            AxiPlotX2.TitleText = "合弯矩(Nm)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_hwj").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_hwj").ToString())
                        End If
                        If Option1(5).Checked = True Then
                            AxiPlotX2.TitleText = "接触力(N/m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_jchl").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_jchl").ToString())
                        End If
                        If Option1(6).Checked = True Then
                            AxiPlotX2.TitleText = "相当应力(MPa)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xgm4").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xgm4").ToString())
                        End If
                        If Option1(7).Checked = True Then
                            AxiPlotX2.TitleText = "节点温度(℃)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_jd_wendu").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_jd_wendu").ToString())
                        End If
                        If Option1(8).Checked = True Then
                            AxiPlotX2.TitleText = "温度变形(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_bx_wd").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_bx_wd").ToString())
                        End If
                        If Option1(9).Checked = True Then
                            AxiPlotX2.TitleText = "轴力变形(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_bx_zl").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_bx_zl").ToString())
                        End If
                        If Option1(10).Checked = True Then
                            AxiPlotX2.TitleText = "鼓胀变形(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_bx_gz").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_bx_gz").ToString())
                        End If
                        If Option1(11).Checked = True Then
                            AxiPlotX2.TitleText = "螺旋变形(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_bx_lx").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_bx_lx").ToString())
                        End If
                        If Option1(12).Checked = True Then
                            AxiPlotX2.TitleText = "综合变形(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_bx_zh").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_bx_zh").ToString())
                        End If
                        If Option1(13).Checked = True Then
                            AxiPlotX2.TitleText = "温度效应(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xy_wd").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xy_wd").ToString())
                        End If
                        If Option1(14).Checked = True Then
                            AxiPlotX2.TitleText = "轴力效应(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xy_zl").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xy_zl").ToString())
                        End If
                        If Option1(15).Checked = True Then
                            AxiPlotX2.TitleText = "鼓胀效应(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xy_gz").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xy_gz").ToString())
                        End If
                        If Option1(16).Checked = True Then
                            AxiPlotX2.TitleText = "螺旋效应(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xy_lx").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xy_lx").ToString())
                        End If
                        If Option1(17).Checked = True Then
                            AxiPlotX2.TitleText = "综合效应(m)"
                            AxiPlotX2.get_YAxis(0).Min = Val(MaxMinCSTable.Rows(0).Item("min_xy_zh").ToString())
                            AxiPlotX2.get_YAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xy_zh").ToString())
                        End If
                        If Option1(18).Checked = True Then
                            AxiPlotX2.TitleText = "安全系数"
                            'AxiPlotX2.YAxis(0).Min = val(MaxMinCSTable.Rows(0).Item("min_aqxs")
                            'AxiPlotX2.YAxis(0).span = val(MaxMinCSTable.Rows(0).Item("max_aqxs")
                            AxiPlotX2.get_YAxis(0).Min = 1
                            AxiPlotX2.get_YAxis(0).Span = 4
                        End If
                    End If
                End If
                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                '*********************************************************************************************************************
                ' 选安全系数时显示，许用安全系数
                '*********************************************************************************************************************
                ' 默认前两个都不显示
                AxiPlotX2.get_Channel(0).Visible = False
                AxiPlotX2.get_Channel(0).VisibleInLegend = False
                AxiPlotX2.get_Channel(1).Visible = False
                AxiPlotX2.get_Channel(1).VisibleInLegend = False
                AxiPlotX2.get_Labels(1).Visible = False
                ' 默认前两个都不显示
                If Option1(18).Checked = True Then
                    AxiPlotX2.get_Channel(0).TitleText = "许用安全系数"
                    AxiPlotX2.get_Channel(0).Visible = True
                    AxiPlotX2.get_Channel(0).VisibleInLegend = True
                    AxiPlotX2.get_Channel(1).Visible = False
                    AxiPlotX2.get_Channel(1).VisibleInLegend = False
                    SQL_command = "select DISTINCT 节点编号,节点下深m from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 节点编号"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    While RECreader.Read()
                        AxiPlotX2.get_Channel(0).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(Text20.Text))
                    End While
                    RECreader.Close()
                    AxiPlotX2.get_Channel(0).TraceLineWidth = 2
                    AxiPlotX2.get_Channel(1).TraceLineWidth = 2
                End If
                '*********************************************************************************************************************
                ' 选等效轴力时显示管体和接头的抗拉强度
                '*********************************************************************************************************************
                If Option1(3).Checked = True Then
                    AxiPlotX2.get_Labels(1).Visible = True
                    AxiPlotX2.get_Labels(1).Caption = "抗拉安全系数=" & Trim(TextBox9.Text)
                    AxiPlotX2.get_Channel(0).TitleText = "管体抗拉强度kN"
                    AxiPlotX2.get_Channel(0).Visible = True
                    AxiPlotX2.get_Channel(0).VisibleInLegend = True
                    AxiPlotX2.get_Channel(1).TitleText = "接头抗拉强度kN"
                    AxiPlotX2.get_Channel(1).Visible = True
                    AxiPlotX2.get_Channel(1).VisibleInLegend = True
                    SQL_command = "select DISTINCT 节点编号,节点下深m,管体抗拉kN,接头抗拉kN from 节点情况表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 节点编号"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    While RECreader.Read()
                        temp = Val(RECreader.Item("管体抗拉kN").ToString())
                        If Val(TextBox9.Text) <> 0 And Val(TextBox9.Text) >= 1 Then
                            temp = temp / Val(TextBox9.Text)
                        End If
                        AxiPlotX2.get_Channel(0).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                        temp = Val(RECreader.Item("接头抗拉kN").ToString())
                        If Val(TextBox9.Text) <> 0 And Val(TextBox9.Text) >= 1 Then
                            temp = temp / Val(TextBox9.Text)
                        End If
                        AxiPlotX2.get_Channel(1).AddXY(Val(RECreader.Item("节点下深m").ToString()), temp)
                    End While
                    RECreader.Close()
                    AxiPlotX2.get_Channel(0).TraceLineWidth = 2
                    AxiPlotX2.get_Channel(1).TraceLineWidth = 2
                End If
                '*********************************************************************************************************************
                ' 各种曲线绘制
                '*********************************************************************************************************************
                If dst.Tables("gk_xhmc_Table").Rows.Count > 0 Then
                    i = 2
                    For Each gk_row In dst.Tables("gk_xhmc_Table").Select
                        SQL_command = "select * from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & gk_row.Item("工况序号").ToString & " order by 节点编号"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                        RECreader = EXECOleDbCommand.ExecuteReader()
                        EXECOleDbCommand.Dispose()
                        If RECreader.HasRows Then
                            While RECreader.Read()
                                If Option1(0).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("管内压力MPa").ToString()))
                                End If
                                If Option1(1).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("管外压力MPa").ToString()))
                                End If
                                If Option1(2).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("真实轴力N").ToString()) / 1000)
                                End If
                                If Option1(3).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("等效轴力N").ToString()) / 1000)
                                    'AxiPlotX2.get_Channel(1).AddXY(Val(RECreader.Item("节点下深m").ToString()), (-1) * Val(RECreader.Item("Fecrh_N").ToString()) / 1000)
                                End If
                                If Option1(4).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("合弯矩Nm").ToString()))
                                End If
                                If Option1(5).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("接触力N").ToString()))
                                End If
                                If Option1(6).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("合成应力MPa").ToString()))
                                End If
                                If Option1(7).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("节点温度℃").ToString()))
                                End If
                                If Option1(8).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("温度变形m").ToString()))
                                End If
                                If Option1(9).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("轴力变形m").ToString()))
                                End If
                                If Option1(10).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("鼓胀变形m").ToString()))
                                End If
                                If Option1(11).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("螺旋变形m").ToString()))
                                End If
                                If Option1(12).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("综合变形m").ToString()))
                                End If
                                If Option1(13).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("温度效应m").ToString()))
                                End If
                                If Option1(14).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("轴力效应m").ToString()))
                                End If
                                If Option1(15).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("鼓胀效应m").ToString()))
                                End If
                                If Option1(16).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("螺旋效应m").ToString()))
                                End If
                                If Option1(17).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("综合效应m").ToString()))
                                End If
                                If Option1(18).Checked = True Then
                                    AxiPlotX2.get_Channel(i).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("安全系数").ToString()))
                                End If
                            End While
                            AxiPlotX2.get_Channel(i).TraceLineWidth = 2
                        End If
                        RECreader.Close()
                        i = i + 1
                    Next
                End If
                cn_userdb.Close()
                cn_userdb.Dispose()
            Else
                msg_prompt = "没有结果数据可显示，请点击[计算]按钮进行计算。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
            Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            msg_prompt = "    按数据类型显示页读取节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '*********************************************************************************************************************************************
    ' 点击“帮助”按钮
    '*********************************************************************************************************************************************
    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    ' 点击“退出”按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Close()
    End Sub
    '*********************************************************************************************
    '“点击按工况显示-计算结果图形”页“绘图”按钮
    '*********************************************************************************************
    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Dim gk_xh_str As String
        'Call draw_init()
        On Error GoTo errhandler
        AxiPlotX1.get_XAxis(0).Title = "井深(m)"
        gk_xh_str = "0"
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
            gk_xh_str = IIf(BindingSource1.Current("工况序号").ToString = "", "0", BindingSource1.Current("工况序号").ToString)
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 节点计算参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & gk_xh_str & " order by 节点编号"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            While RECreader.Read()
                If Check(0).CheckState = 1 Then
                    AxiPlotX1.get_Channel(0).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("管内压力MPa").ToString()))
                End If
                If Check(1).CheckState = 1 Then
                    AxiPlotX1.get_Channel(1).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("管外压力MPa").ToString()))
                End If
                If Check(2).CheckState = 1 Then
                    AxiPlotX1.get_Channel(2).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("真实轴力N").ToString()) / 1000)
                End If
                If Check(3).CheckState = 1 Then
                    AxiPlotX1.get_Channel(3).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("等效轴力N").ToString()) / 1000)
                End If
                If Check(4).CheckState = 1 Then
                    AxiPlotX1.get_Channel(4).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("合弯矩Nm").ToString()))
                End If
                If Check(5).CheckState = 1 Then
                    AxiPlotX1.get_Channel(5).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("接触力N").ToString()))
                End If
                If Check(6).CheckState = 1 Then
                    AxiPlotX1.get_Channel(6).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("合成应力MPa").ToString()))
                End If
                If Check(7).CheckState = 1 Then
                    AxiPlotX1.get_Channel(7).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("节点温度℃").ToString()))
                End If
                If Check(8).CheckState = 1 Then
                    AxiPlotX1.get_Channel(8).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("温度变形m").ToString()))
                End If
                If Check(9).CheckState = 1 Then
                    AxiPlotX1.get_Channel(9).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("轴力变形m").ToString()))
                End If
                If Check(10).CheckState = 1 Then
                    AxiPlotX1.get_Channel(10).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("鼓胀变形m").ToString()))
                End If
                If Check(11).CheckState = 1 Then
                    AxiPlotX1.get_Channel(11).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("螺旋变形m").ToString()))
                End If
                If Check(12).CheckState = 1 Then
                    AxiPlotX1.get_Channel(12).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("综合变形m").ToString()))
                End If
                If Check(13).CheckState = 1 Then
                    AxiPlotX1.get_Channel(13).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("温度效应m").ToString()))
                End If
                If Check(14).CheckState = 1 Then
                    AxiPlotX1.get_Channel(14).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("轴力效应m").ToString()))
                End If
                If Check(15).CheckState = 1 Then
                    AxiPlotX1.get_Channel(15).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("鼓胀效应m").ToString()))
                End If
                If Check(16).CheckState = 1 Then
                    AxiPlotX1.get_Channel(16).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("螺旋效应m").ToString()))
                End If
                If Check(17).CheckState = 1 Then
                    AxiPlotX1.get_Channel(17).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("综合效应m").ToString()))
                End If
                If Check(18).CheckState = 1 Then
                    AxiPlotX1.get_Channel(18).AddXY(Val(RECreader.Item("节点下深m").ToString()), Val(RECreader.Item("安全系数").ToString()))
                End If
                '            If Check(13).Value = 1 Then
                '                AxiPlotX1.Channel(13).AddXY SQL_rst_jd.Fields("节点下深m"), SQL_rst_jd.Fields("等效悬持力N")
                '            End If
            End While
        Else
            msg_prompt = "没有计算数据，无法绘图！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "    按工况显示-计算结果图形绘制读取节点计算数据时出错，可能是因软件升级老数据不完整所致，请检查各输入界面参数是否输入完整，点击[计算]按钮再次计算。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************
    '“按工况显示-计算结果图形”页绘图选择check选择
    '*********************************************************************************************
    Private Sub Check_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check.CheckStateChanged
        Dim Index As Short = Check.GetIndex(eventSender)
        Call draw_init()
    End Sub
    '**************************************************************************************************
    '“按工况显示-计算结果图形”页图形初始化程序，前提是需要该页工况列表不为空
    '**************************************************************************************************
    Private Sub draw_init()
        Dim draw_caption As String
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim gk_xh_str As String
        Dim max_canshu(19) As Double
        Dim min_canshu(19) As Double
        Dim titletext(19) As String
        Dim canshu_ok(19) As Boolean
        gk_xh_str = "0"
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("工况序号"))) Then
            gk_xh_str = IIf(BindingSource1.Current("工况序号").ToString = "", "0", BindingSource1.Current("工况序号").ToString)
        End If
        draw_caption = BindingSource1.Current("工况序号").ToString & ":" & BindingSource1.Current("工况名称").ToString
        AxiPlotX1.TitleText = draw_caption
        AxiPlotX1.TitleVisible = False

        AxiPlotX1.RemoveAllChannels()
        AxiPlotX1.RemoveAllYAxes()
        AxiPlotX1.RemoveAllXAxes()
        AxiPlotX1.AddXAxis()
        AxiPlotX1.get_XAxis(0).Title = "井深(m)"
        AxiPlotX1.get_XAxis(0).TitleFontColor = &H0&
        AxiPlotX1.get_XAxis(0).TitleRotated = False
        AxiPlotX1.get_XAxis(0).TitleShow = True
        AxiPlotX1.get_XAxis(0).Min = 0
        AxiPlotX1.get_XAxis(0).ScaleLinesColor = &H0&
        AxiPlotX1.get_XAxis(0).ScaleLineShow = True
        AxiPlotX1.get_XAxis(0).LabelsFontColor = &H0&
        AxiPlotX1.get_XAxis(0).LabelsVisible = True
        AxiPlotX1.get_XAxis(0).ReverseScale = True
        AxiPlotX1.get_XAxis(0).ZOrder = 0
        For i = 0 To 18
            canshu_ok(i) = False
            j = AxiPlotX1.AddChannel()
            k = AxiPlotX1.AddYAxis()
            AxiPlotX1.get_Channel(i).Visible = False
            AxiPlotX1.get_Channel(i).Color = GetUint32color(i + 1)
            AxiPlotX1.get_Channel(i).YAxisName = AxiPlotX1.get_YAxis(i).Name
            AxiPlotX1.get_YAxis(i).Visible = False
            AxiPlotX1.get_YAxis(i).ScaleLinesColor = GetUint32color(i + 1)
            AxiPlotX1.get_YAxis(i).LabelsFontColor = &H0&
            AxiPlotX1.get_YAxis(i).LabelsVisible = False
            AxiPlotX1.get_YAxis(i).AlignRefAxisName = AxiPlotX1.get_YAxis(i).Name
            AxiPlotX1.get_YAxis(i).ZOrder = i + 2
            AxiPlotX1.get_YAxis(i).TitleShow = False
            AxiPlotX1.get_YAxis(i).ScaleLineShow = False
            AxiPlotX1.get_Channel(i).VisibleInLegend = False
        Next
        AxiPlotX1.get_Labels(0).Caption = draw_caption
        AxiPlotX1.get_Labels(0).Visible = True
        AxiPlotX1.get_Labels(0).ZOrder = 21
        AxiPlotX1.get_ToolBar(0).ZOrder = 22
        If MaxMinCSTable.Rows(0).Item("max_xs").ToString() <> "" Then
            AxiPlotX1.get_XAxis(0).Span = Val(MaxMinCSTable.Rows(0).Item("max_xs").ToString())
            max_canshu(0) = Val(MaxMinCSTable.Rows(0).Item("max_gnyl").ToString())
            min_canshu(0) = Val(MaxMinCSTable.Rows(0).Item("min_gnyl").ToString())
            titletext(0) = "管内压力(MPa)"
            canshu_ok(0) = True
            max_canshu(1) = Val(MaxMinCSTable.Rows(0).Item("max_gwyl").ToString())
            min_canshu(1) = Val(MaxMinCSTable.Rows(0).Item("min_gwyl").ToString())
            titletext(1) = "管外压力(MPa)"
            canshu_ok(1) = True
            max_canshu(2) = Val(MaxMinCSTable.Rows(0).Item("max_shz").ToString()) / 1000
            min_canshu(2) = Val(MaxMinCSTable.Rows(0).Item("min_shz").ToString()) / 1000
            titletext(2) = "真实轴力(kN)"
            canshu_ok(2) = True
            max_canshu(3) = Val(MaxMinCSTable.Rows(0).Item("max_xz").ToString()) / 1000
            min_canshu(3) = Val(MaxMinCSTable.Rows(0).Item("min_xz").ToString()) / 1000
            titletext(3) = "等效轴力(kN)"
            canshu_ok(3) = True
            max_canshu(4) = Val(MaxMinCSTable.Rows(0).Item("max_hwj").ToString())
            min_canshu(4) = Val(MaxMinCSTable.Rows(0).Item("min_hwj").ToString())
            titletext(4) = "合弯矩(Nm)"
            canshu_ok(4) = True
            max_canshu(5) = Val(MaxMinCSTable.Rows(0).Item("max_jchl").ToString())
            min_canshu(5) = Val(MaxMinCSTable.Rows(0).Item("min_jchl").ToString())
            titletext(5) = "接触力(N/m)"
            canshu_ok(5) = True
            max_canshu(6) = Val(MaxMinCSTable.Rows(0).Item("max_xgm4").ToString())
            min_canshu(6) = Val(MaxMinCSTable.Rows(0).Item("min_xgm4").ToString())
            titletext(6) = "相当应力MPa"
            canshu_ok(6) = True
            max_canshu(7) = Val(MaxMinCSTable.Rows(0).Item("max_jd_wendu").ToString())
            min_canshu(7) = Val(MaxMinCSTable.Rows(0).Item("min_jd_wendu").ToString())
            titletext(7) = "节点温度(℃)"
            canshu_ok(7) = True
            max_canshu(8) = Val(MaxMinCSTable.Rows(0).Item("max_bx_wd").ToString())
            min_canshu(8) = Val(MaxMinCSTable.Rows(0).Item("min_bx_wd").ToString())
            titletext(8) = "温度变形(m)"
            canshu_ok(8) = True
            max_canshu(9) = Val(MaxMinCSTable.Rows(0).Item("max_bx_zl").ToString())
            min_canshu(9) = Val(MaxMinCSTable.Rows(0).Item("min_bx_zl").ToString())
            titletext(9) = "轴力变形(m)"
            canshu_ok(9) = True
            max_canshu(10) = Val(MaxMinCSTable.Rows(0).Item("max_bx_gz").ToString())
            min_canshu(10) = Val(MaxMinCSTable.Rows(0).Item("min_bx_gz").ToString())
            titletext(10) = "鼓胀变形(m)"
            canshu_ok(10) = True
            max_canshu(11) = Val(MaxMinCSTable.Rows(0).Item("max_bx_lx").ToString())
            min_canshu(11) = Val(MaxMinCSTable.Rows(0).Item("min_bx_lx").ToString())
            titletext(11) = "螺旋变形(m)"
            canshu_ok(11) = True
            max_canshu(12) = Val(MaxMinCSTable.Rows(0).Item("max_bx_zh").ToString())
            min_canshu(12) = Val(MaxMinCSTable.Rows(0).Item("min_bx_zh").ToString())
            titletext(12) = "综合变形(m)"
            canshu_ok(12) = True
            max_canshu(13) = Val(MaxMinCSTable.Rows(0).Item("max_xy_wd").ToString())
            min_canshu(13) = Val(MaxMinCSTable.Rows(0).Item("min_xy_wd").ToString())
            titletext(13) = "温度效应(m)"
            canshu_ok(13) = True
            max_canshu(14) = Val(MaxMinCSTable.Rows(0).Item("max_xy_zl").ToString())
            min_canshu(14) = Val(MaxMinCSTable.Rows(0).Item("min_xy_zl").ToString())
            titletext(14) = "轴力效应(m)"
            canshu_ok(14) = True
            max_canshu(15) = Val(MaxMinCSTable.Rows(0).Item("max_xy_gz").ToString())
            min_canshu(15) = Val(MaxMinCSTable.Rows(0).Item("min_xy_gz").ToString())
            titletext(15) = "鼓胀效应(m)"
            canshu_ok(15) = True
            max_canshu(16) = Val(MaxMinCSTable.Rows(0).Item("max_xy_lx").ToString())
            min_canshu(16) = Val(MaxMinCSTable.Rows(0).Item("min_xy_lx").ToString())
            titletext(16) = "螺旋效应(m)"
            canshu_ok(16) = True
            max_canshu(17) = Val(MaxMinCSTable.Rows(0).Item("max_xy_zh").ToString())
            min_canshu(17) = Val(MaxMinCSTable.Rows(0).Item("min_xy_zh").ToString())
            titletext(17) = "综合效应(m)"
            canshu_ok(17) = True
            max_canshu(18) = Val(MaxMinCSTable.Rows(0).Item("max_aqxs").ToString())
            min_canshu(18) = Val(MaxMinCSTable.Rows(0).Item("min_aqxs").ToString())
            titletext(18) = "安全系数"
            canshu_ok(18) = True
            AxiPlotX1.ClearAllData()
            For i = 0 To 18
                If (canshu_ok(i) = True) Then
                    If Check(i).CheckState = 1 Then
                        AxiPlotX1.get_Channel(i).Visible = True
                        AxiPlotX1.get_YAxis(i).Visible = True
                        AxiPlotX1.get_YAxis(i).LabelsVisible = True
                        AxiPlotX1.get_YAxis(i).ScaleLineShow = True
                        AxiPlotX1.get_Channel(i).VisibleInLegend = True
                        AxiPlotX1.get_YAxis(i).Min = min_canshu(i) - (max_canshu(i) - min_canshu(i)) * 0.05
                        AxiPlotX1.get_YAxis(i).Span = (max_canshu(i) - min_canshu(i)) * 1.1
                        AxiPlotX1.get_Channel(i).TitleText = titletext(i)
                        AxiPlotX1.get_Channel(i).TraceLineWidth = 2
                    End If
                End If
            Next i
            '对于压力显示，为了便于对比，让坐标分度一致
            If Check(0).CheckState = 1 And Check(1).CheckState = 1 Then
                If min_canshu(0) > min_canshu(1) Then
                    AxiPlotX1.get_YAxis(0).Min = min_canshu(1)
                    AxiPlotX1.get_YAxis(1).Min = min_canshu(1)
                Else
                    AxiPlotX1.get_YAxis(0).Min = min_canshu(0)
                    AxiPlotX1.get_YAxis(1).Min = min_canshu(0)
                End If
                If max_canshu(0) > max_canshu(1) Then
                    AxiPlotX1.get_YAxis(0).Span = max_canshu(0)
                    AxiPlotX1.get_YAxis(1).Span = max_canshu(0)
                Else
                    AxiPlotX1.get_YAxis(0).Span = max_canshu(1)
                    AxiPlotX1.get_YAxis(1).Span = max_canshu(1)
                End If
            End If
            '对于轴力显示，为了便于对比，让坐标分度一致
            If Check(2).CheckState = 1 And Check(3).CheckState = 1 Then
                If min_canshu(2) > min_canshu(3) Then
                    AxiPlotX1.get_YAxis(2).Min = min_canshu(3)
                    AxiPlotX1.get_YAxis(3).Min = min_canshu(3)
                Else
                    AxiPlotX1.get_YAxis(2).Min = min_canshu(2)
                    AxiPlotX1.get_YAxis(3).Min = min_canshu(2)
                End If
                If max_canshu(2) > max_canshu(3) Then
                    AxiPlotX1.get_YAxis(2).Span = max_canshu(2)
                    AxiPlotX1.get_YAxis(3).Span = max_canshu(2)
                Else
                    AxiPlotX1.get_YAxis(2).Span = max_canshu(3)
                    AxiPlotX1.get_YAxis(3).Span = max_canshu(3)
                End If
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    ' 点击“计算”按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        If IsNumeric(Text1.Text) = False Then
            msg_prompt = "请输入正确的节点间距值！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Call write_chanshu("管柱力学分析-节点间距", CStr(Text1.Text))
        If lxfx_prs(Val(Text1.Text)) = False Then
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Call Fill_MaxMinCSTable()
        Call draw_init()
        Call page12_ssglb_fill()
        'Call page1_DataGridView_flash()
        Call page12_jdlb_fill()
        Call draw_init2()
        Call page2_dxjdlb_fill()
        Call draw_init3()
        Call page3_jdlb_fill()
        'Call draw_AFEDP_curve()
    End Sub
    '*********************************************************************************************************************************************
    '点击“绘图”按钮事件
    '*********************************************************************************************************************************************
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '绘制井身结构图,含管柱
        Call drawWellStruction(Picture1, 2)
    End Sub
End Class