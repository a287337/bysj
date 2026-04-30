Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class ipt_gzhu
    '*********************************************************************************************************************************************
    '                                                   关于管柱组合输入与管理界面的说明
    ' 程序升级记事：
    '                                                                                           秦彦斌 2019年11月04日最后整理
    '  （1）让SSTab1各页标题宽度一样，不要挤在一起。
    '  （2）用程序的方法设置DataGridView的外观很方便，不用在界面属性中设置，避免遗漏或忘记如何操作。
    '   (3) fill_gzhgrid过程中，可让DataGridView中的某些列不显示。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    '油井管
    Private yjg_Table As DataTable = b_dst.Tables.Add("yjg_Table")
    '节流工具
    Private jlgj_Table As DataTable = b_dst.Tables.Add("jlgj_Table")
    '封隔器
    Private fgq_Table As DataTable = b_dst.Tables.Add("fgq_Table")
    '开关工具
    Private kggj_Table As DataTable = b_dst.Tables.Add("kggj_Table")
    '伸缩管
    Private shsg_Table As DataTable = b_dst.Tables.Add("shsg_Table")
    '锚定工具
    Private mdgj_Table As DataTable = b_dst.Tables.Add("mdgj_Table")
    '普通钻杆
    Private ptzg_Table As DataTable = b_dst.Tables.Add("ptzg_Table")
    '射孔枪   李润洲2025年6月
    Private skqd_Table As DataTable = b_dst.Tables.Add("skqd_Table")
    '筛管   李润洲2025年7月24日
    Private sg_Table As DataTable = b_dst.Tables.Add("sg_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub ipt_gzhu_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub ipt_gzhu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim gzh_yjxz As String
        Dim pageCount As Integer
        Dim totalWidth As Integer
        Dim pageWidth As Integer

        '管柱共性数据
        Text7.Text = ""
        Text8.Text = CStr(0)
        Text9.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text12.Text = ""
        '油井管
        Text13.Text = CStr(0.0#)
        Text14.Text = ""
        Text110.Text = ""
        Text15.Text = CStr(0.0#)
        Text17.Text = CStr(0.0#)
        Text21.Text = CStr(0.0#)
        Text20.Text = CStr(0.0#)
        Text16.Text = CStr(0.0#)
        Text18.Text = CStr(206842.72#)
        Text19.Text = CStr(0.3)
        Text113.Text = CStr(0.0#)
        Text114.Text = CStr(0.0#)
        Text109.Text = ""
        TextBox1.Text = CStr(0.0000124#)
        '开关元件
        Text30.Text = ""
        Text75.Text = ""
        Text92.Text = CStr(0.0#)
        Text25.Text = CStr(0.0#)
        Text76.Text = CStr(0.0#)
        Text88.Text = CStr(0.0#)
        Text29.Text = CStr(0.0#)
        Text27.Text = CStr(0.0#)
        Text28.Text = CStr(0.0#)
        Text26.Text = CStr(0.0#)
        Text89.Text = ""
        Text90.Text = ""
        Text93.Text = ""
        Text91.Text = ""
        '封隔器
        Text33.Text = CStr(0.0#)
        Text34.Text = CStr(0.0#)
        Text35.Text = CStr(0.0#)
        Text36.Text = CStr(0.0#)
        Text37.Text = CStr(0.0#)
        Text38.Text = CStr(0.0#)
        Text39.Text = CStr(0.0#)
        Text40.Text = CStr(0.0#)
        Combo1.Text = ""
        Text22.Text = ""
        Text35.Text = ""
        Combo1.Text = ""
        Text54.Text = CStr(0.0#)
        Text55.Text = CStr(0.0#)
        Text52.Text = ""
        Text58.Text = ""
        Text53.Text = CStr(0.0#)
        Text23.Text = ""
        Text24.Text = ""
        Text39.Text = CStr(0.0#)
        Text40.Text = CStr(0.0#)
        Text36.Text = CStr(0.0#)
        Text37.Text = CStr(0.0#)
        Text34.Text = CStr(0.0#)
        Text56.Text = CStr(0.0#)
        Text33.Text = CStr(0.0#)
        Text57.Text = CStr(0.0#)
        Text38.Text = CStr(0.0#)
        Text31.Text = CStr(0.0#)
        Text32.Text = CStr(0.0#)
        Text41.Text = CStr(0.0#)
        Text42.Text = CStr(0.0#)
        Text43.Text = CStr(0.0#)
        Text44.Text = CStr(0.0#)
        Text45.Text = ""
        Combo3.Text = "不能"
        Textf2.Text = 0.0
        Textf1.Text = 0.0

        '伸缩元件
        Text49.Text = ""
        Text47.Text = ""
        Text67.Text = ""
        Text68.Text = ""
        Text60.Text = CStr(0.0#)
        Text61.Text = CStr(0.0#)
        Text69.Text = ""
        Text59.Text = ""
        Text62.Text = CStr(0.0#)
        Text64.Text = CStr(0.0#)
        Text46.Text = CStr(0.0#)
        Text66.Text = CStr(0.0#)
        Text48.Text = CStr(0.0#)
        Text65.Text = CStr(0.0#)
        Text63.Text = CStr(0.0#)
        Text50.Text = CStr(0.0#)
        Text51.Text = CStr(0.0#)
        TextBox3.Text = CStr(0.0#)
        '锚定工具
        Text79.Text = ""
        Combo2.Text = ""
        Text80.Text = CStr(0.0#)
        Text85.Text = ""
        Text74.Text = CStr(0.0#)
        Text73.Text = CStr(0.0#)
        Text78.Text = ""
        Text86.Text = CStr(0.0#)
        Text77.Text = ""
        Text72.Text = CStr(0.0#)
        Text84.Text = CStr(0.0#)
        Text83.Text = CStr(0.0#)
        Text81.Text = CStr(0.0#)
        Text87.Text = CStr(0.0#)
        Text71.Text = CStr(0.0#)
        Text82.Text = CStr(0.0#)
        Text70.Text = ""
        '普通钻杆
        Text108.Text = CStr(0.0#)
        Text105.Text = CStr(0.0#)
        Text107.Text = CStr(0.0#)
        Text111.Text = CStr(0.0#)
        Text112.Text = ""
        Text106.Text = ""
        Text104.Text = CStr(0.0#)
        Text103.Text = CStr(0.0#)
        Text102.Text = ""
        Text101.Text = ""
        Text100.Text = CStr(0.0#)
        Text99.Text = CStr(0.0#)
        Text98.Text = CStr(0.0#)
        Text97.Text = CStr(0.0#)
        Text96.Text = CStr(206842.72#)
        Text95.Text = CStr(0.3)
        Text94.Text = CStr(0.0#)
        TextBox2.Text = CStr(0.0000124#)

        '射孔枪弹 李润洲2025年6月
        Texts1.Text = ""
        Texts2.Text = ""
        Texts5.Text = CStr(0.0#)
        Texts6.Text = CStr(0.0#)
        Texts7.Text = CStr(0.0#)
        Texts9.Text = ""
        'Texts10.Text = "" 炸药名
        Texts11.Text = CStr(0.0#)
        Texts12.Text = CStr(0.0#)
        Texts13.Text = CStr(0.0#)
        Texts14.Text = CStr(0.0#)
        Texts15.Text = CStr(0.0#)
        Texts19.Text = ""
        Texts20.Text = ""

        '筛管 李润洲2025年7月24日
        Textk1.Text = ""
        Textk2.Text = CStr(0.0#)
        Textk3.Text = CStr(0.0#)
        Textk4.Text = CStr(0.0#)
        Textk5.Text = CStr(0.0#)
        Textk6.Text = CStr(0.0#)
        Textk7.Text = CStr(0.0#)
        Textk8.Text = CStr(0.0#)
        Textk9.Text = CStr(0.0#)
        Textk10.Text = CStr(0.0#)
        Textk11.Text = CStr(0.0#)
        Textk12.Text = ""
        Textk13.Text = CStr(0.0#)
        Textk14.Text = ""
        Textk15.Text = CStr(0.0#)

        Label_0.Text = well_name & "井油气井参数"
        Label_24.Text = well_name & "井" & zuoye_name & "管柱组成"
        Me.Text = well_name & "井" & zuoye_name & "管柱组合数据输入与修改"
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
        '管柱长度汇总
        Label_25.Text = "共0米"
        SQL_command = "select sum(元件长度m) as 总长度  from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            Label_25.Text = "共" & RECreader.Item("总长度").ToString & "米"
        End If
        RECreader.Close()
        cn_userdb.Close()

        '用ADO.NET给ListBox1赋值，原来用Adodc2。
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        ListBox1.Items.Clear()
        '*******************************************************************************
        '李润洲2025年6月修改，在data_str中将管柱元件性质表中的射孔枪的enable设置为了1
        '2025年7月25日，插入了筛管记录
        '*******************************************************************************
        SQL_command = "select 管柱元件性质 from 管柱元件性质 where enabled=1"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        While RECreader.Read
            gzh_yjxz = RECreader.Item("管柱元件性质").ToString
            ListBox1.Items.Add(gzh_yjxz)
        End While
        Combo1.Items.Clear()
        SQL_command = "select * from 封隔器坐封方式"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        While RECreader.Read
            gzh_yjxz = RECreader.Item("坐封方式").ToString
            Combo1.Items.Add(gzh_yjxz)
        End While
        'Combo2.Items.Clear()
        'SQL_command = "select * from 封隔器定位方式"
        'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        'RECreader = EXECOleDbCommand.ExecuteReader()
        'While RECreader.Read
        '    gzh_yjxz = RECreader.Item("定位方式").ToString
        '    Combo2.Items.Add(gzh_yjxz)
        'End While
        ListBox2.Items.Clear()
        SQL_command = "select DISTINCT 开关类型 from 开关元件 order by 开关类型"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        While RECreader.Read
            gzh_yjxz = RECreader.Item("开关类型").ToString
            ListBox2.Items.Add(gzh_yjxz)
        End While
        RECreader.Close()
        cn_basedb.Close()

        '（1）让SSTab1各页标题宽度一样，不要挤在一起。（2）确定哪些管柱原件选择页可显示（这个VS实现不了了）。
        '这个Fixed设置是必须的
        SSTab1.SizeMode = TabSizeMode.Fixed
        '设置标签宽度
        totalWidth = SSTab1.Width
        pageCount = SSTab1.TabPages.Count
        '最后-1 因为tabcontrol有留margin，得空出margin的空间
        pageWidth = totalWidth / pageCount - 1
        'ItemSize在SizeMode = TabSizeMode.Fixed才生效
        SSTab1.ItemSize = New Size(pageWidth, SSTab1.ItemSize.Height)
        Call sstabshow()
        '管柱数据列表DataGridView1设置并填充函数fill_gzhgrid()，原来用DataGrid7和Adodc4
        Call fill_gzhgrid()
        '油井管数据列表DataGridView2设置并填充函数fill_yjggrid()，原来用 DataGrid1和Adodc3
        Call fill_yjggrid()
        '节流工具数据列表DataGridView3设置并填充函数fill_jlgjgrid()，原来用 DataGrid5和Adodc8
        Call fill_jlgjgrid()
        '封隔器数据列表DataGridView4设置并填充函数fill_fgqgrid()，原来用 DataGrid3和Adodc6
        Call fill_fgqgrid()
        '开关工具数据列表DataGridView5设置并填充函数fill_kggjgrid()，原来用 DataGrid4和Adodc7
        Call fill_kggjgrid(Text30.Text)
        '伸缩管数据列表DataGridView6设置并填充函数fill_shsggrid()，原来用 DataGrid6和Adodc9
        Call fill_shsggrid()
        '锚定工具数据列表DataGridView7设置并填充函数fill_mdgjgrid()，原来用 DataGrid2和Adodc5
        Call fill_mdgjgrid()
        '普通钻杆数据列表DataGridView8设置并填充函数fill_ptzggrid()，原来用 DataGrid9和Adodc11
        Call fill_ptzggrid()
        '填炸药名列表listbox3，射孔枪弹数据列表DataGridView11设置并填充函数fill_skqdgrid()，李润洲2025年6月
        If fill_zym() Then
            Call fill_skqdgrid()
        End If
        '筛管数据列表DataGridView12设置并填充函数fill_sgGrid()，李润洲2025年7月24日
        Call fill_sgGrid()
        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
    End Sub
    '*********************************************************************************************************************************************
    '管柱数据列表DataGridView1设置并填充函数fill_gzhgrid()，原来用DataGrid7和Adodc4
    '*********************************************************************************************************************************************
    Private Sub fill_gzhgrid()
        Dim i As Integer
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select 元件序号,元件名称,元件性质,元件外径mm,元件内径mm,元件长度m from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'order by 元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gzh_Table.Clear()
        ad.Fill(gzh_Table)
        ad.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
        BindingSource1.DataSource = gzh_Table
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
        DataGridView1.Columns(0).Width = 24
        DataGridView1.Columns(1).Width = 150
        For i = 3 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Visible = False
        Next i
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '油井管数据列表DataGridView2设置并填充函数fill_yjggrid()，原来用 DataGrid1和Adodc3
    '*********************************************************************************************************************************************
    Private Sub fill_yjggrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [油管规格],[油管外径mm],[油管壁厚mm],[材料],[扣型]," _
                & "[单位长度质量kg/m],[屈服应力MPa],[抗外挤强度MPa],[抗内压强度MPa],[抗拉强度kN]," _
                & "[接头抗内压强度MPa],[接头抗拉强度kN],[备注],[弹性模量MPa],[泊松比]," _
                & "[热膨胀系数] " _
                & " from 油管 order by [油管外径mm] desc,[油管壁厚mm] desc,[材料] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        yjg_Table.Clear()
        ad.Fill(yjg_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource2.DataSource = yjg_Table
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
        DataGridView2.Columns(1).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(2).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(3).Width = (DataGridView2.Width - 26) * 0.1
        DataGridView2.Columns(4).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(5).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(6).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(7).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(8).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(9).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(10).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(11).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(12).Width = (DataGridView2.Width - 26) * 0.2
        DataGridView2.Columns(13).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(14).Width = (DataGridView2.Width - 26) * 0.05
        DataGridView2.Columns(15).Width = (DataGridView2.Width - 26) * 0.1
        DataGridView2.Refresh()
        DataGridView2.Show()
    End Sub
    '*********************************************************************************************************************************************
    '节流工具数据列表DataGridView3设置并填充函数fill_jlgjgrid()，原来用 DataGrid5和Adodc8
    '*********************************************************************************************************************************************
    Private Sub fill_jlgjgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 节流元件 order by [外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        jlgj_Table.Clear()
        ad.Fill(jlgj_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource3.DataSource = jlgj_Table
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
        DataGridView3.Refresh()
        DataGridView3.Show()
    End Sub
    '*********************************************************************************************************************************************
    '封隔器数据列表DataGridView4设置并填充函数fill_fgqgrid()，原来用 DataGrid3和Adodc6
    '*********************************************************************************************************************************************
    Private Sub fill_fgqgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [封隔器名称],[型号],[最大外径(mm)],[最小通径(mm)],[长度(m)]," _
                    & "[坐封方式],[重量kg],[生产厂家],[温度范围下℃],[温度范围上℃]," _
                    & "[主体材料],[主材屈服强度MPa],[上端扣型],[极限承压(MPa)],[下端扣型]," _
                    & "[坐封力(kN)],[坐封压差(MPa)],[最小坐封压力MPa],[最大坐封压力MPa],[抗内压强度MPa]," _
                    & "[抗外压强度MPa],[极限载荷(kN)],[中心管外径(mm)],[中心管内径(mm)],[备注],[能否反洗井] from 封隔器 order by [最大外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        fgq_Table.Clear()
        ad.Fill(fgq_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource4.DataSource = fgq_Table
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
        DataGridView4.Refresh()
        DataGridView4.Show()
    End Sub
    '*********************************************************************************************************************************************
    'ListBox2的SelectedIndexChanged事件处理：（1）给 Text30.Text赋值；（2）重新生成待选开关工具列表
    '*********************************************************************************************************************************************
    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        Text30.Text = ListBox2.Text
        Call fill_kggjgrid(Text30.Text)
    End Sub
    '*********************************************************************************************************************************************
    '开关工具数据列表DataGridView5设置并填充函数fill_kggjgrid()，原来用 DataGrid4和Adodc7
    '*********************************************************************************************************************************************
    Private Sub fill_kggjgrid(ByVal kggj_lx As String)
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        If kggj_lx = "" Then
            SQL_command = "select * from 开关元件 order by [外径(mm)] desc"
        Else
            SQL_command = "select * from 开关元件 where 开关类型='" & Trim(kggj_lx) & "' order by [外径(mm)] desc"
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        kggj_Table.Clear()
        ad.Fill(kggj_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource5.DataSource = kggj_Table
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
        DataGridView5.Refresh()
        DataGridView5.Show()
    End Sub
    '*********************************************************************************************************************************************
    '伸缩管数据列表DataGridView6设置并填充函数fill_shsggrid()，原来用 DataGrid6和Adodc9
    '*********************************************************************************************************************************************
    Private Sub fill_shsggrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 伸缩元件 order by [外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        shsg_Table.Clear()
        ad.Fill(shsg_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource6.DataSource = shsg_Table
        DataGridView6.ClearSelection()
        DataGridView6.DataSource = BindingSource6
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
        DataGridView6.Refresh()
        DataGridView6.Show()
    End Sub
    '*********************************************************************************************************************************************
    '锚定工具数据列表DataGridView7设置并填充函数fill_mdgjgrid()，原来用 DataGrid2和Adodc5
    '*********************************************************************************************************************************************
    Private Sub fill_mdgjgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select 型号,工具名称,坐卡方式,最大外径mm,最小通径mm,总体长度m,重量kg,工作压力MPa," _
                    & "温度范围下℃,温度范围上℃,小坐卡压力MPa,大坐卡压力MPa,上端扣型,下端扣型,生产厂家,最大锚定力kN," _
                    & "最大解锚力kN,抗内压强度MPa,抗外压强度MPa,抗拉强度kN,备注 " _
                    & " from 锚定工具表 order by 最大外径mm desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        mdgj_Table.Clear()
        ad.Fill(mdgj_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource7.DataSource = mdgj_Table
        DataGridView7.ClearSelection()
        DataGridView7.DataSource = BindingSource7
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
        DataGridView7.Refresh()
        DataGridView7.Show()
    End Sub
    '*********************************************************************************************************************************************
    '普通钻杆数据列表DataGridView8设置并填充函数fill_ptzggrid()，原来用 DataGrid9和Adodc11
    '*********************************************************************************************************************************************
    Private Sub fill_ptzggrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 钻杆 order by [钻杆外径(mm)]"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        ptzg_Table.Clear()
        ad.Fill(ptzg_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource8.DataSource = ptzg_Table
        DataGridView8.ClearSelection()
        DataGridView8.DataSource = BindingSource8
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
        DataGridView8.Refresh()
        DataGridView8.Show()
    End Sub
    '*********************************************************************************************************************************************
    '射孔枪弹数据列表DataGridView11设置并填充函数fill_skqdgrid()，李润洲2025年6月
    '*********************************************************************************************************************************************
    Private Sub fill_skqdgrid()
        Try
            Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                cn_basedb.Open()
                SQL_command = "select [射孔器名称],[射孔弹名称],[炸药名称],[装药量g],[装药密度g╱cm3],[射孔密度孔╱m]," _
                            & "[射孔相位°],[射孔器分类],[外径mm],[壁厚mm],[耐压MPa]," _
                            & "[耐温°C],[平均孔径mm],[平均穿深mm]," _
                            & "[备注] from 枪弹数据表 order by [炸药名称],[装药量g],[装药密度g╱cm3],[射孔密度孔╱m]"
                Using ad As New OleDbDataAdapter(SQL_command, cn_basedb)
                    skqd_Table.Clear()
                    ad.Fill(skqd_Table)
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "基础数据库-枪弹表访问异常！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
        BindingSource11.DataSource = skqd_Table
        DataGridView11.ClearSelection()
        DataGridView11.DataSource = BindingSource11
        DataGridView11.ResetBindings()
        DataGridView11.AutoGenerateColumns = True
        DataGridView11.AllowUserToAddRows = False
        DataGridView11.AllowUserToDeleteRows = False
        DataGridView11.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView11.MultiSelect = False
        DataGridView11.RowHeadersWidth = 24
        DataGridView11.ReadOnly = True
        DataGridView11.Columns(0).Width = 300
        DataGridView11.Columns(1).Width = 200
        DataGridView11.Columns(7).Width = 200
        DataGridView11.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView11.Refresh()
        DataGridView11.Show()
    End Sub
    '*********************************************************************************************************************************************
    '筛管数据列表DataGridView12设置并填充函数fill_sgGrid()，李润洲2025年7月24日
    '*********************************************************************************************************************************************
    Private Sub fill_sgGrid()
        Try
            Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                cn_basedb.Open()
                SQL_command = "select 规格名称,外径mm,壁厚mm,单根长度m,单根质量kg," _
                                & "主体材料,接头抗拉伸强度kN,主材屈服强度MPa,抗内压强度MPa,抗外挤强度MPa," _
                                & "耐温℃,耐压差MPa,[孔隙率%],有效流通面积cm2╱m,拦截的最小颗粒直径mm," _
                                & "生产厂家,备注 from 筛管 order by 外径mm,壁厚mm,主体材料"
                Using ad As New OleDbDataAdapter(SQL_command, cn_basedb)
                    sg_Table.Clear()
                    ad.Fill(sg_Table)
                End Using
            End Using
        Catch ex As OleDbException
            msg_prompt = "基础数据库-筛管表访问异常！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
        BindingSource12.DataSource = sg_Table
        DataGridView12.ClearSelection()
        DataGridView12.DataSource = BindingSource12
        DataGridView12.ResetBindings()
        DataGridView12.AutoGenerateColumns = True
        DataGridView12.AllowUserToAddRows = False
        DataGridView12.AllowUserToDeleteRows = False
        DataGridView12.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView12.MultiSelect = False
        DataGridView12.RowHeadersWidth = 24
        DataGridView12.ReadOnly = True
        DataGridView12.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView12.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView12.Refresh()
        DataGridView12.Show()
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
    'ListBox1的SelectedIndexChanged事件处理：（1）给 Text12.Text赋值；（2）让SSTab1相应的页选中显示
    '*********************************************************************************************************************************************
    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Text12.Text = ListBox1.Text
        Call sstabshow()
    End Sub
    '*********************************************************************************************************************************************
    '根据 Text12.Text，让SSTab1相应的页选中显示
    '*********************************************************************************************************************************************
    Private Sub sstabshow()
        '李润洲2025年6月增加GrouptBox2（射孔枪）和GroupBox3（筛管）的可见与否设置
        Select Case Text12.Text
            Case "油井管"
                SSTab1.SelectedIndex = 0
                _Frame2_0.Visible = True
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "节流工具"
                SSTab1.SelectedIndex = 1
                _Frame2_0.Visible = False
                _Frame2_1.Visible = True
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "封隔器"
                SSTab1.SelectedIndex = 2
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = True
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "开关工具"
                SSTab1.SelectedIndex = 3
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = True
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "伸缩管"
                SSTab1.SelectedIndex = 4
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = True
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "锚定工具"
                SSTab1.SelectedIndex = 5
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = True
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
            Case "普通钻杆"
                SSTab1.SelectedIndex = 6
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = True
                GroupBox2.Visible = False
                GroupBox3.Visible = False
                '李润洲2025年6月增加
            Case "射孔枪"
                SSTab1.SelectedIndex = 7
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = True
                GroupBox3.Visible = False
                '李润洲2025年7月24日增加
            Case "筛管"
                SSTab1.SelectedIndex = 8
                _Frame2_0.Visible = False
                _Frame2_1.Visible = False
                _Frame2_2.Visible = False
                _Frame2_3.Visible = False
                _Frame2_4.Visible = False
                _Frame2_5.Visible = False
                _Frame2_83.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = True
        End Select
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
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中管柱元件列表中的某行。原来用DataGrid7和Adodc4
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Call myrefresh()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的管柱元件列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh()
        Dim xiashen As Double
        Dim zongchang As Double
        If IsNothing(Me.BindingSource1.Current) Then Exit Sub
        If (Not IsDBNull(Me.BindingSource1.Current("元件序号"))) Then
            Text7.Text = BindingSource1.Current("元件名称").ToString
            Text8.Text = IIf(BindingSource1.Current("元件序号").ToString = "", "0", BindingSource1.Current("元件序号").ToString)
            Text9.Text = IIf(BindingSource1.Current("元件长度m").ToString = "", "0.0", BindingSource1.Current("元件长度m").ToString)
            Text10.Text = IIf(BindingSource1.Current("元件外径mm").ToString = "", "0.0", BindingSource1.Current("元件外径mm").ToString)
            Text11.Text = IIf(BindingSource1.Current("元件内径mm").ToString = "", "0.0", BindingSource1.Current("元件内径mm").ToString)
            Text12.Text = BindingSource1.Current("元件性质").ToString
            Call sstabshow()
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            Select Case Text12.Text
                Case "油井管"
                    SQL_command = "select * from 管柱_油管表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text13.Text = RECreader.Item("油管壁厚mm").ToString
                        Text14.Text = RECreader.Item("油管钢级").ToString
                        Text15.Text = RECreader.Item("单位长重kg╱m").ToString
                        Text16.Text = RECreader.Item("抗拉强度kN").ToString
                        Text17.Text = RECreader.Item("屈服强度MPa").ToString
                        Text20.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text21.Text = RECreader.Item("抗挤强度MPa").ToString
                        Text18.Text = RECreader.Item("弹性模量MPa").ToString
                        Text19.Text = RECreader.Item("泊松比").ToString
                        Text110.Text = RECreader.Item("扣型").ToString
                        Text113.Text = RECreader.Item("接头抗内压强度MPa").ToString
                        Text114.Text = RECreader.Item("接头抗拉强度kN").ToString
                        Text109.Text = RECreader.Item("备注").ToString
                        TextBox1.Text = RECreader.Item("热膨胀系数").ToString
                    End If
                Case "节流工具"
                    SQL_command = "select * from 管柱_节流元件 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text31.Text = RECreader.Item("流孔内径mm").ToString
                        Text32.Text = RECreader.Item("流孔长度mm").ToString
                        Text41.Text = RECreader.Item("抗外挤强度MPa").ToString
                        Text42.Text = RECreader.Item("重量kg").ToString
                        Text43.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text44.Text = RECreader.Item("抗拉强度kN").ToString
                        Text45.Text = RECreader.Item("节流流向").ToString
                    End If
                Case "封隔器"
                    SQL_command = "select * from 管柱_封隔定位元件 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text22.Text = RECreader.Item("型号").ToString
                        Text35.Text = RECreader.Item("生产厂家").ToString
                        Combo1.Text = RECreader.Item("坐封方式").ToString
                        Text54.Text = RECreader.Item("温度范围下℃").ToString
                        Text55.Text = RECreader.Item("温度范围上℃").ToString
                        Text52.Text = RECreader.Item("主体材料").ToString
                        Text58.Text = RECreader.Item("备注").ToString
                        Text53.Text = RECreader.Item("主材屈服强度MPa").ToString
                        Text23.Text = RECreader.Item("上端扣型").ToString
                        Text24.Text = RECreader.Item("下端扣型").ToString
                        Text39.Text = RECreader.Item("极限压差MPa").ToString
                        Text40.Text = RECreader.Item("重量kg").ToString
                        Text36.Text = RECreader.Item("坐封力kN").ToString
                        Text37.Text = RECreader.Item("坐封压力MPa").ToString
                        Text34.Text = RECreader.Item("最小坐封压力MPa").ToString
                        Text56.Text = RECreader.Item("最大坐封压力MPa").ToString
                        Text33.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text57.Text = RECreader.Item("抗外压强度MPa").ToString
                        Text38.Text = RECreader.Item("极限载荷kN").ToString
                        Combo3.Text = RECreader.Item("能否反洗井").ToString
                        Textf2.Text = RECreader.Item("插管外径mm").ToString
                        Textf1.Text = RECreader.Item("插管内径mm").ToString
                    End If
                Case "开关工具"
                    SQL_command = "select * from 管柱_开关元件 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text30.Text = RECreader.Item("开关类型").ToString
                        Text75.Text = RECreader.Item("型号").ToString
                        Text92.Text = RECreader.Item("压力等级MPa").ToString
                        Text25.Text = RECreader.Item("重量kg").ToString
                        Text76.Text = RECreader.Item("温度范围下℃").ToString
                        Text88.Text = RECreader.Item("温度范围上℃").ToString
                        Text29.Text = RECreader.Item("极限压差MPa").ToString
                        Text27.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text28.Text = RECreader.Item("抗外挤强度MPa").ToString
                        Text26.Text = RECreader.Item("抗拉强度kN").ToString
                        Text89.Text = RECreader.Item("上端扣型").ToString
                        Text90.Text = RECreader.Item("下端扣型").ToString
                        Text93.Text = RECreader.Item("生产厂家").ToString
                        Text91.Text = RECreader.Item("备注").ToString
                    End If
                Case "伸缩管"
                    SQL_command = "select * from 管柱_伸缩元件 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text49.Text = RECreader.Item("零件号").ToString
                        Text47.Text = RECreader.Item("生产厂家").ToString
                        Text67.Text = RECreader.Item("上端扣型").ToString
                        Text68.Text = RECreader.Item("下端扣型").ToString
                        Text60.Text = RECreader.Item("温度范围下℃").ToString
                        Text61.Text = RECreader.Item("温度范围上℃").ToString
                        Text69.Text = RECreader.Item("主体材料").ToString
                        Text59.Text = RECreader.Item("备注").ToString
                        Text62.Text = RECreader.Item("主材屈服强度MPa").ToString
                        Text64.Text = RECreader.Item("压力等级MPa").ToString
                        Text46.Text = RECreader.Item("重量kg").ToString
                        Text66.Text = RECreader.Item("全缩短长度m").ToString
                        Text48.Text = RECreader.Item("伸缩行程m").ToString
                        Text65.Text = RECreader.Item("抗拉强度kN").ToString
                        Text63.Text = RECreader.Item("伸缩动作载荷kN").ToString
                        Text50.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text51.Text = RECreader.Item("抗外挤强度MPa").ToString
                        TextBox3.Text = RECreader.Item("伸缩动作压力MPa").ToString
                    End If
                Case "锚定工具"
                    SQL_command = "select * from 管柱_锚定元件 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text79.Text = RECreader.Item("型号").ToString
                        Text85.Text = RECreader.Item("生产厂家").ToString
                        Combo2.Text = RECreader.Item("坐卡方式").ToString
                        Text86.Text = RECreader.Item("小坐卡压力MPa").ToString
                        Text72.Text = RECreader.Item("大坐卡压力MPa").ToString
                        Text74.Text = RECreader.Item("温度范围下℃").ToString
                        Text73.Text = RECreader.Item("温度范围上℃").ToString
                        Text80.Text = RECreader.Item("重量kg").ToString
                        Text78.Text = RECreader.Item("上端扣型").ToString
                        Text77.Text = RECreader.Item("下端扣型").ToString
                        Text81.Text = RECreader.Item("工作压力MPa").ToString
                        Text84.Text = RECreader.Item("最大锚定力kN").ToString
                        Text83.Text = RECreader.Item("最大解锚力kN").ToString
                        Text87.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text71.Text = RECreader.Item("抗外压强度MPa").ToString
                        Text82.Text = RECreader.Item("抗拉强度kN").ToString
                        Text70.Text = RECreader.Item("备注").ToString
                    End If
                Case "普通钻杆"
                    SQL_command = "select * from 管柱_普通钻杆 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    If RECreader.Read Then
                        Text108.Text = RECreader.Item("钻杆壁厚mm").ToString
                        Text105.Text = RECreader.Item("管体抗拉强度kN").ToString
                        Text107.Text = RECreader.Item("单根长度m").ToString
                        Text111.Text = RECreader.Item("单位长度质量kgpm").ToString
                        Text112.Text = RECreader.Item("加厚型式").ToString
                        Text106.Text = RECreader.Item("备注").ToString
                        Text104.Text = RECreader.Item("接头外径mm").ToString
                        Text103.Text = RECreader.Item("屈服强度MPa").ToString
                        Text102.Text = RECreader.Item("扣型").ToString
                        Text101.Text = RECreader.Item("钢级").ToString
                        Text100.Text = RECreader.Item("管体抗扭强度Nm").ToString
                        Text99.Text = RECreader.Item("接头抗扭强度Nm").ToString
                        Text98.Text = RECreader.Item("抗内压强度MPa").ToString
                        Text97.Text = RECreader.Item("抗挤强度MPa").ToString
                        Text96.Text = RECreader.Item("弹性模量MPa").ToString
                        Text95.Text = RECreader.Item("泊松比").ToString
                        Text94.Text = RECreader.Item("接头抗拉强度kN").ToString
                        TextBox2.Text = RECreader.Item("热膨胀系数").ToString
                    End If
                    '李润洲  2025年6月增加
                Case "射孔枪"
                    SQL_command = "select * from 管柱_射孔枪弹表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text
                    Using oleDbCmd As New OleDbCommand(SQL_command, cn_userdb)
                        Using reader As OleDbDataReader = oleDbCmd.ExecuteReader()
                            If reader.Read Then
                                Texts1.Text = reader.Item("射孔器分类").ToString
                                Texts2.Text = reader.Item("壁厚mm").ToString
                                Texts5.Text = reader.Item("耐压MPa").ToString
                                Texts6.Text = reader.Item("射孔密度孔╱m").ToString
                                Texts7.Text = reader.Item("射孔相位°").ToString
                                Texts9.Text = reader.Item("射孔弹名称").ToString
                                Texts20.Text = reader.Item("炸药名称").ToString
                                Texts11.Text = reader.Item("装药量g").ToString
                                Texts12.Text = reader.Item("装药密度g╱cm3").ToString
                                Texts13.Text = reader.Item("耐温°C").ToString
                                Texts14.Text = reader.Item("平均孔径mm").ToString
                                Texts15.Text = reader.Item("平均穿深mm").ToString
                                Texts19.Text = reader.Item("备注").ToString
                            End If
                        End Using
                    End Using
                    '李润洲  2025年7与25日增加
                Case "筛管"
                    SQL_command = "select * from 管柱_筛管表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text

                    Using oleDbCmd As New OleDbCommand(SQL_command, cn_userdb)
                        Using RECreader1 As OleDbDataReader = oleDbCmd.ExecuteReader()
                            If RECreader1.Read Then
                                Textk1.Text = RECreader1.Item("主体材料").ToString
                                Textk2.Text = RECreader1.Item("单根长度m").ToString
                                Textk3.Text = RECreader1.Item("单根质量kg").ToString
                                Textk4.Text = RECreader1.Item("孔隙率%").ToString
                                Textk5.Text = RECreader1.Item("有效流通面积cm2╱m").ToString
                                Textk6.Text = RECreader1.Item("耐压差MPa").ToString
                                Textk7.Text = RECreader1.Item("耐温℃").ToString
                                Textk8.Text = RECreader1.Item("抗内压强度MPa").ToString
                                Textk9.Text = RECreader1.Item("抗外挤强度MPa").ToString
                                Textk10.Text = RECreader1.Item("主材屈服强度MPa").ToString
                                Textk11.Text = RECreader1.Item("拦截的最小颗粒直径mm").ToString
                                Textk12.Text = RECreader1.Item("生产厂家").ToString
                                Textk13.Text = RECreader1.Item("接头抗拉伸强度kN").ToString
                                Textk14.Text = RECreader1.Item("备注").ToString
                                Textk15.Text = RECreader1.Item("壁厚mm").ToString

                            End If
                        End Using
                    End Using
            End Select
            xiashen = 0.0#
            zongchang = 0.0#
            SQL_command = "select sum(元件长度m) as 下入长度  from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号<=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            If RECreader.Read Then
                xiashen = Val(RECreader.Item("下入长度").ToString)
            End If
            SQL_command = "select sum(元件长度m) as 总长度  from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            If RECreader.Read Then
                zongchang = Val(RECreader.Item("总长度").ToString)
            End If
            RECreader.Close()
            cn_userdb.Close()
            Label_25.Text = "下入到" & CStr(xiashen) & "米，管柱总长" & CStr(zongchang) & "米。"
            If Val(Text5.Text) < zongchang Then
                msg_prompt = "管柱比完钻井深长，数据不合理！"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用油井管"按钮
    '*********************************************************************************************************************************************
    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
        Call myrefresh2()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView2单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Call myrefresh2()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的油井管列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh2()
        If IsNothing(Me.BindingSource2.Current) Then
            msg_prompt = "请从列表中选择油井管。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource2.Current("油管规格")) Then
                msg_prompt = "请从列表中选择油井管。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource2.Current("油管规格").ToString
                Text10.Text = IIf(BindingSource2.Current("油管外径mm").ToString = "", "0.0", BindingSource2.Current("油管外径mm").ToString)
                Text13.Text = IIf(BindingSource2.Current("油管壁厚mm").ToString = "", "0.0", BindingSource2.Current("油管壁厚mm").ToString)
                Text11.Text = CStr(Val(Text10.Text) - Val(Text13.Text) - Val(Text13.Text))
                Text14.Text = BindingSource2.Current("材料").ToString
                Text15.Text = IIf(BindingSource2.Current("单位长度质量kg/m").ToString = "", "0.0", BindingSource2.Current("单位长度质量kg/m").ToString)
                Text16.Text = IIf(BindingSource2.Current("抗拉强度kN").ToString = "", "0.0", BindingSource2.Current("抗拉强度kN").ToString)
                Text17.Text = IIf(BindingSource2.Current("屈服应力MPa").ToString = "", "0.0", BindingSource2.Current("屈服应力MPa").ToString)
                Text18.Text = IIf(BindingSource2.Current("弹性模量MPa").ToString = "", "0.0", BindingSource2.Current("弹性模量MPa").ToString)
                Text19.Text = IIf(BindingSource2.Current("泊松比").ToString = "", "0.0", BindingSource2.Current("泊松比").ToString)
                Text20.Text = IIf(BindingSource2.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource2.Current("抗内压强度MPa").ToString)
                Text21.Text = IIf(BindingSource2.Current("抗外挤强度MPa").ToString = "", "0.0", BindingSource2.Current("抗外挤强度MPa").ToString)
                Text110.Text = BindingSource2.Current("扣型").ToString
                Text109.Text = BindingSource2.Current("备注").ToString
                Text113.Text = IIf(BindingSource2.Current("接头抗内压强度MPa").ToString = "", "0.0", BindingSource2.Current("接头抗内压强度MPa").ToString)
                Text114.Text = IIf(BindingSource2.Current("接头抗拉强度kN").ToString = "", "0.0", BindingSource2.Current("接头抗拉强度kN").ToString)
                TextBox1.Text = IIf(BindingSource2.Current("热膨胀系数").ToString = "", "0.0", BindingSource2.Current("热膨胀系数").ToString)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用节流工具"按钮
    '*********************************************************************************************************************************************
    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
        Call myrefresh3()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView3单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView3_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        Call myrefresh3()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的节流工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh3()
        If IsNothing(Me.BindingSource3.Current) Then
            msg_prompt = "请从列表中选择节流工具。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource3.Current("名称")) Then
                msg_prompt = "请从列表中选择节流工具。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource3.Current("名称").ToString
                Text9.Text = IIf(BindingSource3.Current("长度(m)").ToString = "", "0.0", BindingSource3.Current("长度(m)").ToString)
                Text10.Text = IIf(BindingSource3.Current("外径(mm)").ToString = "", "0.0", BindingSource3.Current("外径(mm)").ToString)
                Text11.Text = IIf(BindingSource3.Current("内径(mm)").ToString = "", "0.0", BindingSource3.Current("内径(mm)").ToString)
                Text31.Text = IIf(BindingSource3.Current("节流孔内径(mm)").ToString = "", "0.0", BindingSource3.Current("节流孔内径(mm)").ToString)
                Text32.Text = IIf(BindingSource3.Current("节流孔长度(mm)").ToString = "", "0.0", BindingSource3.Current("节流孔长度(mm)").ToString)
                Text41.Text = IIf(BindingSource3.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource3.Current("抗外挤强度(MPa)").ToString)
                Text42.Text = IIf(BindingSource3.Current("重量(kg)").ToString = "", "0.0", BindingSource3.Current("重量(kg)").ToString)
                Text43.Text = IIf(BindingSource3.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource3.Current("抗内压强度(MPa)").ToString)
                Text44.Text = IIf(BindingSource3.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource3.Current("抗拉强度(kN)").ToString)
                Text45.Text = BindingSource3.Current("节流流向").ToString
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用封隔器"按钮
    '*********************************************************************************************************************************************
    Private Sub Command15_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command15.Click
        Call myrefresh4()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView4单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView4_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellDoubleClick
        Call myrefresh4()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的封隔器列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh4()
        If IsNothing(Me.BindingSource4.Current) Then
            msg_prompt = "请从列表中选择封隔器。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource4.Current("封隔器名称")) Then
                msg_prompt = "请从列表中选择封隔器。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource4.Current("封隔器名称").ToString
                Text9.Text = IIf(BindingSource4.Current("长度(m)").ToString = "", "0.0", BindingSource4.Current("长度(m)").ToString)
                Text10.Text = IIf(BindingSource4.Current("最大外径(mm)").ToString = "", "0.0", BindingSource4.Current("最大外径(mm)").ToString)
                Text11.Text = IIf(BindingSource4.Current("最小通径(mm)").ToString = "", "0.0", BindingSource4.Current("最小通径(mm)").ToString)
                Text22.Text = BindingSource4.Current("型号").ToString
                Text35.Text = BindingSource4.Current("生产厂家").ToString
                Combo1.Text = BindingSource4.Current("坐封方式").ToString
                Text54.Text = IIf(BindingSource4.Current("温度范围下℃").ToString = "", "0.0", BindingSource4.Current("温度范围下℃").ToString)
                Text55.Text = IIf(BindingSource4.Current("温度范围上℃").ToString = "", "0.0", BindingSource4.Current("温度范围上℃").ToString)
                Text52.Text = BindingSource4.Current("主体材料").ToString
                Text58.Text = BindingSource4.Current("备注").ToString
                Text53.Text = IIf(BindingSource4.Current("主材屈服强度MPa").ToString = "", "0.0", BindingSource4.Current("主材屈服强度MPa").ToString)
                Text23.Text = BindingSource4.Current("上端扣型").ToString
                Text24.Text = BindingSource4.Current("下端扣型").ToString
                Text39.Text = IIf(BindingSource4.Current("极限承压(MPa)").ToString = "", "0.0", BindingSource4.Current("极限承压(MPa)").ToString)
                Text40.Text = IIf(BindingSource4.Current("重量kg").ToString = "", "0.0", BindingSource4.Current("重量kg").ToString)
                Text36.Text = IIf(BindingSource4.Current("坐封力(kN)").ToString = "", "0.0", BindingSource4.Current("坐封力(kN)").ToString)
                Text37.Text = IIf(BindingSource4.Current("坐封压差(MPa)").ToString = "", "0.0", BindingSource4.Current("坐封压差(MPa)").ToString)
                Text34.Text = IIf(BindingSource4.Current("最小坐封压力MPa").ToString = "", "0.0", BindingSource4.Current("最小坐封压力MPa").ToString)
                Text56.Text = IIf(BindingSource4.Current("最大坐封压力MPa").ToString = "", "0.0", BindingSource4.Current("最大坐封压力MPa").ToString)
                Text33.Text = IIf(BindingSource4.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource4.Current("抗内压强度MPa").ToString)
                Text57.Text = IIf(BindingSource4.Current("抗外压强度MPa").ToString = "", "0.0", BindingSource4.Current("抗外压强度MPa").ToString)
                Text38.Text = IIf(BindingSource4.Current("极限载荷(kN)").ToString = "", "0.0", BindingSource4.Current("极限载荷(kN)").ToString)
                Combo3.Text = BindingSource4.Current("能否反洗井").ToString
                Textf2.Text = IIf(BindingSource4.Current("中心管外径(mm)").ToString = "", "0.0", BindingSource4.Current("中心管外径(mm)").ToString)
                Textf1.Text = IIf(BindingSource4.Current("中心管内径(mm)").ToString = "", "0.0", BindingSource4.Current("中心管内径(mm)").ToString)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用开关工具"按钮
    '*********************************************************************************************************************************************
    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
        Call myrefresh5()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView5单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView5_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView5.CellDoubleClick
        Call myrefresh5()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的开关工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh5()
        If IsNothing(Me.BindingSource5.Current) Then
            msg_prompt = "请从列表中选择开关工具。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource5.Current("名称")) Then
                msg_prompt = "请从列表中选择开关工具。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource5.Current("名称").ToString
                Text9.Text = IIf(BindingSource5.Current("长度(m)").ToString = "", "0.0", BindingSource5.Current("长度(m)").ToString)
                Text10.Text = IIf(BindingSource5.Current("外径(mm)").ToString = "", "0.0", BindingSource5.Current("外径(mm)").ToString)
                Text11.Text = IIf(BindingSource5.Current("内径(mm)").ToString = "", "0.0", BindingSource5.Current("内径(mm)").ToString)
                Text30.Text = BindingSource5.Current("开关类型").ToString
                Text75.Text = BindingSource5.Current("型号").ToString
                Text92.Text = IIf(BindingSource5.Current("压力等级MPa").ToString = "", "0.0", BindingSource5.Current("压力等级MPa").ToString)
                Text25.Text = IIf(BindingSource5.Current("重量(kg)").ToString = "", "0.0", BindingSource5.Current("重量(kg)").ToString)
                Text76.Text = IIf(BindingSource5.Current("温度范围下℃").ToString = "", "0.0", BindingSource5.Current("温度范围下℃").ToString)
                Text88.Text = IIf(BindingSource5.Current("温度范围上℃").ToString = "", "0.0", BindingSource5.Current("温度范围上℃").ToString)
                Text29.Text = IIf(BindingSource5.Current("极限压差(MPa)").ToString = "", "0.0", BindingSource5.Current("极限压差(MPa)").ToString)
                Text27.Text = IIf(BindingSource5.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource5.Current("抗内压强度(MPa)").ToString)
                Text28.Text = IIf(BindingSource5.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource5.Current("抗外挤强度(MPa)").ToString)
                Text26.Text = IIf(BindingSource5.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource5.Current("抗拉强度(kN)").ToString)
                Text89.Text = BindingSource5.Current("上端扣型").ToString
                Text90.Text = BindingSource5.Current("下端扣型").ToString
                Text93.Text = BindingSource5.Current("生产厂家").ToString
                Text91.Text = BindingSource5.Current("备注").ToString
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用伸缩管"按钮
    '*********************************************************************************************************************************************
    Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
        Call myrefresh6()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView6单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView6_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellDoubleClick
        Call myrefresh6()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的伸缩管列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh6()
        If IsNothing(Me.BindingSource6.Current) Then
            msg_prompt = "请从列表中选择伸缩管。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource6.Current("元件名称")) Then
                msg_prompt = "请从列表中选择伸缩管。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource6.Current("元件名称").ToString
                Text9.Text = IIf(BindingSource6.Current("全缩短长度m").ToString = "", "0.0", BindingSource6.Current("全缩短长度m").ToString)
                Text10.Text = IIf(BindingSource6.Current("外径(mm)").ToString = "", "0.0", BindingSource6.Current("外径(mm)").ToString)
                Text11.Text = IIf(BindingSource6.Current("内径(mm)").ToString = "", "0.0", BindingSource6.Current("内径(mm)").ToString)
                Text49.Text = BindingSource6.Current("零件号").ToString
                Text47.Text = BindingSource6.Current("生产厂家").ToString
                Text67.Text = BindingSource6.Current("上端扣型").ToString
                Text68.Text = BindingSource6.Current("下端扣型").ToString
                Text60.Text = IIf(BindingSource6.Current("温度范围下℃").ToString = "", "0.0", BindingSource6.Current("温度范围下℃").ToString)
                Text61.Text = IIf(BindingSource6.Current("温度范围上℃").ToString = "", "0.0", BindingSource6.Current("温度范围上℃").ToString)
                Text69.Text = BindingSource6.Current("主体材料").ToString
                Text59.Text = BindingSource6.Current("备注").ToString
                Text62.Text = IIf(BindingSource6.Current("主材屈服强度MPa").ToString = "", "0.0", BindingSource6.Current("主材屈服强度MPa").ToString)
                Text64.Text = IIf(BindingSource6.Current("压力等级MPa").ToString = "", "0.0", BindingSource6.Current("压力等级MPa").ToString)
                Text46.Text = IIf(BindingSource6.Current("重量(kg)").ToString = "", "0.0", BindingSource6.Current("重量(kg)").ToString)
                Text66.Text = IIf(BindingSource6.Current("全缩短长度m").ToString = "", "0.0", BindingSource6.Current("全缩短长度m").ToString)
                Text48.Text = IIf(BindingSource6.Current("伸缩行程(m)").ToString = "", "0.0", BindingSource6.Current("伸缩行程(m)").ToString)
                Text65.Text = IIf(BindingSource6.Current("抗拉强度kN").ToString = "", "0.0", BindingSource6.Current("抗拉强度kN").ToString)
                Text50.Text = IIf(BindingSource6.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource6.Current("抗内压强度(MPa)").ToString)
                Text51.Text = IIf(BindingSource6.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource6.Current("抗外挤强度(MPa)").ToString)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用锚定工具"按钮
    '*********************************************************************************************************************************************
    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
        Call myrefresh7()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView7单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView7_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellDoubleClick
        Call myrefresh7()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的锚定工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh7()
        If IsNothing(Me.BindingSource7.Current) Then
            msg_prompt = "请从列表中选择锚定工具。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource7.Current("工具名称")) Then
                msg_prompt = "请从列表中选择锚定工具。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource7.Current("工具名称").ToString
                Text9.Text = IIf(BindingSource7.Current("总体长度m").ToString = "", "0.0", BindingSource7.Current("总体长度m").ToString)
                Text10.Text = IIf(BindingSource7.Current("最大外径mm").ToString = "", "0.0", BindingSource7.Current("最大外径mm").ToString)
                Text11.Text = IIf(BindingSource7.Current("最小通径mm").ToString = "", "0.0", BindingSource7.Current("最小通径mm").ToString)
                Text79.Text = BindingSource7.Current("型号").ToString
                Combo2.Text = BindingSource7.Current("坐卡方式").ToString
                Text80.Text = IIf(BindingSource7.Current("重量kg").ToString = "", "0.0", BindingSource7.Current("重量kg").ToString)    '表示质量，存的就是质量，不转换。表设计之初的问题。
                Text85.Text = BindingSource7.Current("生产厂家").ToString
                Text74.Text = IIf(BindingSource7.Current("温度范围下℃").ToString = "", "0.0", BindingSource7.Current("温度范围下℃").ToString)
                Text73.Text = IIf(BindingSource7.Current("温度范围上℃").ToString = "", "0.0", BindingSource7.Current("温度范围上℃").ToString)
                Text78.Text = BindingSource7.Current("上端扣型").ToString
                Text86.Text = IIf(BindingSource7.Current("小坐卡压力MPa").ToString = "", "0.0", BindingSource7.Current("小坐卡压力MPa").ToString)
                Text77.Text = BindingSource7.Current("下端扣型").ToString
                Text72.Text = IIf(BindingSource7.Current("大坐卡压力MPa").ToString = "", "0.0", BindingSource7.Current("大坐卡压力MPa").ToString)
                Text84.Text = IIf(BindingSource7.Current("最大锚定力kN").ToString = "", "0.0", BindingSource7.Current("最大锚定力kN").ToString)
                Text83.Text = IIf(BindingSource7.Current("最大解锚力kN").ToString = "", "0.0", BindingSource7.Current("最大解锚力kN").ToString)
                Text81.Text = IIf(BindingSource7.Current("工作压力MPa").ToString = "", "0.0", BindingSource7.Current("工作压力MPa").ToString)
                Text87.Text = IIf(BindingSource7.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource7.Current("抗内压强度MPa").ToString)
                Text71.Text = IIf(BindingSource7.Current("抗外压强度MPa").ToString = "", "0.0", BindingSource7.Current("抗外压强度MPa").ToString)
                Text82.Text = IIf(BindingSource7.Current("抗拉强度kN").ToString = "", "0.0", BindingSource7.Current("抗拉强度kN").ToString)
                Text70.Text = BindingSource7.Current("型号").ToString
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"↑选用普通钻杆"按钮
    '*********************************************************************************************************************************************
    Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command10.Click
        Call myrefresh8()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView8单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView8_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView8.CellDoubleClick
        Call myrefresh8()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的普通钻杆列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh8()
        If IsNothing(Me.BindingSource8.Current) Then
            msg_prompt = "请从列表中选择普通钻杆。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource8.Current("钻杆规格")) Then
                msg_prompt = "请从列表中选择普通钻杆。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text7.Text = BindingSource8.Current("钻杆规格").ToString
                Text10.Text = IIf(BindingSource8.Current("钻杆外径(mm)").ToString = "", "0.0", BindingSource8.Current("钻杆外径(mm)").ToString)
                Text108.Text = IIf(BindingSource8.Current("钻杆壁厚(mm)").ToString = "", "0.0", BindingSource8.Current("钻杆壁厚(mm)").ToString)
                Text11.Text = CStr(Val(Text10.Text) - Val(Text108.Text) - Val(Text108.Text))
                Text105.Text = IIf(BindingSource8.Current("管体抗拉强度kN").ToString = "", "0.0", BindingSource8.Current("管体抗拉强度kN").ToString)
                Text107.Text = IIf(BindingSource8.Current("单根长度(m)").ToString = "", "0.0", BindingSource8.Current("单根长度(m)").ToString)
                Text111.Text = IIf(BindingSource8.Current("单位长度质量(Kg/m)").ToString = "", "0.0", BindingSource8.Current("单位长度质量(Kg/m)").ToString)
                Text112.Text = BindingSource8.Current("加厚型式").ToString
                Text106.Text = BindingSource8.Current("备注").ToString
                Text104.Text = IIf(BindingSource8.Current("接头外径mm").ToString = "", "0.0", BindingSource8.Current("接头外径mm").ToString)
                Text103.Text = IIf(BindingSource8.Current("屈服强度MPa").ToString = "", "0.0", BindingSource8.Current("屈服强度MPa").ToString)
                Text102.Text = BindingSource8.Current("扣型").ToString
                Text101.Text = BindingSource8.Current("钢级").ToString
                Text100.Text = IIf(BindingSource8.Current("管体抗扭强度Nm").ToString = "", "0.0", BindingSource8.Current("管体抗扭强度Nm").ToString)
                Text99.Text = IIf(BindingSource8.Current("接头抗扭强度Nm").ToString = "", "0.0", BindingSource8.Current("接头抗扭强度Nm").ToString)
                Text98.Text = IIf(BindingSource8.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource8.Current("抗内压强度MPa").ToString)
                Text97.Text = IIf(BindingSource8.Current("抗挤强度MPa").ToString = "", "0.0", BindingSource8.Current("抗挤强度MPa").ToString)
                Text96.Text = IIf(BindingSource8.Current("弹性模量MPa").ToString = "", "0.0", BindingSource8.Current("弹性模量MPa").ToString)
                Text95.Text = IIf(BindingSource8.Current("泊松比").ToString = "", "0.0", BindingSource8.Current("泊松比").ToString)
                Text94.Text = IIf(BindingSource8.Current("接头抗拉强度kN").ToString = "", "0.0", BindingSource8.Current("接头抗拉强度kN").ToString)
                TextBox2.Text = IIf(BindingSource8.Current("热膨胀系数").ToString = "", "0.0", BindingSource8.Current("热膨胀系数").ToString)
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
    '*********************************************************************************************************************************************
    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        On Error GoTo errhandler
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的管柱数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn_userdb.Close()
                Exit Sub
            End If
            SQL_command = "delete * from 管柱_油管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_普通钻杆 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '*********************************************************************************************************************************************
            '删除管柱-射孔枪数据，李润洲2025年7月1日增加
            '*********************************************************************************************************************************************
            SQL_command = "delete * from 管柱_射孔枪弹表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '*********************************************************************************************************************************************
            '*********************************************************************************************************************************************
            '删除管柱-射孔枪数据，李润洲2025年7月24日增加
            '*********************************************************************************************************************************************
            SQL_command = "delete * from 管柱_筛管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '2025.8.2 李润洲增加
            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            'SQL_command = "delete * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            'EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段爆轰计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段封隔器计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔油管应力计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '*********************************************************************************************************************************************
            SQL_command = "DELETE * from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn_userdb.Close()
            Call fill_gzhgrid()
            Call drawWellStruction(Picture1, 2)
            msg_prompt = "管柱数据删除完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            RECreader.Close()
            cn_userdb.Close()
            msg_prompt = "未找到元件序号对应的数据，请选好要删除的数据！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "元件序号数据错误，请选好要删除的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        On Error GoTo errhandler
        If Text7.Text = "" Then
            msg_prompt = "请输入元件名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) <= 0 Or (Not IsNumeric(Text8.Text)) Then
            msg_prompt = "请输入合适的元件序号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text9.Text) <= 0 Or (Not IsNumeric(Text9.Text)) Then
            msg_prompt = "请输入合适的元件长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text12.Text = "" Then
            msg_prompt = "请选择元件类型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text10.Text) <= 0 Or (Not IsNumeric(Text10.Text)) Then
            msg_prompt = "请输入合适的元件外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text11.Text) <= 0 Or (Not IsNumeric(Text11.Text)) Then
            msg_prompt = "请输入合适的元件内径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '无枪身射孔枪没有明确的整体壁厚参数，所以允许其内径=外径，允许壁厚为0
        '李润洲2025年7月25日增加判断分支
        If Text12.Text = "射孔枪" Then
            If Val(Text11.Text) < 0 Or (Not IsNumeric(Text11.Text)) Then
                msg_prompt = "请输入合适的元件内径。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
            If Val(Text10.Text) < Val(Text11.Text) Then
                msg_prompt = "外径小于内径，不合理，请输入合适的元件外径。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
        Else
            If Val(Text11.Text) <= 0 Or (Not IsNumeric(Text11.Text)) Then
                msg_prompt = "请输入合适的元件内径。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
            If Val(Text10.Text) <= Val(Text11.Text) Then
                msg_prompt = "外径小于等于内径，不合理，请输入合适的元件外径。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
        End If
        Select Case Text12.Text
            Case "油井管"
                If Text14.Text = "" Then
                    msg_prompt = "请输入或选择油井管材料的钢级。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Trim(Text110.Text) = "" Then
                    msg_prompt = "请输入或选择油井管接头的扣型。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text15.Text) <= 0 Or (Not IsNumeric(Text15.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管单位长质量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text17.Text) <= 0 Or (Not IsNumeric(Text17.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管屈服强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text21.Text) <= 0 Or (Not IsNumeric(Text21.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管抗外挤强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text20.Text) <= 0 Or (Not IsNumeric(Text20.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管抗内压强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text16.Text) <= 0 Or (Not IsNumeric(Text16.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管管体抗拉强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(TextBox1.Text) <= 0 Or (Not IsNumeric(TextBox1.Text)) Then
                    msg_prompt = "请输入或选择合适的油井管热膨胀系数。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "节流工具"
                If Val(Text31.Text) = 0 Then
                    msg_prompt = "请输入或选择节流孔内径。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text32.Text) = 0 Then
                    msg_prompt = "请输入或选择节流孔长度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text42.Text) = 0 Then
                    msg_prompt = "请输入或选择节流元件重量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text45.Text = "" Then
                    msg_prompt = "请输入或选择节流流向。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "封隔器"
                If Val(Text40.Text) <= 0 Or (Not IsNumeric(Text40.Text)) Then
                    msg_prompt = "请输入或选择合适的封隔器重量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text22.Text = "" Then
                    msg_prompt = "请输入或选择合适的封隔器型号。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not (Combo1.Text = "旋转下放" Or Combo1.Text = "投球打压" Or Combo1.Text = "上提下放" Or Combo1.Text = "水力扩张") Then
                    msg_prompt = "请选择合适的封隔器坐封方式。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Textf2.Text) <= 0 Or (Not IsNumeric(Textf2.Text)) Then
                    msg_prompt = "请输入或选择合适的封隔器中心管外径。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Textf1.Text) <= 0 Or (Not IsNumeric(Textf1.Text)) Then
                    msg_prompt = "请输入或选择合适的封隔器中心管内径。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "开关工具"
                If Val(Text25.Text) <= 0 Or (Not IsNumeric(Text25.Text)) Then
                    msg_prompt = "请输入或选择合适的开关元件重量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Trim(Text30.Text) = "" Then
                    msg_prompt = "请输入或选择开关工具类型。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Trim(Text75.Text) = "" Then
                    msg_prompt = "请输入开关工具型号。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "伸缩管"
                If Val(Text46.Text) <= 0 Or (Not IsNumeric(Text46.Text)) Then
                    msg_prompt = "请输入合适的伸缩元件重量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text48.Text) <= 0 Or (Not IsNumeric(Text48.Text)) Then
                    msg_prompt = "请输入合适的伸缩元件伸缩行程。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not (Val(Text9.Text) >= Val(Text66.Text) And Val(Text9.Text) <= (Val(Text66.Text) + Val(Text48.Text))) Then
                    msg_prompt = "对于伸缩元件，元件长度(m)为下入长度，须满足：全缩短长度<=元件长度<=（全缩短长度+伸缩行程），请输入合适的元件长度(m)值。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text63.Text) < 0 Or (Not IsNumeric(Text63.Text)) Then
                    msg_prompt = "请输入合适的伸缩动作载荷（让销钉剪断的轴向力）。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(TextBox3.Text) < 0 Or (Not IsNumeric(TextBox3.Text)) Then
                    msg_prompt = "请输入合适的伸缩动作压力（让销钉剪断的内外压力差）。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text63.Text) <> 0 And Val(TextBox3.Text) = 0 Then
                    msg_prompt = "   伸缩动作载荷不为零，说明伸缩管装有销钉，那么让销钉剪断的内外压力差是不会为零的，请输入合适的伸缩动作压力。" & Chr(13) & Chr(10) _
                    & "说明：伸缩管产品说明书中一般有剪断单颗销钉的剪切力和剪切压力，剪切力除以剪切压力为活塞面积，此应为定值。建议伸缩动作载荷除以伸缩动作压力所得值与单颗销钉算得的值一致。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Val(Text63.Text) = 0 And Val(TextBox3.Text) <> 0 Then
                    msg_prompt = "伸缩动作压力不为零，说明伸缩管装有销钉，那么让销钉剪断的轴向力是不会为零的，请输入合适的伸缩动作载荷。" & Chr(13) & Chr(10) _
                    & "说明：伸缩管产品说明书中一般有剪断单颗销钉的剪切力和剪切压力，剪切力除以剪切压力为活塞面积，此应为定值。建议伸缩动作载荷除以伸缩动作压力所得值与单颗销钉算得的值一致。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "锚定工具"
                If Val(Text80.Text) <= 0 Or (Not IsNumeric(Text80.Text)) Then
                    msg_prompt = "请输入或选择合适的锚定工具重量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text79.Text = "" Then
                    msg_prompt = "请输入或选择合适的锚定工具型号。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not (Combo2.Text = "油管打压" Or Combo2.Text = "内外压差" Or Combo2.Text = "下放管柱") Then
                    msg_prompt = "请选择合适的锚定工具坐卡方式。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
            Case "普通钻杆"
                If Text7.Text = "" Then
                    Text7.Focus()
                    msg_prompt = "请输入或选择钻杆规格。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text10.Text) Or IsDBNull(Text10.Text) Or Val(Text10.Text) = 0 Then
                    Text10.Focus()
                    msg_prompt = "请输入合法的钻杆外径。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text108.Text) Or IsDBNull(Text108.Text) Or Val(Text108.Text) = 0 Then
                    Text108.Focus()
                    msg_prompt = "请输入合法的钻杆壁厚。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text107.Text) Or IsDBNull(Text107.Text) Or Val(Text107.Text) = 0 Then
                    Text107.Focus()
                    msg_prompt = "请输入合法的单根长度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text104.Text) Or IsDBNull(Text104.Text) Or Val(Text104.Text) = 0 Then
                    Text104.Focus()
                    msg_prompt = "请输入合法的接头外径。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text112.Text = "" Then
                    msg_prompt = "请输入加厚型式。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text111.Text) Or IsDBNull(Text111.Text) Or Val(Text111.Text) = 0 Then
                    Text111.Focus()
                    msg_prompt = "请输入合法的单位长度质量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text102.Text = "" Then
                    msg_prompt = "请输入扣型。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Text101.Text = "" Then
                    msg_prompt = "请输入钢级。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text103.Text) Or IsDBNull(Text103.Text) Or Val(Text103.Text) = 0 Then
                    Text103.Focus()
                    msg_prompt = "请输入合法的屈服强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text105.Text) Or IsDBNull(Text105.Text) Or Val(Text105.Text) = 0 Then
                    Text105.Focus()
                    msg_prompt = "请输入合法的管体抗拉强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text94.Text) Or IsDBNull(Text94.Text) Or Val(Text94.Text) = 0 Then
                    Text94.Focus()
                    msg_prompt = "请输入合法的接头抗拉强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text98.Text) Or IsDBNull(Text98.Text) Or Val(Text98.Text) = 0 Then
                    Text98.Focus()
                    msg_prompt = "请输入合法的抗内压强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text96.Text) Or IsDBNull(Text96.Text) Or Val(Text96.Text) = 0 Then
                    Text96.Focus()
                    msg_prompt = "请输入合法的弹性模量。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text97.Text) Or IsDBNull(Text97.Text) Or Val(Text97.Text) = 0 Then
                    Text97.Focus()
                    msg_prompt = "请输入合法的抗挤强度。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(Text95.Text) Or IsDBNull(Text95.Text) Or Val(Text95.Text) = 0 Then
                    Text95.Focus()
                    msg_prompt = "请输入合法的泊松比。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                If Not IsNumeric(TextBox2.Text) Or IsDBNull(TextBox2.Text) Or Val(TextBox2.Text) = 0 Then
                    Text95.Focus()
                    msg_prompt = "请输入合法的热膨胀系数。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    Exit Sub
                End If
                '李润洲2025年6月增加
            Case "射孔枪"
                '无枪身射孔枪大多时候没有壁厚参数，所以壁厚允许为0，该字段只检查是否是数字，重载CheckWigetData
                If Not CheckWigetData(Texts2, "壁厚") Then Exit Sub '不是数字时退出
                If Not CheckWigetData(Text10, "射孔枪类型", False) Then Exit Sub '为空退出
                If Not CheckWigetData(Texts20, "炸药名称", False) Then Exit Sub '为空退出
                If Not CheckWigetData(Texts9, "射孔弹名称", False) Then Exit Sub '为空退出
                If Not CheckWigetData(Texts5, "耐压", True, False) Then Exit Sub '为空、不是数字、小于等于0时退出
                If Not CheckWigetData(Texts6, "射孔密度", True, False) Then Exit Sub
                If Not CheckWigetData(Texts11, "装药量", True, False) Then Exit Sub
                If Not CheckWigetData(Texts12, "装药密度", True, False) Then Exit Sub
                If Not CheckWigetData(Texts7, "射孔相位", True, False) Then Exit Sub
                If Not CheckWigetData(Texts13, "耐温", True, False) Then Exit Sub
                If Not CheckWigetData(Texts14, "平均孔径", True, False) Then Exit Sub
                If Not CheckWigetData(Texts15, "平均穿深", True, False) Then Exit Sub
                '李润洲2025年7月25日增加
            Case "筛管"
                If Not CheckWigetData(Textk15, "壁厚", True, False) Then Exit Sub
                If Not CheckWigetData(Textk1, "主体材料", False) Then Exit Sub
                If Not CheckWigetData(Textk2, "单根长度", True, False) Then Exit Sub
                If Not CheckWigetData(Textk3, "单根质量", True, False) Then Exit Sub
                If Not CheckWigetData(Textk4, "孔隙率", True, False) Then Exit Sub
                If Not CheckWigetData(Textk5, "有效流通面积", True, False) Then Exit Sub
                If Not CheckWigetData(Textk6, "耐压差", True, True) Then Exit Sub
                If Not CheckWigetData(Textk7, "耐温", True, False) Then Exit Sub
                If Not CheckWigetData(Textk8, "抗内压强度", True, False) Then Exit Sub
                If Not CheckWigetData(Textk9, "抗外挤强度", True, False) Then Exit Sub
                If Not CheckWigetData(Textk10, "主材屈服轻度", True, False) Then Exit Sub
                If Not CheckWigetData(Textk11, "拦截的最小颗粒直径", True, False) Then Exit Sub
                If Not CheckWigetData(Textk13, "接头抗拉伸强度", True, False) Then Exit Sub

        End Select
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.Read Then
            msg_prompt = "是否要建立新的管柱数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn_userdb.Close()
                Exit Sub
            End If
            SQL_command = "delete * from 管柱_油管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_普通钻杆 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()

            '2025.6.18 李润洲增加
            SQL_command = "delete * from 管柱_射孔枪弹表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '2025.7.24 李润洲增加
            SQL_command = "delete * from 管柱_筛管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & Text8.Text
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '2025.8.2 李润洲增加
            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            'SQL_command = "delete * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            'EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段爆轰计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段封隔器计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔油管应力计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()


            SQL_command = "insert into 管柱数据表(作业名称,元件序号,元件名称,元件性质,元件长度m,元件外径mm,元件内径mm,井号) values (" _
                & "'" & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Text7.Text & "','" & Text12.Text & "'," & CStr(Text9.Text) & "," & CStr(Text10.Text) & "," & CStr(Text11.Text) & ",'" & well_name & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn_userdb.Close()
            Call save_case()
            Call fill_gzhgrid()
            msg_prompt = "新的管柱元件数据保存完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            msg_prompt = "是否要保存对所选序号的管柱元件数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn_userdb.Close()
                Exit Sub
            End If
            SQL_command = "delete * from 管柱_油管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 管柱_普通钻杆 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()

            '2025.6.18 李润洲增加
            SQL_command = "delete * from 管柱_射孔枪弹表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '2025.7.24 李润洲增加
            SQL_command = "delete * from 管柱_筛管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            '2025.8.2 李润洲增加
            SQL_command = "delete * from 工况_射孔夹层表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            'SQL_command = "delete * from 射孔工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            'EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段爆轰计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔段封隔器计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "delete * from 射孔油管应力计算参数 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()


            SQL_command = "update 管柱数据表 set " & " 元件名称='" & Text7.Text & "'," & " 元件性质='" & Text12.Text & "'," _
                & " 元件长度m=" & CStr(Text9.Text) & "," & " 元件外径mm=" & CStr(Text10.Text) & "," & " 元件内径mm=" & CStr(Text11.Text) _
                & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(Text8.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn_userdb.Close()
            Call save_case()
            Call fill_gzhgrid()
            msg_prompt = "管柱元件数据保存完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Call drawWellStruction(Picture1, 2)
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    Private Sub save_case()
        Dim SQL_command As String
        'Dim jh_xs As String
        Dim bzh As String
        Dim tishi As String
        Dim rec_value As String
        Dim basedb_chg As Boolean

        bzh = (Today) & "输入" & well_name & "井数据时加入。"
        basedb_chg = False
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        Select Case Text12.Text
            Case "油井管"
                SQL_command = "insert into 管柱_油管表(作业名称,元件序号,元件名称,油管壁厚mm,油管钢级," _
                    & "单位长重kg╱m,抗拉强度kN,屈服强度MPa,弹性模量MPa,泊松比," _
                    & "抗内压强度MPa,抗挤强度MPa,扣型,接头抗内压强度MPa,接头抗拉强度kN," _
                    & "备注,热膨胀系数,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Text7.Text & "'," & CStr(Text13.Text) & ",'" & Trim(Text14.Text) & "'," _
                    & CStr(Text15.Text) & "," & CStr(Text16.Text) & "," & CStr(Text17.Text) & "," & CStr(Text18.Text) & "," & CStr(Text19.Text) & "," _
                    & CStr(Text20.Text) & "," & CStr(Text21.Text) & ",'" & Trim(Text110.Text) & "'," & CStr(Text113.Text) & "," & CStr(Text114.Text) & ",'" _
                    & Text109.Text & "'," & CStr(TextBox1.Text) & ",'" & well_name & " ')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 油管 " _
                                    & " where [材料]='" & Trim(Text14.Text) & "' and [油管外径mm]= " & CStr(Text10.Text) _
                                    & " and [油管壁厚mm]=" & CStr(Text13.Text) & " and  [单位长度质量kg/m]=" & CStr(Text15.Text) _
                                    & " and [扣型]='" & Trim(Text110.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此油管加入到油管数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 油管(" & "[油管规格],[材料],[油管外径mm], [油管壁厚mm],[单位长度质量kg/m]," _
                        & "[泊松比],[弹性模量MPa],[屈服应力MPa],[抗内压强度MPa],[抗外挤强度MPa]," _
                        & "[抗拉强度kN],[扣型],[接头抗内压强度MPa],[接头抗拉强度kN],[备注],[热膨胀系数]) values (" _
                        & "'" & Trim(Text7.Text) & "','" & Trim(Text14.Text) & "'," & CStr(Text10.Text) & "," & CStr(Text13.Text) & "," & CStr(Text15.Text) & "," _
                        & CStr(Text19.Text) & "," & CStr(Text18.Text) & "," & CStr(Text17.Text) & "," & CStr(Text20.Text) & "," & CStr(Text21.Text) & "," _
                        & CStr(Text16.Text) & ",'" & Trim(Text110.Text) & "'," & CStr(Text113.Text) & "," & CStr(Text114.Text) & ",'" & Trim(bzh) & "'," & CStr(TextBox1.Text) & ")"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                Else
                    '油管库及界面中的某一数据均不为空或0，但不相等，应提示用户是否用界面数据更新油管库数据。
                    tishi = ""
                    rec_value = RECreader.Item("油管规格").ToString
                    If rec_value <> "" And Trim(Text7.Text) <> "" And rec_value <> Text7.Text Then
                        tishi = "油管规格：数据库中为‘" & rec_value & "’，界面中为‘" & Trim(Text7.Text) & "’"
                    End If
                    rec_value = RECreader.Item("屈服应力MPa").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text17.Text) <> 0.0# And rec_value <> Text17.Text Then
                        If tishi = "" Then
                            tishi = "屈服极限：数据库中=" & rec_value & "，界面中=" & CStr(Text17.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "屈服极限：数据库中=" & rec_value & "，界面中=" & CStr(Text17.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("抗外挤强度MPa").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text21.Text) <> 0.0# And rec_value <> Text21.Text Then
                        If tishi = "" Then
                            tishi = "抗挤强度：数据库中=" & rec_value & "，界面中=" & CStr(Text21.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "抗挤强度：数据库中=" & rec_value & "，界面中=" & CStr(Text21.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("抗内压强度MPa").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text20.Text) <> 0.0# And rec_value <> Text20.Text Then
                        If tishi = "" Then
                            tishi = "管体抗内压强度：数据库中=" & rec_value & "，界面中=" & CStr(Text20.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "管体抗内压强度：数据库中=" & rec_value & "，界面中=" & CStr(Text20.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("抗拉强度kN").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text16.Text) <> 0.0# And rec_value <> Text16.Text Then
                        If tishi = "" Then
                            tishi = "管体抗拉强度：数据库中=" & rec_value & "，界面中=" & CStr(Text16.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "管体抗拉强度：数据库中=" & rec_value & "，界面中=" & CStr(Text16.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("接头抗内压强度MPa").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text113.Text) <> 0.0# And rec_value <> Text113.Text Then
                        If tishi = "" Then
                            tishi = "接头抗内压强度：数据库中=" & rec_value & "，界面中=" & CStr(Text113.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "接头抗内压强度：数据库中=" & RECreader.Item("接头抗内压强度MPa").ToString & "，界面中=" & CStr(Text113.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("接头抗拉强度kN").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text114.Text) <> 0.0# And rec_value <> Text114.Text Then
                        If tishi = "" Then
                            tishi = "接头抗拉强度：数据库中=" & rec_value & "，界面中=" & CStr(Text114.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "接头抗拉强度：数据库中=" & RECreader.Item("接头抗拉强度kN").ToString & "，界面中=" & CStr(Text114.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("弹性模量MPa").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text18.Text) <> 0.0# And rec_value <> Text18.Text Then
                        If tishi = "" Then
                            tishi = "弹性模量：数据库中=" & rec_value & "，界面中=" & CStr(Text18.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "弹性模量：数据库中=" & RECreader.Item("弹性模量MPa").ToString & "，界面中=" & CStr(Text18.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("泊松比").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(Text19.Text) <> 0.0# And rec_value <> Text19.Text Then
                        If tishi = "" Then
                            tishi = "泊松比：数据库中=" & rec_value & "，界面中=" & CStr(Text19.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "泊松比：数据库中=" & RECreader.Item("泊松比").ToString & "，界面中=" & CStr(Text19.Text)
                        End If
                    End If
                    rec_value = RECreader.Item("热膨胀系数").ToString
                    If (rec_value <> "" And Val(rec_value) <> 0.0#) And Val(TextBox1.Text) <> 0.0# And rec_value <> TextBox1.Text Then
                        If tishi = "" Then
                            tishi = "热膨胀系数：数据库中=" & rec_value & "，界面中=" & CStr(TextBox1.Text)
                        Else
                            tishi = tishi & Chr(13) & Chr(10) & "热膨胀系数：数据库中=" & RECreader.Item("热膨胀系数").ToString & "，界面中=" & CStr(TextBox1.Text)
                        End If
                    End If
                    If tishi <> "" Then
                        msg_prompt = "油管数据库中下列数据与界面中的值不同：" & Chr(13) & Chr(10) & tishi & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "是否用界面中的数据更新油管数据库？"
                        msg_buttons = 4 + 32
                        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        If msg_return = 6 Then
                            bzh = (Today) & "输入" & well_name & "井数据时最后更新。"
                            SQL_command = "update 油管 set [备注]='" & bzh & "',[油管规格]='" & Trim(Text7.Text) & "'," & " [屈服应力MPa]=" & CStr(Text17.Text) _
                                & ",[抗外挤强度MPa]=" & CStr(Text21.Text) & "," & " [抗内压强度MPa]=" & CStr(Text20.Text) & ",[抗拉强度kN]=" & CStr(Text16.Text) & "," _
                                & " [接头抗内压强度MPa]=" & CStr(Text113.Text) & ",[接头抗拉强度kN]=" & CStr(Text114.Text) & "," _
                                & " [弹性模量MPa]=" & CStr(Text18.Text) & ",[泊松比]=" & CStr(Text19.Text) & ",[热膨胀系数]=" & CStr(TextBox1.Text) _
                                & " where [材料]='" & Trim(Text14.Text) & "' and [油管外径mm]= " & CStr(Text10.Text) & " and [油管壁厚mm]=" & CStr(Text13.Text) _
                                & " and  [单位长度质量kg/m]=" & CStr(Text15.Text) & " and [扣型]='" & Trim(Text110.Text) & "'"
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                            EXECOleDbCommand.ExecuteNonQuery()
                            basedb_chg = True
                        End If
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    Call fill_yjggrid()
                End If
            Case "节流工具"
                SQL_command = "insert into 管柱_节流元件(作业名称,元件序号,元件名称,重量kg,流孔内径mm,流孔长度mm,抗拉强度kN,抗外挤强度MPa,抗内压强度MPa,节流流向,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Text7.Text & "'," & CStr(Text42.Text) & "," & CStr(Text31.Text) & "," _
                    & CStr(Text32.Text) & "," & CStr(Text44.Text) & "," & CStr(Text41.Text) & "," & CStr(Text43.Text) & ",'" & Text45.Text & "','" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 节流元件 " _
                        & " where [外径(mm)]=" & CStr(Text10.Text) & " and [内径(mm)]= " & CStr(Text11.Text) _
                        & " and [长度(m)]=" & CStr(Text9.Text) & " and  [节流流向]='" & Text45.Text & "' and  [名称]='" & Trim(Text7.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此节流元件加入到节流元件数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 节流元件(" _
                            & "[名称],[外径(mm)],[内径(mm)], [重量(kg)], [长度(m)]," _
                            & "[节流孔内径(mm)],[节流孔长度(mm)],[抗内压强度(MPa)],[抗外挤强度(MPa)],[抗拉强度(kN)]," _
                            & "[节流流向]) values (" _
                            & "'" & Trim(Text7.Text) & "'," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text42.Text) & "," & CStr(Text9.Text) & "," _
                            & CStr(Text31.Text) & "," & CStr(Text32.Text) & "," & CStr(Text43.Text) & "," & CStr(Text41.Text) & "," & CStr(Text44.Text) & ",'" _
                            & Text45.Text & "')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    Call fill_jlgjgrid()
                End If
            Case "封隔器"
                '******************************************************************************************************************************************
                '李润洲2025年7月10日修改，用插管外径存放中心管外径，插管内径存放中心管内径
                '******************************************************************************************************************************************
                SQL_command = "insert into 管柱_封隔定位元件(作业名称,元件序号,元件名称,极限压差MPa,极限载荷kN,坐封压力MPa,坐封力kN,重量kg,坐封方式,井号," _
                        & "温度范围下℃,温度范围上℃,抗内压强度MPa,抗外压强度MPa,型号,生产厂家,上端扣型,下端扣型,主体材料,主材屈服强度MPa," _
                        & "最小坐封压力MPa,最大坐封压力MPa,备注,能否反洗井,插管外径mm,插管内径mm) values ('" _
                        & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Text7.Text & "'," & CStr(Text39.Text) & "," & CStr(Text38.Text) & "," _
                        & CStr(Text37.Text) & "," & CStr(Text36.Text) & "," & CStr(Text40.Text) & ",'" & Combo1.Text & "','" & well_name & "'," _
                        & CStr(Text54.Text) & "," & CStr(Text55.Text) & "," & CStr(Text33.Text) & "," & CStr(Text57.Text) & ",'" & Text22.Text & "','" _
                        & Text35.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Text52.Text & "'," & CStr(Text53.Text) & "," _
                        & CStr(Text34.Text) & "," & CStr(Text56.Text) & ",'" & Text58.Text & "','" & Combo3.Text & "'," & Textf2.Text & "," & Textf1.Text & ")"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 封隔器 " & " where [最大外径(mm)]=" & CStr(Text10.Text) & " and [最小通径(mm)]= " _
                    & CStr(Text11.Text) & " and [长度(m)]=" & CStr(Text9.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Trim(Text22.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此封隔器加入到封隔器数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 封隔器 ([型号],[坐封方式],[最大外径(mm)],[最小通径(mm)],[长度(m)]," _
                            & "[封隔器名称],[重量kg],[生产厂家],[温度范围下℃],[温度范围上℃]," _
                            & "[主体材料],[主材屈服强度MPa],[上端扣型],[极限承压(MPa)],[下端扣型]," _
                            & "[坐封力(kN)],[坐封压差(MPa)],[最小坐封压力MPa],[最大坐封压力MPa],[抗内压强度MPa]," _
                            & "[抗外压强度MPa],[极限载荷(kN)],[备注],[能否反洗井],[中心管外径(mm)],[中心管内径(mm)]) values (" _
                            & "'" & Trim(Text22.Text) & "','" & Combo1.Text & "'," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text9.Text) & "," _
                            & "'" & Text7.Text & "'," & CStr(Text40.Text) & ",'" & Text35.Text & "'," & CStr(Text54.Text) & "," & CStr(Text55.Text) & "," _
                            & "'" & Text52.Text & "'," & CStr(Text53.Text) & ",'" & Text23.Text & "'," & CStr(Text39.Text) & ",'" & Text24.Text & "'," _
                            & CStr(Text36.Text) & "," & CStr(Text37.Text) & "," & CStr(Text34.Text) & "," & CStr(Text56.Text) & "," & CStr(Text33.Text) & "," _
                            & CStr(Text57.Text) & "," & CStr(Text38.Text) & ",'" & bzh & "','" & Combo3.Text & "'," & Textf2.Text & "," & Textf1.Text & ")"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    Call fill_fgqgrid()
                End If
            Case "开关工具"
                SQL_command = "insert into 管柱_开关元件(作业名称,元件序号,元件名称,重量kg,抗拉强度kN," _
                        & "极限压差MPa,抗外挤强度MPa,抗内压强度MPa,温度范围下℃,温度范围上℃," _
                        & "压力等级MPa,型号,开关类型,上端扣型,下端扣型," _
                        & "生产厂家,备注,井号) values (" _
                        & "'" & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Text7.Text & "'," & CStr(Text25.Text) & "," & CStr(Text26.Text) & "," _
                        & CStr(Text29.Text) & "," & CStr(Text28.Text) & "," & CStr(Text27.Text) & "," & CStr(Text76.Text) & "," & CStr(Text88.Text) & "," _
                        & CStr(Text92.Text) & ",'" & Text75.Text & "','" & Text30.Text & "','" & Text89.Text & "','" & Text90.Text & "','" _
                        & Text93.Text & "','" & Text91.Text & "','" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 开关元件 " & " where [外径(mm)]=" & CStr(Text10.Text) & " and [内径(mm)]= " & CStr(Text11.Text) _
                    & " and [长度(m)]=" & CStr(Text9.Text) & " and  [开关类型]='" & Trim(Text30.Text) & "'" & " and [型号]='" & Trim(Text75.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此开关元件加入到开关元件数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 开关元件(" _
                            & "[名称],[外径(mm)],[内径(mm)], [重量(kg)], [长度(m)]," _
                            & "[抗拉强度(kN)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[极限压差(MPa)],[开关类型]," _
                            & "[温度范围下℃],[温度范围上℃],[压力等级MPa],[型号],[上端扣型]," _
                            & "[下端扣型],[生产厂家],[备注]) values (" & "'" _
                            & Trim(Text7.Text) & "'," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text25.Text) & "," & CStr(Text9.Text) & "," _
                            & CStr(Text26.Text) & "," & CStr(Text28.Text) & "," & CStr(Text27.Text) & "," & CStr(Text29.Text) & ",'" & CStr(Text30.Text) & "'," _
                            & CStr(Text76.Text) & "," & CStr(Text88.Text) & "," & CStr(Text92.Text) & ",'" & Trim(Text75.Text) & "','" & Trim(Text89.Text) & "','" _
                            & Trim(Text90.Text) & "','" & Trim(Text93.Text) & "','" & Trim(bzh) & "')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    fill_kggjgrid(Text30.Text)
                End If
            Case "伸缩管"
                SQL_command = "insert into 管柱_伸缩元件(作业名称,元件序号,元件名称,重量kg,伸缩行程m," _
                        & "抗外挤强度MPa,抗内压强度MPa,井号,温度范围下℃,温度范围上℃," _
                        & "压力等级MPa,抗拉强度kN,上端扣型,下端扣型,主体材料," _
                        & "主材屈服强度MPa,零件号,生产厂家,备注,全缩短长度m," _
                        & "伸缩动作载荷kN,伸缩动作压力MPa) values (" & "'" _
                        & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Trim(Text7.Text) & "'," & CStr(Text46.Text) & "," & CStr(Text48.Text) & "," _
                        & CStr(Text51.Text) & "," & CStr(Text50.Text) & ",'" & well_name & "'," & CStr(Text60.Text) & "," & CStr(Text61.Text) & "," _
                        & CStr(Text64.Text) & "," & CStr(Text65.Text) & ",'" & Trim(Text67.Text) & "','" & Trim(Text68.Text) & "','" & Trim(Text69.Text) & "'," _
                        & CStr(Text62.Text) & ",'" & Trim(Text49.Text) & "','" & Trim(Text47.Text) & "','" & Trim(Text59.Text) & "'," & CStr(Text66.Text) & "," _
                        & CStr(Text63.Text) & "," & CStr(TextBox3.Text) & ")"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 伸缩元件 " & " where [外径(mm)]=" & CStr(Text10.Text) & " and [内径(mm)]= " & CStr(Text11.Text) _
                        & " and [全缩短长度m]=" & CStr(Text66.Text) & "and [伸缩行程(m)]=" & CStr(Text48.Text) & " and [零件号]='" & Trim(Text49.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此伸缩元件加入到伸缩元件数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 伸缩元件([元件名称],[外径(mm)],[内径(mm)], [重量(kg)], [全缩短长度m]," _
                                & "[伸缩行程(m)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[压力等级MPa],[抗拉强度kN]," _
                                & "[上端扣型],[下端扣型],[主体材料],[主材屈服强度MPa],[温度范围下℃]," _
                                & "[温度范围上℃],[零件号],[生产厂家],[备注]) values (" & "'" _
                                & Trim(Text7.Text) & "'," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text46.Text) & "," & Text66.Text & "," _
                                & CStr(Text48.Text) & "," & CStr(Text51.Text) & "," & CStr(Text50.Text) & "," & CStr(Text64.Text) & "," & CStr(Text65.Text) & "," & "'" _
                                & Trim(Text67.Text) & "','" & Trim(Text68.Text) & "','" & Trim(Text69.Text) & "'," & CStr(Text62.Text) & "," & CStr(Text60.Text) & "," _
                                & CStr(Text61.Text) & ",'" & Trim(Text49.Text) & "','" & Trim(Text47.Text) & "','" & Trim(bzh) & "')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    fill_shsggrid()
                End If
            Case "锚定工具"
                SQL_command = "insert into 管柱_锚定元件(作业名称,元件序号,元件名称,型号,重量kg," _
                        & "工作压力MPa,温度范围下℃,温度范围上℃,坐卡方式,小坐卡压力MPa," _
                        & "大坐卡压力MPa,上端扣型,下端扣型,生产厂家,最大锚定力kN," _
                        & "最大解锚力kN,抗内压强度MPa,抗外压强度MPa,抗拉强度kN,备注,井号) values (" & "'" _
                        & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Trim(Text7.Text) & "','" & Trim(Text79.Text) & "'," & CStr(Text80.Text) & "," _
                        & CStr(Text81.Text) & "," & CStr(Text74.Text) & "," & CStr(Text73.Text) & ",'" & Combo2.Text & "'," & CStr(Text86.Text) & "," _
                        & CStr(Text72.Text) & ",'" & Trim(Text78.Text) & "','" & Trim(Text77.Text) & "','" & Trim(Text85.Text) & "'," & CStr(Text84.Text) & "," _
                        & CStr(Text83.Text) & "," & CStr(Text87.Text) & "," & CStr(Text71.Text) & "," & CStr(Text82.Text) & ",'" & Trim(Text70.Text) & "','" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 锚定工具表 " & " where 最大外径mm=" & CStr(Text10.Text) & " and 最小通径mm= " & CStr(Text11.Text) _
                        & " and 总体长度m=" & CStr(Text9.Text) & " and  坐卡方式='" & Combo2.Text & "' " & " and 型号='" & Trim(Text79.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.Read Then
                    msg_prompt = "是否将此锚定工具加入到锚定工具数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 锚定工具表 (型号,坐卡方式,最大外径mm,最小通径mm,总体长度m," _
                        & "工具名称,重量kg,生产厂家,温度范围下℃,温度范围上℃," _
                        & "上端扣型,小坐卡压力MPa,下端扣型,大坐卡压力MPa,最大锚定力kN," _
                        & "最大解锚力kN,工作压力MPa,抗内压强度MPa,抗外压强度MPa,抗拉强度kN," _
                        & "备注) values (" & "'" _
                        & Trim(Text79.Text) & "','" & Combo2.Text & "'," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text9.Text) & "," & "'" _
                        & Text7.Text & "'," & CStr(Text80.Text) & ",'" & Text85.Text & "'," & CStr(Text74.Text) & "," & CStr(Text73.Text) & "," & "'" _
                        & Text78.Text & "'," & CStr(Text86.Text) & ",'" & Text77.Text & "'," & CStr(Text72.Text) & "," & CStr(Text84.Text) & "," _
                        & CStr(Text83.Text) & "," & CStr(Text81.Text) & "," & CStr(Text87.Text) & "," & CStr(Text71.Text) & "," & CStr(Text82.Text) & "," & "'" _
                        & bzh & " ')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    fill_mdgjgrid()
                End If
            Case "普通钻杆"
                SQL_command = "insert into 管柱_普通钻杆(作业名称,元件序号,规格名称,钻杆外径mm,钻杆壁厚mm," _
                        & "钢级,管体抗拉强度kN,屈服强度MPa,弹性模量MPa,泊松比," _
                        & "加厚型式,单根长度m,单位长度质量kgpm,扣型,管体抗扭强度Nm," _
                        & "接头抗扭强度Nm,接头抗拉强度kN,抗内压强度MPa,抗挤强度MPa,接头外径mm," _
                        & "备注,热膨胀系数,井号) values (" & "'" _
                        & zuoye_name & "'," & CStr(Text8.Text) & ",'" & Trim(Text7.Text) & "'," & CStr(Text10.Text) & "," & CStr(Text108.Text) & ",'" _
                        & Trim(Text101.Text) & "'," & CStr(Text105.Text) & "," & CStr(Text103.Text) & "," & CStr(Text96.Text) & "," & CStr(Text95.Text) & ",'" _
                        & Trim(Text112.Text) & "'," & CStr(Text107.Text) & "," & CStr(Text111.Text) & ",'" & Trim(Text102.Text) & "'," & CStr(Text100.Text) & "," _
                        & CStr(Text99.Text) & "," & CStr(Text94.Text) & "," & CStr(Text98.Text) & "," & CStr(Text97.Text) & "," & CStr(Text104.Text) & ",'" _
                        & Trim(Text106.Text) & "'," & CStr(TextBox2.Text) & ",'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 钻杆 " & " where [钻杆规格]='" & Trim(Text7.Text) & "' and [钻杆外径(mm)]= " & CStr(Text10.Text) _
                        & " and [钻杆壁厚(mm)]=" & CStr(Text108.Text) & " and  [单根长度(m)]=" & CStr(Text107.Text) & " and  [单位长度质量(Kg/m)]=" & CStr(Text111.Text)
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.HasRows Then
                    msg_prompt = "是否将此钻杆加入到普通钻杆数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into  钻杆 ([钻杆规格],[钻杆外径(mm)],[钻杆壁厚(mm)],[管体抗拉强度kN],[单根长度(m)]," _
                                & "[单位长度质量(Kg/m)],[加厚型式],[备注],[钢级],[扣型]," _
                                & "[屈服强度MPa],[弹性模量MPa],[泊松比],[管体抗扭强度Nm],[接头抗扭强度Nm]," _
                                & "[接头抗拉强度kN],[抗内压强度MPa],[抗挤强度MPa],[接头外径mm],[热膨胀系数]) values (" & "'" _
                                & Trim(Text7.Text) & "'," & Str(Val(Text10.Text)) & "," & Str(Val(Text108.Text)) & "," & Str(Val(Text105.Text)) & "," & Str(Val(Text107.Text)) & "," _
                                & Str(Val(Text111.Text)) & ",'" & Trim(Text112.Text) & "','" & IIf(Trim(Text101.Text) = "", bzh, Trim(Text101.Text)) & "','" & Trim(Text101.Text) & " ','" & Trim(Text102.Text) & "'," _
                                & Str(Val(Text103.Text)) & "," & Str(Val(Text96.Text)) & "," & Str(Val(Text95.Text)) & "," & Str(Val(Text100.Text)) & "," & Str(Val(Text99.Text)) & "," _
                                & Str(Val(Text94.Text)) & "," & Str(Val(Text98.Text)) & "," & Str(Val(Text97.Text)) & "," & Str(Val(Text104.Text)) & Str(Val(TextBox2.Text)) & ")"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    fill_ptzggrid()
                End If
                '2025.6.18 李润洲增加
            Case "射孔枪"
                SQL_command = "insert into 管柱_射孔枪弹表(作业名称,元件序号,射孔器分类,元件名称,壁厚mm," _
                        & "耐压MPa,[射孔密度孔╱m],射孔相位°,射孔弹名称,炸药名称," _
                        & "装药量g,[装药密度g╱cm3] ,[耐温°C],[平均孔径mm]," _
                        & "[平均穿深mm],备注,井号) values (" & "'" _
                        & zuoye_name & "'," & Text8.Text & ",'" & Trim(Texts1.Text) & "','" & Trim(Text7.Text) & "'," & Texts2.Text & "," & Texts5.Text & "," _
                        & Texts6.Text & "," & Texts7.Text & ",'" & Texts9.Text & "','" & Texts20.Text & "'," & Texts11.Text & "," _
                        & Texts12.Text & "," & Texts13.Text & "," & Texts14.Text & "," & Texts15.Text & ",'" _
                        & Texts19.Text & "','" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 枪弹数据表 " & " where [射孔器名称]='" & Trim(Text7.Text) & "' and [射孔弹名称]='" & CStr(Texts9.Text) _
                        & "' and [炸药名称]='" & CStr(Texts20.Text) & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.HasRows Then
                    msg_prompt = "是否将此射孔枪加入到射孔枪弹数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 枪弹数据表(射孔器分类,射孔器名称,外径mm,壁厚mm," _
                        & "耐压MPa,[射孔密度孔╱m],射孔相位°,射孔弹名称,炸药名称," _
                        & "装药量g,[装药密度g╱cm3] ,[耐温°C],[套管外径mm],[平均孔径mm]," _
                        & "[平均穿深mm],流动效率,钢靶孔径mm,钢靶穿深mm," _
                        & "备注) values (" & "'" _
                        & Texts1.Text & "', '" & Text7.Text & "'," & Text10.Text & "," & Texts2.Text & "," & Texts5.Text & "," _
                        & Texts6.Text & "," & Texts7.Text & ",'" & Texts9.Text & "','" & Texts20.Text & "'," & Texts11.Text & "," _
                        & Texts12.Text & "," & Texts13.Text & ",0.0" & "," & Texts14.Text & "," & Texts15.Text & ",0,0,0,'" _
                        & Texts19.Text & "')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg Then
                    Call fill_skqdgrid()
                End If
                '2025.7.25 李润洲增加
            Case "筛管"
                SQL_command = "insert into 管柱_筛管表(作业名称,元件序号,元件名称,壁厚mm,单根长度m,单根质量kg," _
                        & "主体材料,接头抗拉伸强度kN,主材屈服强度MPa,抗内压强度MPa,抗外挤强度MPa," _
                        & "耐压差MPa,耐温℃,[孔隙率%],有效流通面积cm2╱m,拦截的最小颗粒直径mm," _
                        & "生产厂家,备注,井号) values (" & "'" _
                        & zuoye_name & "'," & Text8.Text & ",'" & Trim(Text7.Text) & "'," & Textk15.Text & "," & Textk2.Text & "," & Textk3.Text & ",'" _
                        & Textk1.Text & "'," & Textk13.Text & "," & Textk10.Text & "," & Textk8.Text & "," & Textk9.Text & "," _
                        & Textk6.Text & "," & Textk7.Text & "," & Textk4.Text & "," & Textk5.Text & "," & Textk11.Text & ",'" _
                        & Textk12.Text & "','" & Textk14.Text & "','" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "select * from 筛管 " & " where 外径mm=" & Trim(Text10.Text) & " and 壁厚mm=" & Trim(Textk15.Text) _
                        & " and 单根长度m=" & Trim(Textk2.Text)
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.HasRows Then
                    msg_prompt = "是否将此筛管加入到筛管数据库中？"
                    msg_buttons = 4 + 32
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    If msg_return = 6 Then
                        SQL_command = "insert into 筛管(规格名称,外径mm,壁厚mm,单根长度m,单根质量kg," _
                        & "主体材料,接头抗拉伸强度kN,主材屈服强度MPa,抗内压强度MPa,抗外挤强度MPa," _
                        & "耐压差MPa,耐温℃,[孔隙率%],有效流通面积cm2╱m,拦截的最小颗粒直径mm," _
                        & "生产厂家,备注) values (" & "'" _
                        & Trim(Text7.Text) & "'," & Text10.Text & "," & Textk15.Text & "," & Textk2.Text & "," & Textk3.Text & ",'" _
                        & Textk1.Text & "'," & Textk13.Text & "," & Textk10.Text & "," & Textk8.Text & "," & Textk9.Text & "," _
                        & Textk6.Text & "," & Textk7.Text & "," & Textk4.Text & "," & Textk5.Text & "," & Textk11.Text & ",'" _
                        & Textk12.Text & "','" & Textk14.Text & "')"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        basedb_chg = True
                    End If
                End If
                RECreader.Close()
                cn_basedb.Close()
                If basedb_chg = True Then
                    fill_sgGrid()
                End If
        End Select
        cn_userdb.Close()
    End Sub
    '*********************************************************************************************************************************************
    '节流工具节流流向List选择赋值
    '   UPGRADE_WARNING: 初始化窗体时可能激发事件 List2.SelectedIndexChanged。 
    '   单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"”
    '*********************************************************************************************************************************************
    Private Sub List2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List2.SelectedIndexChanged
        Text45.Text = List2.Text
    End Sub
    '*********************************************************************************************************************************************
    '炸药名设置listbox3，李润洲2025年8月1日
    '*********************************************************************************************************************************************
    Private Function fill_zym() As Boolean
        fill_zym = True
        ListBox3.Items.Clear()
        Try
            Using cn_basedb As New System.Data.OleDb.OleDbConnection(AdoConString)
                cn_basedb.Open()
                SQL_command = "select 炸药名 from 炸药表 order by 炸药名"
                Using oleDbCmd As New OleDbCommand(SQL_command, cn_basedb)
                    Using reader As OleDbDataReader = oleDbCmd.ExecuteReader()
                        While reader.Read
                            ListBox3.Items.Add(reader.Item(0).ToString)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As OleDbException
            fill_zym = False
            msg_prompt = "基础数据库枪弹数据表访问异常。" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
    End Function
    '单击选择射孔枪按钮  李润洲2025年6月
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        refresh_li()
    End Sub
    '双击射孔枪grid  李润洲2025年6月
    Private Sub DataGridView11_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView11.DoubleClick
        refresh_li()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的射孔段列表中的数据填写界面中各文本框，李润洲2025年6月
    '*********************************************************************************************************************************************
    Private Sub refresh_li()
        If (IsNothing(Me.BindingSource11.Current)) Then Exit Sub
        If (Not IsDBNull(Me.BindingSource11.Current("射孔器名称"))) Then
            Texts1.Text = If(BindingSource11.Current("射孔器分类") Is DBNull.Value, "", BindingSource11.Current("射孔器分类").ToString)
            Text7.Text = If(BindingSource11.Current("射孔器名称") Is DBNull.Value, "", BindingSource11.Current("射孔器名称").ToString)
            Text10.Text = If(BindingSource11.Current("外径mm") Is DBNull.Value, "0.0", BindingSource11.Current("外径mm").ToString)
            Texts5.Text = If(BindingSource11.Current("耐压MPa") Is DBNull.Value, "0.0", BindingSource11.Current("耐压MPa").ToString)
            Texts6.Text = If(BindingSource11.Current("射孔密度孔╱m") Is DBNull.Value, "0.0", BindingSource11.Current("射孔密度孔╱m").ToString)
            Texts7.Text = If(BindingSource11.Current("射孔相位°") Is DBNull.Value, "0.0", BindingSource11.Current("射孔相位°").ToString)
            Texts9.Text = If(BindingSource11.Current("射孔弹名称") Is DBNull.Value, "", BindingSource11.Current("射孔弹名称").ToString)
            Texts20.Text = If(BindingSource11.Current("炸药名称") Is DBNull.Value, "", BindingSource11.Current("炸药名称").ToString)
            Texts11.Text = If(BindingSource11.Current("装药量g") Is DBNull.Value, "0.0", BindingSource11.Current("装药量g").ToString)
            Texts12.Text = If(BindingSource11.Current("装药密度g╱cm3") Is DBNull.Value, "0.0", BindingSource11.Current("装药密度g╱cm3").ToString)
            Texts13.Text = If(BindingSource11.Current("耐温°C") Is DBNull.Value, "0.0", BindingSource11.Current("耐温°C").ToString)
            Texts14.Text = If(BindingSource11.Current("平均孔径mm") Is DBNull.Value, "0.0", BindingSource11.Current("平均孔径mm").ToString)
            Texts15.Text = If(BindingSource11.Current("平均穿深mm") Is DBNull.Value, "", BindingSource11.Current("平均穿深mm").ToString)
            Texts19.Text = If(BindingSource11.Current("备注") Is DBNull.Value, "", BindingSource11.Current("备注").ToString)
            Texts2.Text = If(BindingSource11.Current("壁厚mm") Is DBNull.Value, "0", BindingSource11.Current("壁厚mm").ToString)
            If (BindingSource11.Current("外径mm") Is DBNull.Value Or BindingSource11.Current("壁厚mm") Is DBNull.Value) Then
                Text11.Text = "0.0"
            Else
                Text11.Text = Convert.ToSingle(BindingSource11.Current("外径mm")) - 2 * Convert.ToSingle(BindingSource11.Current("壁厚mm"))
            End If
            '用套管外径匹配一下井筒套管外径，排除不合适的射孔枪，目前还没有加上
            'If Val(Text5.Text) < Val(TextBox5.Text) Then
            '    msg_prompt = "射孔底界比完钻井深深，数据不合理！"
            '    msg_buttons = 0 + 48
            '    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            'End If
        End If
    End Sub
    '单击选择筛管按钮  李润洲2025年7月
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        refresh_li_sg()
    End Sub

    '双击筛管grid  李润洲2025年7月
    Private Sub DataGridView12_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView12.DoubleClick
        refresh_li_sg()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的筛管列表中的数据填写界面中各文本框，李润洲2025年6月
    '*********************************************************************************************************************************************
    Private Sub refresh_li_sg()
        With BindingSource12
            If (IsNothing(.Current)) Then Exit Sub
            If (Not IsDBNull(.Current("规格名称"))) Then
                Text7.Text = If(.Current("规格名称") Is DBNull.Value, "", .Current("规格名称").ToString)
                Text10.Text = If(.Current("外径mm") Is DBNull.Value, "0.0", .Current("外径mm").ToString)
                Textk1.Text = If(.Current("主体材料") Is DBNull.Value, "", .Current("主体材料").ToString)
                Textk2.Text = If(.Current("单根长度m") Is DBNull.Value, "0.0", .Current("单根长度m").ToString)
                Textk3.Text = If(.Current("单根质量kg") Is DBNull.Value, "0.0", .Current("单根质量kg").ToString)
                Textk4.Text = If(.Current("孔隙率%") Is DBNull.Value, "0.0", .Current("孔隙率%").ToString)
                Textk5.Text = If(.Current("有效流通面积cm2╱m") Is DBNull.Value, "0.0", .Current("有效流通面积cm2╱m").ToString)
                Textk6.Text = If(.Current("耐压差MPa") Is DBNull.Value, "0.0", .Current("耐压差MPa").ToString)
                Textk7.Text = If(.Current("耐温℃") Is DBNull.Value, "0.0", .Current("耐温℃").ToString)
                Textk8.Text = If(.Current("抗内压强度MPa") Is DBNull.Value, "0.0", .Current("抗内压强度MPa").ToString)
                Textk9.Text = If(.Current("抗外挤强度MPa") Is DBNull.Value, "0.0", .Current("抗外挤强度MPa").ToString)
                Textk10.Text = If(.Current("主材屈服强度MPa") Is DBNull.Value, "0.0", .Current("主材屈服强度MPa").ToString)
                Textk11.Text = If(.Current("拦截的最小颗粒直径mm") Is DBNull.Value, "0.0", .Current("拦截的最小颗粒直径mm").ToString)
                Textk12.Text = If(.Current("生产厂家") Is DBNull.Value, "", .Current("生产厂家").ToString)
                Textk13.Text = If(.Current("接头抗拉伸强度kN") Is DBNull.Value, "0.0", .Current("接头抗拉伸强度kN").ToString)
                Textk14.Text = If(.Current("备注") Is DBNull.Value, "", .Current("备注").ToString)
                Textk15.Text = If(.Current("壁厚mm") Is DBNull.Value, "0.0", .Current("壁厚mm").ToString)

                If (.Current("外径mm") Is DBNull.Value Or .Current("壁厚mm") Is DBNull.Value) Then
                    Text11.Text = "0.0"
                Else
                    Text11.Text = Convert.ToSingle(.Current("外径mm")) - 2 * Convert.ToSingle(.Current("壁厚mm"))
                End If
            End If
        End With
    End Sub
    '*********************************************************************************************************************************************
    '用选中的炸药名填写texts20，李润洲2025年8月1日
    '*********************************************************************************************************************************************
    Private Sub ListBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox3.Click
        If ListBox3.SelectedItem IsNot Nothing Then
            ' 将选中项的内容填充到TextBox中
            Texts20.Text = ListBox3.SelectedItem.ToString()
        End If
    End Sub
End Class