Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmipt_wellstruct
    '*********************************************************************************************************************************************
    '                                                   关于井身结构数据输入及管理窗体的说明
    '说明：
    '   "不能在 Load 事件处理程序中调用此(DrawLine)方法。如果已调整该窗体的大小或者其他窗体遮蔽了该窗体，则不会重绘所绘制的内容。若要自动重绘内容，
    '应该重写OnPaint 方法。"
    ' -----------------------摘自VS2008帮助:ms-help://MS.VSCC.v90/MS.MSDNQTR.v90.chs/dv_fxmclignrl/html/55c1dbeb-75d0-430c-9814-a24b8971ad8c.htm
    '
    ' 程序升级记事：
    '                                                                                           秦彦斌 2019年10月7日最后整理
    '  （1）因为不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并
    '关闭计时器。经试验，利用窗体的Paint事件画井身结构图会出现界面已显示好但图没画的情况。
    '  （2）摸索掌握了VS中TreeView控件，Graphics类绘图方法法。
    '
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private dst As New DataSet("well_dst")
    Private wst_Table As DataTable = dst.Tables.Add("wst_Table")
    Private caslib_Table As New DataTable
    Private SQL_command As String
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmipt_wellstruct_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面load
    '*********************************************************************************************************************************************
    Private Sub frmipt_wellstruct_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Label_0.Text = well_name & "井油气井参数"
        Label_24.Text = well_name & "井套管参数"
        Me.Text = well_name & "井井身结构数据输入与管理"

        '用ADO.NET给井基本数据赋值
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位 from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
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
        cn.Close()

        Text7.Text = CStr(0)
        Combo1.Text = ""
        Combo2.Text = ""
        Text9.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text12.Text = ""
        Text13.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text15.Text = CStr(0.0#)
        Text16.Text = CStr(0.0#)
        Text17.Text = ""
        Text18.Text = CStr(0.3)
        Text8.Text = CStr(206842.0#)
        Text19.Text = CStr(0.0#)
        Text20.Text = CStr(0.0#)
        Text21.Text = CStr(0.0#)
        Text22.Text = CStr(0.0#)
        Text23.Text = ""
        Text24.Text = ""
        Text25.Text = CStr(0.0#)
        Text28.Text = ""
        Text29.Text = CStr(0.0#)
        Text26.Text = CStr(0.0#)
        Text27.Text = CStr(0.0#)
        TextBox1.Text = CStr(1.3#)
        TextBox2.Text = CStr(1.8#)

        '设置并填充DataGridView1
        DataGridView1.MultiSelect = False
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Call my_gridview1_flash()
        Call tree_change()
        'MsgBox("调试暂停！", , "系统提示")
        '不能在 Load 事件处理程序中调用画图(如DrawLine)方法，故设计时器Time1，200毫秒触发，在触发事件处理程序中调用画图函数画井身结构图并关闭计时器
        Timer1.Interval = 200
        Timer1.Start()
    End Sub
    '*********************************************************************************************************************************************
    'Timer1的Tick事件处理：（1）画井身结构图；（2）关闭计时器
    '*********************************************************************************************************************************************
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '绘制井身结构图,不含管柱
        Call drawWellStruction(Picture1, 1)
        Timer1.Stop()
    End Sub
    '*********************************************************************************************************************************************
    '设置并填充DataGridView1
    '*********************************************************************************************************************************************
    Private Sub my_gridview1_flash()
        cn = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn.Open()
        SQL_command = "select [套管规格],[钢级],[套管外径(mm)],[壁厚(mm)],[扣型],[单位长质量(kg/m)],[屈服极限(MPa)],[抗挤强度(MPa)],[抗内压强度(MPa)],[抗拉强度(kN)]," _
                    & "[接头抗内压强度(MPa)],[接头抗拉强度(kN)],[备注] from 套管  order by [套管外径(mm)] desc,[壁厚(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn)
        caslib_Table.Clear()
        ad.Fill(caslib_Table)
        ad.Dispose()
        cn.Close()
        cn.Dispose()

        BindingSource1.DataSource = caslib_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.Columns(0).Width = 250
        DataGridView1.Columns(1).Width = 90
        DataGridView1.Columns(2).Width = 90
        DataGridView1.Columns(3).Width = 90
        DataGridView1.Columns(4).Width = 90
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).Width = 90
        DataGridView1.Columns(8).Width = 90
        DataGridView1.Columns(9).Width = 90
        DataGridView1.Columns(10).Width = 90
        DataGridView1.Columns(11).Width = 90
        DataGridView1.Columns(12).Width = 375
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '给treeview添加节点
    '*********************************************************************************************************************************************
    Private Sub tree_change()
        Dim Key As String
        Dim text_Renamed As String
        Dim i As Short
        Dim csh As Short
        Dim dsh As Short
        Dim node1 As New System.Windows.Forms.TreeNode
        Dim node2 As New System.Windows.Forms.TreeNode
        Dim ls_Table_1 As New DataTable
        Dim ls_Table_2 As New DataTable
        Dim RECreader_1 As OleDbDataReader
        Dim RECreader_2 As OleDbDataReader

        TreeView1.Nodes.Clear()
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        i = 0
        SQL_command = "select distinct 套管层数 from 套管数据表 where 井号='" & well_name & "' order by 套管层数"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader_1 = EXECOleDbCommand.ExecuteReader()
        While RECreader_1.Read()
            csh = RECreader_1.GetInt32(0)
            Key = "第" & Trim(CStr(csh)) & "层套管"
            text_Renamed = "第" & Trim(CStr(csh)) & "层套管"
            node1 = TreeView1.Nodes.Add(Key, text_Renamed)
            SQL_command = "select distinct 套管段数 from 套管数据表 where 井号='" & well_name & "' and 套管层数= " & CStr(csh) & " order by 套管段数"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            RECreader_2 = EXECOleDbCommand.ExecuteReader()
            While RECreader_2.Read()
                dsh = RECreader_2.GetInt32(0)
                Key = "第" & Trim(CStr(csh)) & "-" & Trim(CStr(dsh)) & "段套管"
                text_Renamed = "第" & Trim(CStr(csh)) & "-" & Trim(CStr(dsh)) & "段套管"
                'UPGRADE_WARNING: Add 方法行为已更改
                'node2 = TreeView1.Nodes.Find(node1.Index, True)(0).Nodes.Add(Key, text_Renamed)
                node2 = TreeView1.Nodes(i).Nodes.Add(Key, text_Renamed)
            End While
            i = i + 1
        End While
        cn.Close()
        cn.Dispose()
        For i = 0 To TreeView1.Nodes.Count - 1
            TreeView1.Nodes.Item(i).Expand()
        Next i
    End Sub
    '*********************************************************************************************************************************************
    '点击“退出”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击“删除”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        On Error GoTo errhandler
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select * from 套管数据表 where 井号='" & well_name & "'and 套管层数= " & CStr(Text7.Text) & " and 套管段数 = " & CStr(Text14.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            msg_prompt = "是否确定要删除所选的套管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn.Close()
                Exit Sub
            End If
            '*********************同时删除“套管磨损分析参数表”中相应的套管段数据******************************************
            SQL_command = "DELETE * from 套管磨损分析参数表 where 井号='" & well_name & "'and 层= " & CStr(Text7.Text) & " and 段 = " & CStr(Text14.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()

            SQL_command = "DELETE * from 套管数据表 where 井号='" & well_name & "'and 套管层数= " & CStr(Text7.Text) & " and 套管段数 = " & CStr(Text14.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn.Close()
            '绘制井身结构图,不含管柱
            Call drawWellStruction(Picture1, 1)
            Call tree_change()
            Call write_chanshu("套管磨损计算-是否重算磨损量", "是")
            msg_prompt = "套管数据删除完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            RECreader.Close()
            cn.Close()
            msg_prompt = "未找到层数、段数对应的数据，请选好要删除的数据！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "层数或段数数据错误，请选好要删除的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击“帮助”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '点击“保存”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        On Error GoTo errhandler
        If Combo1.Text = "" Then
            msg_prompt = "请选择或输入套管类型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo2.Text = "" Then
            msg_prompt = "请选择或输入是否回接。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text7.Text) Then
            msg_prompt = "套管层数应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text7.Text) = 0 Then
            msg_prompt = "套管层数不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text9.Text)) Then
            msg_prompt = "套管外径应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Val(Text9.Text) = 0 Then
            msg_prompt = "套管外径不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text16.Text)) Then
            msg_prompt = "套管壁厚应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text16.Text) = 0 Then
            msg_prompt = "套管壁厚不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text11.Text)) Then
            msg_prompt = "套管下入深度应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text11.Text) = 0 Then
            msg_prompt = "套管下入深度不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text14.Text) = 0 Or (Not IsNumeric(Text14.Text)) Then
            msg_prompt = "请输入套管段数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text15.Text)) Then
            msg_prompt = "下套管时井眼的完钻深度应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text15.Text) = 0 Then
            msg_prompt = "下套管时井眼的完钻深度不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not (Val(Text10.Text) <= Val(Text11.Text) And Val(Text11.Text) <= Val(Text15.Text)) Then
            msg_prompt = "所输入的几个深度数据不符合现实！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text15.Text) > Val(Text5.Text) Then
            msg_prompt = "完钻深度与油气井基本数据中的完钻深度矛盾，数据不合理！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text23.Text = "" Then
            msg_prompt = "请输入套管名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text24.Text = "" Then
            msg_prompt = "请输入或选择套管扣型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text25.Text)) Then
            msg_prompt = "钻头尺寸应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text25.Text) = 0 Then
            msg_prompt = "钻头尺寸不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text26.Text)) Then
            msg_prompt = "套管接头抗内压强度(MPa)值应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text26.Text) = 0 Then
            msg_prompt = "套管接头抗内压强度(MPa)值不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text27.Text)) Then
            msg_prompt = "套管接头抗拉(kN)强度值应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text27.Text) = 0 Then
            msg_prompt = "套管接头抗拉(kN)强度值不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(Text29.Text)) Then
            msg_prompt = "套管单位长度质量[线重(kg/m)]值应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text29.Text) = 0 Then
            msg_prompt = "套管单位长度质量[线重(kg/m)]值不应为零。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(TextBox1.Text)) Then
            msg_prompt = "固井时井液密度(g/cm^3)值应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (Not IsNumeric(TextBox2.Text)) Then
            msg_prompt = "固井水泥浆密度(g/cm^3)值应为数字。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(TextBox1.Text) <= 0 Or Val(TextBox1.Text) > 5 Then
            msg_prompt = "请输入合理的固井时井液密度(g/cm^3)值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(TextBox2.Text) <= 0 Or Val(TextBox2.Text) > 5 Then
            msg_prompt = "请输入合理的固井水泥浆密度(g/cm^3)值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select * from 套管数据表 where 井号='" & well_name & "'and 套管层数= " & CStr(Text7.Text) & " and 套管段数 = " & CStr(Text14.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            msg_prompt = "是否要保存对所选套管数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn.Close()
                Exit Sub
            End If
            SQL_command = "update 套管数据表 set " _
                & " 套管规格='" & Trim(Text12.Text) & "',套管类型='" & Combo1.Text & "'," & " 是否回接='" & Combo2.Text & "',弹性模量MPa=" & CStr(Text8.Text) & "," _
                & " 套管外径mm=" & Str(Val(Text9.Text)) & ",悬挂深度m=" & CStr(Text10.Text) & "," & " 套管下深m=" & CStr(Text11.Text) & ",水泥返深m=" & CStr(Text13.Text) & "," _
                & " 套管段数=" & CStr(Text14.Text) & ",完钻深度m=" & CStr(Text15.Text) & "," & " 套管壁厚mm=" & CStr(Text16.Text) & ",套管钢级=" & "'" & Trim(Text17.Text) & "'," _
                & " 抗内压强度MPa=" & Str(Val(Text19.Text)) & ",抗拉强度kN=" & CStr(Text20.Text) & "," & " 抗挤强度MPa=" & CStr(Text21.Text) & ",泊松比=" & Str(Val(Text18.Text)) & "," _
                & " 套管名称='" & Trim(Text23.Text) & "',扣型='" & Trim(Text24.Text) & "'," & " 单位长质量kgpm=" & CStr(Text29.Text) & ",接头抗内压强度MPa=" & Str(Val(Text26.Text)) & "," _
                & " 接头抗拉强度kN=" & CStr(Text27.Text) & ",备注='" & Trim(Text28.Text) & "'," & " 钻头尺寸mm=" & CStr(Text25.Text) & ",屈服极限MPa =" & CStr(Text22.Text) & "," _
                & " 固井时井液密度=" & CStr(TextBox1.Text) & ",固井水泥浆密度 =" & CStr(TextBox2.Text) _
                & " where 井号='" & well_name & "' and 套管层数= " & Val(Text7.Text) & " and 套管段数 = " & Val(Text14.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
        Else
            msg_prompt = "是否要建立新的一层段套管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                RECreader.Close()
                cn.Close()
                Exit Sub
            End If
            '*********************一旦建立了新的套管层段，就清空“套管磨损分析参数表、冲蚀预测分析结果表”******************************************
            SQL_command = "DELETE * from 套管磨损分析参数表 where 井号='" & well_name & "'and 层= " & CStr(Text7.Text) & " and 段 = " & CStr(Text14.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
            'SQL_command = "DELETE * from 冲蚀预测分析结果表 where 井号='" & well_name & "'"
            'Set SQL_rst = ExecuteSQL(SQL_command, msg_prompt, 2)
            '*********************一旦建立了新的套管层段，就清空“套管磨损分析参数表、冲蚀预测分析结果表”******************************************
            SQL_command = "insert into 套管数据表(" _
                    & "套管规格,   套管层数,   套管类型,     套管段数,         是否回接," _
                    & "套管外径mm, 弹性模量MPa,悬挂深度m,    套管下深m,        水泥返深m," _
                    & "完钻深度m,  套管壁厚mm, 套管钢级,     抗内压强度MPa,    抗拉强度kN," _
                    & "抗挤强度MPa,泊松比,     屈服极限MPa,  井号,            套管名称," _
                    & "扣型,       钻头尺寸mm,单位长质量kgpm,接头抗内压强度MPa,接头抗拉强度kN," _
                    & "备注,固井时井液密度,固井水泥浆密度) values ('" _
                    & Text12.Text & "'," & CStr(Text7.Text) & ",'" & Combo1.Text & "'," & CStr(Text14.Text) & ",'" & Combo2.Text & "'," _
                    & CStr(Text9.Text) & "," & CStr(Text8.Text) & "," & CStr(Text10.Text) & "," & CStr(Text11.Text) & "," & CStr(Text13.Text) & "," _
                    & CStr(Text15.Text) & "," & CStr(Text16.Text) & ",'" & Text17.Text & "'," & CStr(Text19.Text) & "," & CStr(Text20.Text) & "," _
                    & CStr(Text21.Text) & "," & CStr(Text18.Text) & "," & CStr(Text22.Text) & ",'" & well_name & "','" & Text23.Text & "','" _
                    & Text24.Text & "'," & CStr(Text25.Text) & "," & CStr(Text29.Text) & "," & CStr(Text26.Text) & "," & CStr(Text27.Text) & ",'" _
                    & Text28.Text & "'," & CStr(TextBox1.Text) & "," & CStr(TextBox2.Text) & ")"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
        End If
        RECreader.Close()
        cn.Close()
        Call write_chanshu("套管磨损计算-是否重算磨损量", "是")
        Call save_tglib()
        '绘制井身结构图,不含管柱
        Call drawWellStruction(Picture1, 1)
        Call tree_change()
        msg_prompt = "套管数据保存完成。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '检查刚输入的套管数据在套管库中是否存在或数据有变化，根据存在及变化情况进行添加或更新
    '*********************************************************************************************************************************************
    Private Sub save_tglib()
        Dim bzh As String
        Dim data_chged As Boolean
        Dim tishi As String

        cn = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn.Open()
        SQL_command = "select * from 套管 where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            data_chged = False
            '************************************************************************************************
            '    若套管库里的数据与界面中的数据不一致，应酌情更新套管库数据：
            '   （1）套管库里的某一数据为空或0，而界面中不为空或0，应自动更新套管库数据。
            '   （2）套管库及界面中的某一数据均不为空或0，但不相等，应提示用户是否用界面数据更新套管库数据。
            '************************************************************************************************
            '   （1）套管库里的某一数据为空或0，而界面中不为空或0，应自动更新套管库数据。
            If RECreader.Item("套管规格").ToString = "" And Trim(Text12.Text) <> "" Then
                SQL_command = "update 套管 set [套管规格]='" & Trim(Text12.Text) & "'" _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If

            If (RECreader.Item("屈服极限(MPa)").ToString = "" Or Val(RECreader.Item("屈服极限(MPa)").ToString) = 0.0) And Val(Text22.Text) <> 0.0# Then
                SQL_command = "update 套管 set [屈服极限(MPa)]=" & CStr(Text22.Text) & "'" _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("单位长质量(kg/m)").ToString = "" Or Val(RECreader.Item("单位长质量(kg/m)").ToString) = 0.0) And Val(Text29.Text) <> 0.0# Then
                SQL_command = "update 套管 set [单位长质量(kg/m)]=" & CStr(Text29.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("抗挤强度(MPa)").ToString = "" Or Val(RECreader.Item("抗挤强度(MPa)").ToString) = 0.0) And Val(Text21.Text) <> 0.0# Then
                SQL_command = "update 套管 set [抗挤强度(MPa)]=" & CStr(Text21.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("抗内压强度(MPa)").ToString = "" Or Val(RECreader.Item("抗内压强度(MPa)").ToString) = 0.0) And Val(Text19.Text) <> 0.0# Then
                SQL_command = "update 套管 set [抗内压强度(MPa)]=" & CStr(Text19.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("抗拉强度(kN)").ToString = "" Or Val(RECreader.Item("抗拉强度(kN)").ToString) = 0.0) And Val(Text20.Text) <> 0.0# Then
                SQL_command = "update 套管 set [抗拉强度(kN)]=" & CStr(Text20.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("接头抗内压强度(MPa)").ToString = "" Or Val(RECreader.Item("接头抗内压强度(MPa)").ToString) = 0.0) And Val(Text26.Text) <> 0.0# Then
                SQL_command = "update 套管 set [接头抗内压强度(MPa)]=" & CStr(Text26.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If (RECreader.Item("接头抗拉强度(kN)").ToString = "" Or Val(RECreader.Item("接头抗拉强度(kN)").ToString) = 0.0) And Val(Text27.Text) <> 0.0# Then
                SQL_command = "update 套管 set [接头抗拉强度(kN)]=" & CStr(Text27.Text) _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                data_chged = True
            End If
            If data_chged = True Then
                bzh = (Today) & "输入" & well_name & "井数据时最后更新。"
                SQL_command = "update 套管 set [备注]='" & bzh & "'" _
                    & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
            End If
            '   （2）套管库及界面中的某一数据均不为空或0，但不相等，应提示用户是否用界面数据更新套管库数据。
            tishi = ""
            If RECreader.Item("套管规格").ToString <> "" And Trim(Text12.Text) <> "" And RECreader.Item("套管规格").ToString <> Text12.Text Then
                tishi = "套管规格：数据库中为：" & RECreader.Item("套管规格").ToString & "，界面中为：" & Trim(Text12.Text)
            End If
            If RECreader.Item("屈服极限(MPa)").ToString <> "" And Val(Text22.Text) <> 0.0# And RECreader.Item("屈服极限(MPa)").ToString <> Text22.Text Then
                If tishi = "" Then
                    tishi = "屈服极限：数据库中=" & CStr(RECreader.Item("屈服极限(MPa)").ToString) & "，界面中=" & CStr(Text22.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "屈服极限：数据库中=" & CStr(RECreader.Item("屈服极限(MPa)").ToString) & "，界面中=" & CStr(Text22.Text)
                End If
            End If
            If RECreader.Item("单位长质量(kg/m)").ToString <> "" And Val(Text29.Text) <> 0.0# And RECreader.Item("单位长质量(kg/m)").ToString <> Text29.Text Then
                If tishi = "" Then
                    tishi = "单位长质量：数据库中=" & CStr(RECreader.Item("单位长质量(kg/m)").ToString) & "，界面中=" & CStr(Text29.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "单位长质量：数据库中=" & CStr(RECreader.Item("单位长质量(kg/m)").ToString) & "，界面中=" & CStr(Text29.Text)
                End If
            End If
            If RECreader.Item("抗挤强度(MPa)").ToString <> "" And Val(Text21.Text) <> 0.0# And RECreader.Item("抗挤强度(MPa)").ToString <> Text21.Text Then
                If tishi = "" Then
                    tishi = "抗挤强度：数据库中=" & CStr(RECreader.Item("抗挤强度(MPa)").ToString) & "，界面中=" & CStr(Text21.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "抗挤强度：数据库中=" & CStr(RECreader.Item("抗挤强度(MPa)").ToString) & "，界面中=" & CStr(Text21.Text)
                End If
            End If
            If RECreader.Item("抗内压强度(MPa)").ToString <> "" And Val(Text19.Text) <> 0.0# And RECreader.Item("抗内压强度(MPa)").ToString <> Text19.Text Then
                If tishi = "" Then
                    tishi = "管体抗内压强度：数据库中=" & CStr(RECreader.Item("抗内压强度(MPa)").ToString) & "，界面中=" & CStr(Text19.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "管体抗内压强度：数据库中=" & CStr(RECreader.Item("抗内压强度(MPa)").ToString) & "，界面中=" & CStr(Text19.Text)
                End If
            End If
            If RECreader.Item("抗拉强度(kN)").ToString <> "" And Val(Text20.Text) <> 0.0# And RECreader.Item("抗拉强度(kN)").ToString <> Text20.Text Then
                If tishi = "" Then
                    tishi = "管体抗拉强度：数据库中=" & CStr(RECreader.Item("抗拉强度(kN)").ToString) & "，界面中=" & CStr(Text20.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "管体抗拉强度：数据库中=" & CStr(RECreader.Item("抗拉强度(kN)").ToString) & "，界面中=" & CStr(Text20.Text)
                End If
            End If
            If RECreader.Item("接头抗内压强度(MPa)").ToString <> "" And Val(Text26.Text) <> 0.0# And RECreader.Item("接头抗内压强度(MPa)").ToString <> Text26.Text Then
                If tishi = "" Then
                    tishi = "接头抗内压强度：数据库中=" & CStr(RECreader.Item("接头抗内压强度(MPa)").ToString) & "，界面中=" & CStr(Text26.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "接头抗内压强度：数据库中=" & CStr(RECreader.Item("接头抗内压强度(MPa)").ToString) & "，界面中=" & CStr(Text26.Text)
                End If
            End If
            If RECreader.Item("接头抗拉强度(kN)").ToString <> "" And Val(Text27.Text) <> 0.0# And RECreader.Item("接头抗拉强度(kN)").ToString <> Text27.Text Then
                If tishi = "" Then
                    tishi = "接头抗拉强度：数据库中=" & CStr(RECreader.Item("接头抗拉强度(kN)").ToString) & "，界面中=" & CStr(Text27.Text)
                Else
                    tishi = tishi & Chr(13) & Chr(10) & "接头抗拉强度：数据库中=" & CStr(RECreader.Item("接头抗拉强度(kN)").ToString) & "，界面中=" & CStr(Text27.Text)
                End If
            End If
            If tishi <> "" Then
                msg_prompt = "套管数据库中下列数据与界面中的值不同：" & Chr(13) & Chr(10) & tishi & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "是否用界面中的数据更新套管数据库？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return = 6 Then
                    bzh = (Today) & "输入" & well_name & "井数据时最后更新。"
                    SQL_command = "update 套管 set [备注]='" & bzh & "',[套管规格]='" & Trim(Text12.Text) & "'," & " [屈服极限(MPa)]=" & CStr(Text22.Text) & ",[单位长质量(kg/m)]=" & CStr(Text29.Text) & "," & " [抗挤强度(MPa)]=" & CStr(Text21.Text) & ",[抗内压强度(MPa)]=" & CStr(Text19.Text) & "," & " [抗拉强度(kN)]=" & CStr(Text20.Text) & ",[接头抗内压强度(MPa)]=" & CStr(Text26.Text) & "," & " [接头抗拉强度(kN)]=" & CStr(Text27.Text) & " where [套管外径(mm)]=" & Str(Val(Text9.Text)) & " and [壁厚(mm)]= " & Str(Val(Text16.Text)) & " and [钢级]='" & Text17.Text & "' and [扣型]='" & Text24.Text & "'"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                    data_chged = True
                End If
            End If
            RECreader.Close()
            cn.Close()
            If data_chged = True Then
                Call my_gridview1_flash()
            End If
        Else
            msg_prompt = "是否将此套管加入到套管数据库中？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                bzh = (Today) & "输入" & well_name & "井数据时加入。"
                SQL_command = "insert into 套管(" _
                    & "[套管规格],   [钢级],   [套管外径(mm)], [壁厚(mm)], [抗挤强度(MPa)]," _
                    & "[抗内压强度(MPa)],[抗拉强度(kN)],[屈服极限(MPa)],[扣型],[备注]," _
                    & "[单位长质量(kg/m)],[接头抗内压强度(MPa)],[接头抗拉强度(kN)]) values (" & "'" _
                    & Trim(Text12.Text) & "','" & Trim(Text17.Text) & "'," & CStr(Text9.Text) & "," & CStr(Text16.Text) & "," & CStr(Text21.Text) & "," _
                    & CStr(Text19.Text) & "," & CStr(Text20.Text) & "," & CStr(Text22.Text) & ",'" & Trim(Text24.Text) & "','" & bzh & "'," _
                    & CStr(Text29.Text) & "," & Str(Val(Text26.Text)) & "," & Str(Val(Text27.Text)) & ")"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                EXECOleDbCommand.ExecuteNonQuery()
                RECreader.Close()
                cn.Close()
                Call my_gridview1_flash()
            Else
                RECreader.Close()
                cn.Close()
            End If
        End If
        cn.Dispose()
    End Sub
    '*********************************************************************************************************************************************
    '点击“↑选用套管[&A]”按钮事件
    '*********************************************************************************************************************************************
    Private Sub cmdSel_casing_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSel_casing.Click
        Call myrefresh2()
    End Sub
    '*********************************************************************************************************************************************
    '双击DataGridView1单元格任意位置事件
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call myrefresh2()
    End Sub
    '*********************************************************************************************************************************************
    '用选中的套管库中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub myrefresh2()
        If IsNothing(Me.BindingSource1.Current) Then
            msg_prompt = "请从列表中选择套管。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Else
            If IsDBNull(Me.BindingSource1.Current(0)) Then
                msg_prompt = "请从列表中选择套管。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Else
                Text9.Text = IIf(BindingSource1.Current("套管外径(mm)").ToString = "", "0.0", BindingSource1.Current("套管外径(mm)").ToString)
                Text16.Text = IIf(BindingSource1.Current("壁厚(mm)").ToString = "", "0.0", BindingSource1.Current("壁厚(mm)").ToString)
                Text12.Text = BindingSource1.Current("套管规格").ToString
                Text17.Text = BindingSource1.Current("钢级").ToString
                Text19.Text = IIf(BindingSource1.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗内压强度(MPa)").ToString)
                Text20.Text = IIf(BindingSource1.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("抗拉强度(kN)").ToString)
                Text21.Text = IIf(BindingSource1.Current("抗挤强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗挤强度(MPa)").ToString)
                Text22.Text = IIf(BindingSource1.Current("屈服极限(MPa)").ToString = "", "0.0", BindingSource1.Current("屈服极限(MPa)").ToString)
                Text24.Text = BindingSource1.Current("扣型").ToString
                Text28.Text = BindingSource1.Current("备注").ToString
                Text29.Text = IIf(BindingSource1.Current("单位长质量(kg/m)").ToString = "", "0.0", BindingSource1.Current("单位长质量(kg/m)").ToString)
                Text26.Text = IIf(BindingSource1.Current("接头抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("接头抗内压强度(MPa)").ToString)
                Text27.Text = IIf(BindingSource1.Current("接头抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("接头抗拉强度(kN)").ToString)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击treeview节点后，显示相应层段套管数据
    '*********************************************************************************************************************************************
    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim num1 As Integer
        Dim numLayer As Integer
        Dim numSegment As Short
        num1 = InStr(TreeView1.SelectedNode.Name, "-")
        If num1 <> 0 Then
            numLayer = Val(Mid(TreeView1.SelectedNode.Name, num1 - 1))
            numSegment = Val(Mid(TreeView1.SelectedNode.Name, num1 + 1))
            If numLayer <> 0 Then
                '用ADO.NET给套管数据赋值
                cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn.Open()
                SQL_command = "select * from 套管数据表 where 井号='" & well_name & "' and 套管层数 = " & Val(numLayer) & " and 套管段数 = " & Val(CStr(numSegment))
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If RECreader.Read Then
                    Text12.Text = RECreader.Item("套管规格").ToString
                    Text7.Text = RECreader.Item("套管层数").ToString
                    Combo1.Text = RECreader.Item("套管类型").ToString
                    Combo2.Text = RECreader.Item("是否回接").ToString
                    Text9.Text = RECreader.Item("套管外径mm").ToString
                    Text10.Text = RECreader.Item("悬挂深度m").ToString
                    Text11.Text = RECreader.Item("套管下深m").ToString
                    Text13.Text = RECreader.Item("水泥返深m").ToString
                    Text14.Text = RECreader.Item("套管段数").ToString
                    Text15.Text = RECreader.Item("完钻深度m").ToString
                    Text16.Text = RECreader.Item("套管壁厚mm").ToString
                    Text17.Text = RECreader.Item("套管钢级").ToString
                    Text19.Text = RECreader.Item("抗内压强度MPa").ToString
                    Text20.Text = RECreader.Item("抗拉强度kN").ToString
                    Text21.Text = RECreader.Item("抗挤强度MPa").ToString
                    Text22.Text = RECreader.Item("屈服极限MPa").ToString
                    Text8.Text = RECreader.Item("弹性模量MPa").ToString
                    Text18.Text = RECreader.Item("泊松比").ToString
                    Text23.Text = RECreader.Item("套管名称").ToString
                    Text24.Text = RECreader.Item("扣型").ToString
                    Text25.Text = RECreader.Item("钻头尺寸mm").ToString
                    Text26.Text = RECreader.Item("接头抗内压强度MPa").ToString
                    Text27.Text = RECreader.Item("接头抗拉强度kN").ToString
                    Text28.Text = RECreader.Item("备注").ToString
                    Text29.Text = RECreader.Item("单位长质量kgpm").ToString
                    TextBox1.Text = RECreader.Item("固井时井液密度").ToString
                    TextBox2.Text = RECreader.Item("固井水泥浆密度").ToString
                End If
                RECreader.Close()
                cn.Close()
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击“绘图”按钮事件
    '*********************************************************************************************************************************************
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '绘制井身结构图,不含管柱
        Call drawWellStruction(Picture1, 1)
        
    End Sub
End Class