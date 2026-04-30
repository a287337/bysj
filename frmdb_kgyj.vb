Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_kgyj
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库开关工具数据输入、修改、管理窗体的说明
    '                                                                                                       秦彦斌 2022年01月19日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200201,基础数据表维护模块升级到VS2008的第6个。
    '    20220119，重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private kggj_Table As DataTable = b_dst.Tables.Add("kggj_Table")
    Private kggjxfqx_Table As DataTable = b_dst.Tables.Add("kggjxfqx_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_kgyj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_kgyj_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text11.Text = ""
        Text9.Text = ""
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text8.Text = CStr(0.0#)
        Text1.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text4.Text = CStr(0.0#)
        Text16.Text = CStr(0.0#)
        Text18.Text = CStr(0.0#)
        Text17.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text7.Text = CStr(0.0#)
        Text12.Text = ""
        Text5.Text = CStr(0.0#)
        Call fill_kglxlist()
        Call fill_kggjgrid()
        Call fill_kggjxfqxgrid()
    End Sub
    '*********************************************************************************************************************************************
    '开关类型列表填充
    ' "select DISTINCT 开关类型 from 开关元件 order by 开关类型"
    '*********************************************************************************************************************************************
    Private Sub fill_kglxlist()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command =  "select DISTINCT 开关类型 from 开关元件 order by 开关类型"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        ' Shutdown the painting of the ListBox as items are added.
        ListBox1.BeginUpdate()
        ListBox1.Items.Clear()
        While RECreader.Read
            ListBox1.Items.Add(RECreader.Item("开关类型").ToString())
        End While
        ' Allow the ListBox to repaint and display the new items.
        ListBox1.EndUpdate()
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
    End Sub
    '*********************************************************************************************************************************************
    '开关工具数据列表DataGridView1设置并填充函数fill_kggjgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_kggjgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        If Text11.Text = "" Then
            SQL_command = "select [开关类型],[型号],[名称],[外径(mm)],[内径(mm)],[长度(m)],[重量(kg)]," _
                        & "[压力等级MPa],[极限压差(MPa)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[抗拉强度(kN)]," _
                        & "[温度范围下℃],[温度范围上℃],[上端扣型],[下端扣型],[生产厂家],[备注]" _
                        & " from 开关元件 order by [外径(mm)] desc"
        Else
            SQL_command = "select [开关类型],[型号],[名称],[外径(mm)],[内径(mm)],[长度(m)],[重量(kg)]," _
                        & "[压力等级MPa],[极限压差(MPa)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[抗拉强度(kN)]," _
                        & "[温度范围下℃],[温度范围上℃],[上端扣型],[下端扣型],[生产厂家],[备注]" _
                        & " from 开关元件 where 开关类型='" & Trim(Text11.Text) & "' order by [外径(mm)] desc"
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        kggj_Table.Clear()
        ad.Fill(kggj_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = kggj_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.Columns(0).Width = 78
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 150
        DataGridView1.Columns(3).Width = 70
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Width = 70
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中开关工具列表中的某行。
    '用选中的开关工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If Not IsNothing(Me.BindingSource1.Current) Then
            If Not IsDBNull(Me.BindingSource1.Current("型号")) Then
                Text11.Text = BindingSource1.Current("开关类型").ToString
                Text9.Text = BindingSource1.Current("型号").ToString
                Text2.Text = IIf(BindingSource1.Current("外径(mm)").ToString = "", "0.0", BindingSource1.Current("外径(mm)").ToString)
                Text3.Text = IIf(BindingSource1.Current("内径(mm)").ToString = "", "0.0", BindingSource1.Current("内径(mm)").ToString)
                Text8.Text = IIf(BindingSource1.Current("长度(m)").ToString = "", "0.0", BindingSource1.Current("长度(m)").ToString)
                Text1.Text = BindingSource1.Current("名称").ToString
                Text13.Text = BindingSource1.Current("生产厂家").ToString
                Text14.Text = BindingSource1.Current("上端扣型").ToString
                Text15.Text = BindingSource1.Current("下端扣型").ToString
                Text4.Text = IIf(BindingSource1.Current("重量(kg)").ToString = "", "0.0", BindingSource1.Current("重量(kg)").ToString)    '表示质量，存的就是质量，不转换。表设计之初的问题。
                Text16.Text = IIf(BindingSource1.Current("温度范围下℃").ToString = "", "0.0", BindingSource1.Current("温度范围下℃").ToString)
                Text18.Text = IIf(BindingSource1.Current("温度范围上℃").ToString = "", "0.0", BindingSource1.Current("温度范围上℃").ToString)
                Text17.Text = IIf(BindingSource1.Current("压力等级MPa").ToString = "", "0.0", BindingSource1.Current("压力等级MPa").ToString)
                Text10.Text = IIf(BindingSource1.Current("极限压差(MPa)").ToString = "", "0.0", BindingSource1.Current("极限压差(MPa)").ToString)
                Text6.Text = IIf(BindingSource1.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗内压强度(MPa)").ToString)
                Text7.Text = IIf(BindingSource1.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗外挤强度(MPa)").ToString)
                Text12.Text = BindingSource1.Current("备注").ToString
                Text5.Text = IIf(BindingSource1.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("抗拉强度(kN)").ToString)
                Call fill_kggjxfqxgrid()
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '开关工具信封曲线数据列表DataGridView2设置并填充函数fill_kggjxfqxgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_kggjxfqxgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select 序号,轴力kN,压差MPa from 开关元件载荷性能图参数表 where 0>1"
        If Not IsNothing(Me.BindingSource1.Current) Then
            If Not IsDBNull(Me.BindingSource1.Current("型号")) Then
                SQL_command = "select 序号,轴力kN,压差MPa from 开关元件载荷性能图参数表 " _
                        & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                        & " and  [开关类型]='" & Text11.Text & "' " & " and  [型号]='" & Text9.Text & "'" & " order by 序号"
            End If
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        kggjxfqx_Table.Clear()
        ad.Fill(kggjxfqx_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource2.DataSource = kggjxfqx_Table
        DataGridView2.ClearSelection()
        DataGridView2.DataSource = BindingSource2
        DataGridView2.ResetBindings()
        DataGridView2.AutoGenerateColumns = True
        DataGridView2.AllowUserToAddRows = False
        DataGridView2.AllowUserToDeleteRows = False
        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.MultiSelect = False
        DataGridView2.RowHeadersWidth = 24
        DataGridView2.Columns(0).Width = (DataGridView2.Width - 26) * 0.125
        DataGridView2.Columns(1).Width = (DataGridView2.Width - 26) * 0.4
        DataGridView2.Columns(2).Width = (DataGridView2.Width - 26) * 0.4
        DataGridView2.ReadOnly = True
        DataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView2.Refresh()
        DataGridView2.Show()
        Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView2，在单元格的任何部分被单击时事件-选中信封曲线列表中的某行。
    '用选中的信封曲线列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        If IsNothing(Me.BindingSource2.Current) Then
            Text23.Text = CStr(0)
            Text24.Text = CStr(0.0#)
            Text25.Text = CStr(0.0#)
        Else
            If Not IsDBNull(Me.BindingSource2.Current("序号")) Then
                Text23.Text = IIf(BindingSource2.Current("序号").ToString = "", "0", BindingSource2.Current("序号").ToString)
                Text24.Text = IIf(BindingSource2.Current("轴力kN").ToString = "", "0", BindingSource2.Current("轴力kN").ToString)
                Text25.Text = IIf(BindingSource2.Current("压差MPa").ToString = "", "0", BindingSource2.Current("压差MPa").ToString)
            End If
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"帮助[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '点击"退出[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击开关工具类型列表
    '*********************************************************************************************************************************************
    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        Text11.Text = ListBox1.Text
        Call fill_kggjgrid()
        Call fill_kggjxfqxgrid()
    End Sub
    '*********************************************************************************************************************************************
    '主界面点击"删除[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 开关元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的开关元件数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 开关元件 " _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                        & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "DELETE * from 开关元件载荷性能图参数表 " _
                        & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                        & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                basedb_chg = True
                msg_prompt = "开关元件数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的开关元件！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_kglxlist()
            Call fill_kggjgrid()
            Call fill_kggjxfqxgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的开关工具！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        Dim nffxj As String
        Dim kggjxfqx_row As DataRow
        On Error GoTo ErrHandler
        If Trim(Text11.Text) = "" Then
            msg_prompt = "请输入或选择开关工具类型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Text9.Text) = "" Then
            msg_prompt = "请输入开关工具型号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入开关元件外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入开关元件内径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入开关元件长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Text1.Text) = "" Then
            msg_prompt = "请输入开关工具名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) = 0 Then
            msg_prompt = "请输入开关元件重量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 开关元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的开关元件数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into 开关元件(" _
                        & "[名称],[外径(mm)],[内径(mm)], [重量(kg)], [长度(m)]," _
                        & "[抗拉强度(kN)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[极限压差(MPa)],[开关类型]," _
                        & "[温度范围下℃],[备注],[温度范围上℃],[压力等级MPa],[型号]," _
                        & "[上端扣型],[下端扣型],[生产厂家]) values (" _
                        & "'" & Text1.Text & "'," & CStr(Text2.Text) & "," & CStr(Text3.Text) & "," & CStr(Text4.Text) & "," & CStr(Text8.Text) & "," _
                        & CStr(Text5.Text) & "," & CStr(Text7.Text) & "," & CStr(Text6.Text) & "," & CStr(Text10.Text) & ",'" & CStr(Text11.Text) & "'," _
                        & CStr(Text16.Text) & ",'" & Text12.Text & "'," & CStr(Text18.Text) & "," & CStr(Text17.Text) & ",'" & Text9.Text & "','" _
                        & Text14.Text & "','" & Text15.Text & "','" & Text13.Text & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "DELETE * from 开关元件载荷性能图参数表 " _
                        & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                        & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                For Each kggjxfqx_row In kggjxfqx_Table.Select
                    SQL_command = "insert into 开关元件载荷性能图参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,开关类型)  values (" _
                            & kggjxfqx_row.Item("序号").ToString & "," & kggjxfqx_row.Item("轴力kN").ToString & "," & kggjxfqx_row.Item("压差MPa").ToString & ",'" & Text9.Text & "'," & CStr(Text2.Text) & "," _
                            & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Text11.Text & "')"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                    EXECOleDbCommand.ExecuteNonQuery()
                Next
                basedb_chg = True
            End If
        Else
            msg_prompt = "是否要保存对数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "update 开关元件 set " _
                        & " [名称]='" & Text1.Text & "', [重量(kg)]=" & CStr(Text4.Text) & "," & " [抗拉强度(kN)]=" & CStr(Text5.Text) & ", [抗外挤强度(MPa)]=" & CStr(Text7.Text) & "," _
                        & " [抗内压强度(MPa)]=" & CStr(Text6.Text) & ", [极限压差(MPa)]=" & CStr(Text10.Text) & "," & " [生产厂家]='" & Text13.Text & "', [备注]='" & Text12.Text & "'," _
                        & " [上端扣型]='" & Text14.Text & "', [下端扣型]='" & Text15.Text & "'," & " [温度范围下℃]=" & CStr(Text16.Text) & ", [温度范围上℃]=" & CStr(Text18.Text) & "," _
                        & " [压力等级MPa]=" & CStr(Text17.Text) _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                        & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_kglxlist()
            Call fill_kggjgrid()
            Call fill_kggjxfqxgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击移除[&E]"按钮
    '*********************************************************************************************************************************************
    Private Sub zhxnqx_delete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zhxnqx_delete.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 开关元件载荷性能图参数表 " _
                & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            SQL_command = "DELETE * from 开关元件载荷性能图参数表 " _
                    & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                    & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        Else
            msg_prompt = "请选择好要删除的开关元件载荷性能图参数！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_kggjxfqxgrid()
            Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的开关元件载荷性能图参数！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)

    End Sub
    '*********************************************************************************************************************************************
    '点击"添加[&A]"按钮
    '*********************************************************************************************************************************************
    Private Sub zhxnqx_save_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zhxnqx_save.Click
        Dim basedb_chg As Boolean
        Dim textLegal As Short
        On Error GoTo ErrHandler
        If IsNothing(Me.BindingSource1.Current) Or IsDBNull(Me.BindingSource1.Current("型号")) Then
            msg_prompt = "请选择开关工具。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Text9.Text) = "" Then
            msg_prompt = "开关工具型号是关键参数，请输入。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text23.Text) <= 0 Then
            msg_prompt = "请输入合理的序号值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If (isTextLegal(Text23, "序号")) = -1 Then Exit Sub
        If isTextLegal(Text24, "轴力") = -1 Then Exit Sub
        If isTextLegal(Text25, "压差") = -1 Then Exit Sub
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 开关元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            RECreader.Close()
            cn_basedb.Close()
            cn_basedb.Dispose()
            msg_prompt = "请选择开关工具或者先保存开关工具技术参数后选择该开关工具，然后添加载荷性能曲线数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        SQL_command = "select * from 开关元件载荷性能图参数表 " _
                & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                & " and  [开关类型]='" & Text11.Text & "' " & " and [型号]='" & Text9.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            '插入新记录
            SQL_command = "insert into 开关元件载荷性能图参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,开关类型)  values (" _
                    & CStr(Text23.Text) & "," & CStr(Text24.Text) & "," & CStr(Text25.Text) & ",'" & Text9.Text & "'," & CStr(Text2.Text) & "," _
                    & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Text11.Text & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        Else
            SQL_command = "update 开关元件载荷性能图参数表 set " _
                    & " [轴力kN]=" & CStr(Text24.Text) & "," & " [压差MPa]=" & CStr(Text25.Text) & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                    & " and [长度m]=" & CStr(Text8.Text) & " and  [开关类型]='" & Text11.Text & "' " & " and  [型号]='" & Text9.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_kggjxfqxgrid()
            Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"绘载荷性能曲线[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub CDMdraw_xfqx_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CDMdraw_xfqx.Click
        Dim i As Short
        Dim zl() As Double
        Dim yc() As Double
        Dim kggjxfqx_row As DataRow
        ReDim zl(kggjxfqx_Table.Rows.Count - 1)
        ReDim yc(kggjxfqx_Table.Rows.Count - 1)
        i = 0
        For Each kggjxfqx_row In kggjxfqx_Table.Select
            zl(i) = Val(kggjxfqx_row.Item("轴力kN").ToString)
            yc(i) = Val(kggjxfqx_row.Item("压差MPa").ToString)
            i = i + 1
        Next
        Call draw_xfqx(iPlotX1, yc, zl, kggjxfqx_Table.Rows.Count, "轴力(kN)", "压差(MPa)") '
    End Sub
End Class