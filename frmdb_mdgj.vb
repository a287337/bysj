Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_mdgj
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库锚定工具数据输入、修改、管理窗体的说明
    '                                                                                                       秦彦斌 2022年01月19日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200131，基础数据表维护模块升级到VS2008的第5个。
    '    20220119，重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private mdgj_Table As DataTable = b_dst.Tables.Add("mdgj_Table")
    Private mdgjxfqx_Table As DataTable = b_dst.Tables.Add("mdgjxfqx_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_mdgj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_mdgj_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text13.Text = ""
        Combo1.Text = ""
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text8.Text = CStr(0.0#)
        Text1.Text = ""
        Text5.Text = CStr(0.0#)
        Text9.Text = ""
        Text12.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text16.Text = ""
        Text7.Text = CStr(0.0#)
        Text17.Text = ""
        Text4.Text = CStr(0.0#)
        Text19.Text = CStr(0.0#)
        Text21.Text = CStr(0.0#)
        Text20.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text22.Text = ""
        Text23.Text = CStr(0)
        Text24.Text = CStr(0.0#)
        Text25.Text = CStr(0.0#)
        Call fill_mdgjgrid()
        Call fill_mdgjxfqxgrid()
    End Sub
    '*********************************************************************************************************************************************
    '锚定工具数据列表DataGridView1设置并填充函数fill_mdgjgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_mdgjgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select 型号,工具名称,坐卡方式,最大外径mm,最小通径mm,总体长度m,重量kg,工作压力MPa," _
                    & "温度范围下℃,温度范围上℃,小坐卡压力MPa,大坐卡压力MPa,上端扣型,下端扣型,生产厂家,最大锚定力kN," _
                    & "最大解锚力kN,抗内压强度MPa,抗外压强度MPa,抗拉强度kN,备注 " & " from 锚定工具表 order by 最大外径mm desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        mdgj_Table.Clear()
        ad.Fill(mdgj_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = mdgj_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.Columns(0).Width = 150
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中锚定工具列表中的某行。
    '用选中的锚定工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("型号"))) Then
            Text13.Text = BindingSource1.Current("型号").ToString
            Combo1.Text = BindingSource1.Current("坐卡方式").ToString
            Text2.Text = IIf(BindingSource1.Current("最大外径mm").ToString = "", "0.0", BindingSource1.Current("最大外径mm").ToString)
            Text3.Text = IIf(BindingSource1.Current("最小通径mm").ToString = "", "0.0", BindingSource1.Current("最小通径mm").ToString)
            Text8.Text = IIf(BindingSource1.Current("总体长度m").ToString = "", "0.0", BindingSource1.Current("总体长度m").ToString)
            Text1.Text = BindingSource1.Current("工具名称").ToString
            Text5.Text = IIf(BindingSource1.Current("重量kg").ToString = "", "0.0", BindingSource1.Current("重量kg").ToString)    '表示质量，存的就是质量，不转换。表设计之初的问题。
            Text9.Text = BindingSource1.Current("生产厂家").ToString
            Text12.Text = IIf(BindingSource1.Current("温度范围下℃").ToString = "", "0.0", BindingSource1.Current("温度范围下℃").ToString)
            Text11.Text = IIf(BindingSource1.Current("温度范围上℃").ToString = "", "0.0", BindingSource1.Current("温度范围上℃").ToString)
            Text16.Text = BindingSource1.Current("上端扣型").ToString
            Text7.Text = IIf(BindingSource1.Current("小坐卡压力MPa").ToString = "", "0.0", BindingSource1.Current("小坐卡压力MPa").ToString)
            Text17.Text = BindingSource1.Current("下端扣型").ToString
            Text4.Text = IIf(BindingSource1.Current("大坐卡压力MPa").ToString = "", "0.0", BindingSource1.Current("大坐卡压力MPa").ToString)
            Text19.Text = IIf(BindingSource1.Current("最大锚定力kN").ToString = "", "0.0", BindingSource1.Current("最大锚定力kN").ToString)
            Text21.Text = IIf(BindingSource1.Current("最大解锚力kN").ToString = "", "0.0", BindingSource1.Current("最大解锚力kN").ToString)
            Text20.Text = IIf(BindingSource1.Current("工作压力MPa").ToString = "", "0.0", BindingSource1.Current("工作压力MPa").ToString)
            Text10.Text = IIf(BindingSource1.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗内压强度MPa").ToString)
            Text14.Text = IIf(BindingSource1.Current("抗外压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗外压强度MPa").ToString)
            Text6.Text = IIf(BindingSource1.Current("抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("抗拉强度kN").ToString)
            Text22.Text = BindingSource1.Current("备注").ToString
            Call fill_mdgjxfqxgrid()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '锚定工具信封曲线数据列表DataGridView2设置并填充函数fill_mdgjxfqxgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_mdgjxfqxgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select 序号,轴力kN,压差MPa from 锚定工具载荷性能图参数表 where 0>1"
        If Not IsNothing(Me.BindingSource1.Current) Then
            If Not IsDBNull(Me.BindingSource1.Current("型号")) Then
                SQL_command = "select 序号,轴力kN,压差MPa from 锚定工具载荷性能图参数表 " _
                        & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                        & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "' order by 序号"
            End If
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        mdgjxfqx_Table.Clear()
        ad.Fill(mdgjxfqx_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource2.DataSource = mdgjxfqx_Table
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
    '主界面点击"删除[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 锚定工具表 " _
                    & " where 最大外径mm=" & CStr(Text2.Text) & " and 最小通径mm= " & CStr(Text3.Text) & " and 总体长度m=" & CStr(Text8.Text) _
                    & " and  坐卡方式='" & Combo1.Text & "' " & " and 型号='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的锚定工具数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 锚定工具表 " _
                        & " where 最大外径mm=" & CStr(Text2.Text) & " and 最小通径mm= " & CStr(Text3.Text) & " and 总体长度m=" & CStr(Text8.Text) _
                        & " and  坐卡方式='" & Combo1.Text & "' " & " and 型号='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "delete * from 锚定工具载荷性能图参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                        & " and [长度m]=" & CStr(Text8.Text) & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                basedb_chg = True
                msg_prompt = "锚定工具数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的锚定工具！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_mdgjgrid()
            Call fill_mdgjxfqxgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "关键数据有错，请选择好要删除的锚定工具！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        Dim nffxj As String
        Dim mdgjxfqx_row As DataRow
        On Error GoTo errhandler
        If Trim(Text13.Text) = "" Then
            msg_prompt = "锚定工具型号是关键参数，请输入。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "" Then
            msg_prompt = "请输入或选择锚定工具的坐卡方式。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入锚定工具最大外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入锚定工具最小通径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入锚定工具长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Trim(Text1.Text) = "" Then
            msg_prompt = "请输入锚定工具名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text5.Text) = 0 Then
            msg_prompt = "请输入锚定工具质量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 锚定工具表 " _
                    & " where 最大外径mm=" & CStr(Text2.Text) & " and 最小通径mm= " & CStr(Text3.Text) & " and 总体长度m=" & CStr(Text8.Text) _
                    & " and  坐卡方式='" & Combo1.Text & "' " & " and 型号='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的锚定工具数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into 锚定工具表 (型号,坐卡方式,最大外径mm,最小通径mm,总体长度m," _
                        & "工具名称,重量kg,生产厂家,温度范围下℃,温度范围上℃," _
                        & "上端扣型,小坐卡压力MPa,下端扣型,大坐卡压力MPa,最大锚定力kN," _
                        & "最大解锚力kN,工作压力MPa,抗内压强度MPa,抗外压强度MPa,抗拉强度kN," _
                        & "备注) values (" _
                        & "'" & Text13.Text & "','" & Combo1.Text & "'," & CStr(Text2.Text) & "," & CStr(Text3.Text) & "," & CStr(Text8.Text) & "," _
                        & "'" & Text1.Text & "'," & CStr(Text5.Text) & ",'" & Text9.Text & "'," & CStr(Text12.Text) & "," & CStr(Text11.Text) & "," _
                        & "'" & Text16.Text & "'," & CStr(Text7.Text) & ",'" & Text17.Text & "'," & CStr(Text4.Text) & "," & CStr(Text19.Text) & "," _
                        & CStr(Text21.Text) & "," & CStr(Text20.Text) & "," & CStr(Text10.Text) & "," & CStr(Text14.Text) & "," & CStr(Text6.Text) & "," _
                        & "'" & Text22.Text & " ')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "delete * from 锚定工具载荷性能图参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                        & " and [长度m]=" & CStr(Text8.Text) & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                For Each mdgjxfqx_row In mdgjxfqx_Table.Select
                    SQL_command = "insert into 锚定工具载荷性能图参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,坐卡方式)  values (" _
                            & mdgjxfqx_row.Item("序号").ToString & "," & mdgjxfqx_row.Item("轴力kN").ToString & "," & mdgjxfqx_row.Item("压差MPa").ToString & ",'" & Text13.Text & "'," & CStr(Text2.Text) & "," _
                            & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Combo1.Text & "')"
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
                SQL_command = "update 锚定工具表 set " _
                        & " 工具名称='" & Text1.Text & "'," & " 重量kg=" & CStr(Text5.Text) & "," & " 生产厂家='" & Text9.Text & "'," & " 温度范围下℃=" & CStr(Text12.Text) & "," _
                        & " 温度范围上℃=" & CStr(Text11.Text) & "," & " 上端扣型='" & Text16.Text & "'," & " 小坐卡压力MPa=" & CStr(Text7.Text) & "," & " 下端扣型='" & Text17.Text & "'," _
                        & " 大坐卡压力MPa=" & CStr(Text4.Text) & "," & " 最大锚定力kN=" & CStr(Text19.Text) & "," & " 最大解锚力kN=" & CStr(Text21.Text) & "," & " 工作压力MPa=" & CStr(Text20.Text) & "," _
                        & " 抗内压强度MPa=" & CStr(Text10.Text) & "," & " 抗外压强度MPa=" & CStr(Text14.Text) & "," & " 抗拉强度kN=" & CStr(Text6.Text) & "," & " [备注]='" & Text22.Text & "'" _
                        & " where 最大外径mm=" & CStr(Text2.Text) & " and 最小通径mm= " & CStr(Text3.Text) & " and 总体长度m=" & CStr(Text8.Text) _
                        & " and  坐卡方式='" & Combo1.Text & "' " & " and 型号='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_mdgjgrid()
            Call fill_mdgjxfqxgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击移除[&E]"按钮
    '*********************************************************************************************************************************************
    Private Sub zhxnqx_delete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zhxnqx_delete.Click
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 锚定工具载荷性能图参数表 " _
                & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            SQL_command = "delete * from 锚定工具载荷性能图参数表 " _
                    & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                    & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        Else
            msg_prompt = "请选择好要删除的锚定工具载荷性能图参数！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_mdgjxfqxgrid()
            Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "关键数据有错，请选择好要删除的锚定工具载荷性能图参数！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"添加[&A]"按钮
    '*********************************************************************************************************************************************
    Private Sub zhxnqx_save_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zhxnqx_save.Click
        Dim basedb_chg As Boolean
        Dim textLegal As Short
        On Error GoTo errhandler
        If IsNothing(Me.BindingSource1.Current) Or IsDBNull(Me.BindingSource1.Current("型号")) Then
            msg_prompt = "请选择锚定工具。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Text13.Text) = "" Then
            msg_prompt = "锚定工具型号是关键参数，请输入。"
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
        SQL_command = "select * from 锚定工具表 " _
                & " where 最大外径mm=" & CStr(Text2.Text) & " and 最小通径mm= " & CStr(Text3.Text) & " and 总体长度m=" & CStr(Text8.Text) _
                & " and  坐卡方式='" & Combo1.Text & "' " & " and 型号='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            RECreader.Close()
            cn_basedb.Close()
            cn_basedb.Dispose()
            msg_prompt = "请选择锚定工具或者先保存锚定工具技术参数后选择该锚定工具，然后添加载荷性能曲线数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        SQL_command = "select * from 锚定工具载荷性能图参数表 " _
                & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                & " and  [坐卡方式]='" & Combo1.Text & "' " & " and [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            '插入新记录
            SQL_command = "insert into 锚定工具载荷性能图参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,坐卡方式)  values (" _
                    & CStr(Text23.Text) & "," & CStr(Text24.Text) & "," & CStr(Text25.Text) & ",'" & Text13.Text & "'," & CStr(Text2.Text) & "," _
                    & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Combo1.Text & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        Else
            SQL_command = "update 锚定工具载荷性能图参数表 set " _
                    & " [轴力kN]=" & CStr(Text24.Text) & "," & " [压差MPa]=" & CStr(Text25.Text) & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                    & " and [长度m]=" & CStr(Text8.Text) & " and  [坐卡方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            basedb_chg = True
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_mdgjxfqxgrid()
            Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
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
        Dim mdgjxfqx_row As DataRow
        ReDim zl(mdgjxfqx_Table.Rows.Count - 1)
        ReDim yc(mdgjxfqx_Table.Rows.Count - 1)
        i = 0
        For Each mdgjxfqx_row In mdgjxfqx_Table.Select
            zl(i) = Val(mdgjxfqx_row.Item("轴力kN").ToString)
            yc(i) = Val(mdgjxfqx_row.Item("压差MPa").ToString)
            i = i + 1
        Next
        Call draw_xfqx(iPlotX1, zl, yc, mdgjxfqx_Table.Rows.Count, "压差(MPa)", "轴力(kN)") '
    End Sub
End Class