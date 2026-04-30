Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_fgq
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库封隔器数据输入、修改、管理窗体的说明
    '                                                                                                            秦彦斌 2022年1月18日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200130,基础数据表维护模块升级到VS2008的第4个。
    '    20220118，重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private fgq_Table As DataTable = b_dst.Tables.Add("fgq_Table")
    Private fgqxfqx_Table As DataTable = b_dst.Tables.Add("fgqxfqx_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_fgq_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_fgq_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        Text15.Text = ""
        Text18.Text = CStr(0.0#)
        Text16.Text = ""
        Text7.Text = CStr(0.0#)
        Text17.Text = ""
        Text4.Text = CStr(0.0#)
        Text20.Text = CStr(0.0#)
        Text21.Text = CStr(0.0#)
        Text19.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text22.Text = ""
        Text23.Text = CStr(0)
        Text24.Text = CStr(0.0#)
        Text25.Text = CStr(0.0#)
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 封隔器坐封方式"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            Combo1.Items.Add(RECreader.Item("坐封方式").ToString)
        End While
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        Call fill_fgqgrid()
        Call fill_fgqxfqxgrid()
    End Sub
    '*********************************************************************************************************************************************
    '封隔器数据列表DataGridView1设置并填充函数fill_fgqgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_fgqgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [封隔器名称],[型号],[最大外径(mm)],[最小通径(mm)],[长度(m)]," _
                & "[坐封方式],[重量kg],[生产厂家],[温度范围下℃],[温度范围上℃]," _
                & "[主体材料],[主材屈服强度MPa],[上端扣型],[极限承压(MPa)],[下端扣型]," _
                & "[坐封力(kN)],[坐封压差(MPa)],[最小坐封压力MPa],[最大坐封压力MPa],[抗内压强度MPa]," _
                & "[抗外压强度MPa],[极限载荷(kN)],[备注],[能否反洗井] from 封隔器 order by [最大外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        fgq_Table.Clear()
        ad.Fill(fgq_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = fgq_Table
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
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中封隔器列表中的某行。
    '用选中的封隔器列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("型号"))) Then
            Text13.Text = BindingSource1.Current("型号").ToString
            Combo1.Text = BindingSource1.Current("坐封方式").ToString
            Text2.Text = IIf(BindingSource1.Current("最大外径(mm)").ToString = "", "0.0", BindingSource1.Current("最大外径(mm)").ToString)
            Text3.Text = IIf(BindingSource1.Current("最小通径(mm)").ToString = "", "0.0", BindingSource1.Current("最小通径(mm)").ToString)
            Text8.Text = IIf(BindingSource1.Current("长度(m)").ToString = "", "0.0", BindingSource1.Current("长度(m)").ToString)
            Text1.Text = BindingSource1.Current("封隔器名称").ToString
            Text5.Text = IIf(BindingSource1.Current("重量kg").ToString = "", "0.0", BindingSource1.Current("重量kg").ToString)    '表示质量，存的就是质量，不转换。表设计之初的问题。
            Text9.Text = BindingSource1.Current("生产厂家").ToString
            Text12.Text = IIf(BindingSource1.Current("温度范围下℃").ToString = "", "0.0", BindingSource1.Current("温度范围下℃").ToString)
            Text11.Text = IIf(BindingSource1.Current("温度范围上℃").ToString = "", "0.0", BindingSource1.Current("温度范围上℃").ToString)
            Text15.Text = BindingSource1.Current("主体材料").ToString
            Text18.Text = IIf(BindingSource1.Current("主材屈服强度MPa").ToString = "", "0.0", BindingSource1.Current("主材屈服强度MPa").ToString)
            Text16.Text = BindingSource1.Current("上端扣型").ToString
            Text7.Text = IIf(BindingSource1.Current("极限承压(MPa)").ToString = "", "0.0", BindingSource1.Current("极限承压(MPa)").ToString)
            Text17.Text = BindingSource1.Current("下端扣型").ToString
            Text4.Text = IIf(BindingSource1.Current("坐封力(kN)").ToString = "", "0.0", BindingSource1.Current("坐封力(kN)").ToString)
            Text20.Text = IIf(BindingSource1.Current("坐封压差(MPa)").ToString = "", "0.0", BindingSource1.Current("坐封压差(MPa)").ToString)
            Text21.Text = IIf(BindingSource1.Current("最小坐封压力MPa").ToString = "", "0.0", BindingSource1.Current("最小坐封压力MPa").ToString)
            Text19.Text = IIf(BindingSource1.Current("最大坐封压力MPa").ToString = "", "0.0", BindingSource1.Current("最大坐封压力MPa").ToString)
            Text10.Text = IIf(BindingSource1.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗内压强度MPa").ToString)
            Text14.Text = IIf(BindingSource1.Current("抗外压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗外压强度MPa").ToString)
            Text6.Text = IIf(BindingSource1.Current("极限载荷(kN)").ToString = "", "0.0", BindingSource1.Current("极限载荷(kN)").ToString)
            Text22.Text = BindingSource1.Current("备注").ToString
            If BindingSource1.Current("能否反洗井").ToString = "能" Then
                Option1.Checked = True
                Option2.Checked = False
            Else
                Option1.Checked = False
                Option2.Checked = True
            End If
            Call fill_fgqxfqxgrid()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '封隔器信封曲线数据列表DataGridView2设置并填充函数fill_fgqxfqxgrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_fgqxfqxgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select 序号,压差MPa,轴力kN from 封隔器信封曲线参数表 where 0>1"
        If Not IsNothing(Me.BindingSource1.Current) Then
            If Not IsDBNull(Me.BindingSource1.Current("型号")) Then
                SQL_command = "select 序号,压差MPa,轴力kN from 封隔器信封曲线参数表 " _
                        & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) _
                        & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'" _
                        & " order by 序号"
            End If
        End If
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        fgqxfqx_Table.Clear()
        ad.Fill(fgqxfqx_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource2.DataSource = fgqxfqx_Table
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
        On Error GoTo ErrHandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 封隔器 " _
                & " where [最大外径(mm)]=" & CStr(Text2.Text) & " and [最小通径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的封隔器数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 封隔器 " & " where [最大外径(mm)]=" & CStr(Text2.Text) & " and [最小通径(mm)]= " & CStr(Text3.Text) _
                        & " and [长度(m)]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "DELETE  * from 封隔器信封曲线参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                        & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "封隔器数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的封隔器！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_fgqgrid()
            Call fill_fgqxfqxgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的封隔器！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        Dim nffxj As String
        Dim fgqxfqx_row As DataRow
        On Error GoTo ErrHandler
        If Trim(Text13.Text) = "" Then
            msg_prompt = "封隔器型号是关键参数，请输入。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Combo1.Text = "" Then
            msg_prompt = "请输入或选择封隔器的坐封方式。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入封隔器最大外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入封隔器最小通径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入封隔器长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If Trim(Text1.Text) = "" Then
            msg_prompt = "请输入封隔器名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text5.Text) = 0 Then
            msg_prompt = "请输入封隔器质量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Option1.Checked = True Then
            nffxj = "能"
        Else
            nffxj = "不能"
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 封隔器 " & " where [最大外径(mm)]=" & CStr(Text2.Text) & " and [最小通径(mm)]= " & CStr(Text3.Text) _
                & " and [长度(m)]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的封隔器数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into 封隔器 ([型号],[坐封方式],[最大外径(mm)],[最小通径(mm)],[长度(m)]," _
                        & "[封隔器名称],[重量kg],[生产厂家],[温度范围下℃],[温度范围上℃]," _
                        & "[主体材料],[主材屈服强度MPa],[上端扣型],[极限承压(MPa)],[下端扣型]," _
                        & "[坐封力(kN)],[坐封压差(MPa)],[最小坐封压力MPa],[最大坐封压力MPa],[抗内压强度MPa]," _
                        & "[抗外压强度MPa],[极限载荷(kN)],[备注],[能否反洗井]) values (" _
                        & "'" & Text13.Text & "','" & Combo1.Text & "'," & CStr(Text2.Text) & "," & CStr(Text3.Text) & "," & CStr(Text8.Text) & "," _
                        & "'" & Text1.Text & "'," & CStr(Text5.Text) & ",'" & Text9.Text & "'," & CStr(Text12.Text) & "," & CStr(Text11.Text) & "," & "'" _
                        & Text15.Text & "'," & CStr(Text18.Text) & ",'" & Text16.Text & "'," & CStr(Text7.Text) & ",'" & Text17.Text & "'," _
                        & CStr(Text4.Text) & "," & CStr(Text20.Text) & "," & CStr(Text21.Text) & "," & CStr(Text19.Text) & "," & CStr(Text10.Text) & "," _
                        & CStr(Text14.Text) & "," & CStr(Text6.Text) & ",'" & Text22.Text & "','" & nffxj & " ')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "DELETE * from 封隔器信封曲线参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                        & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                For Each fgqxfqx_row In fgqxfqx_Table.Select
                    SQL_command = "insert into 封隔器信封曲线参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,坐封方式)  values (" _
                            & fgqxfqx_row.Item("序号").ToString & "," & fgqxfqx_row.Item("轴力kN").ToString & "," & fgqxfqx_row.Item("压差MPa").ToString & ",'" & Text13.Text & "'," & CStr(Text2.Text) & "," _
                            & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Combo1.Text & "')"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                Next
                basedb_chg = True
            End If
        Else
            msg_prompt = "是否要保存对数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "update 封隔器 set " _
                    & " [封隔器名称]='" & Text1.Text & "'," & " [重量kg]=" & CStr(Text5.Text) & "," & " [生产厂家]='" & Text9.Text & "'," & " [温度范围下℃]=" & CStr(Text12.Text) & "," _
                    & " [温度范围上℃]=" & CStr(Text11.Text) & "," & " [主体材料]='" & Text15.Text & "'," & " [主材屈服强度MPa]=" & CStr(Text18.Text) & "," & " [上端扣型]='" & Text16.Text & "'," _
                    & " [极限承压(MPa)]=" & CStr(Text7.Text) & "," & " [下端扣型]='" & Text17.Text & "'," & " [坐封力(kN)]=" & CStr(Text4.Text) & "," & " [坐封压差(MPa)]=" & CStr(Text20.Text) & "," _
                    & " [最小坐封压力MPa]=" & CStr(Text21.Text) & "," & " [最大坐封压力MPa]=" & CStr(Text19.Text) & "," & " [抗内压强度MPa]=" & CStr(Text10.Text) & "," _
                    & " [抗外压强度MPa]=" & CStr(Text14.Text) & "," & " [极限载荷(kN)]=" & CStr(Text6.Text) & "," & " [能否反洗井]='" & nffxj & "'," & " [备注]='" & Text22.Text & "'" _
                    & "  where [最大外径(mm)]=" & CStr(Text2.Text) & " and [最小通径(mm)]= " & CStr(Text3.Text) _
                    & " and [长度(m)]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_fgqgrid()
            Call fill_fgqxfqxgrid()
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
    Private Sub fgqqx_delete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fgqqx_delete.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 封隔器信封曲线参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            SQL_command = "DELETE * from 封隔器信封曲线参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                    & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            basedb_chg = True
        Else
            msg_prompt = "请选择好要删除的封隔器信封曲线参数！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_fgqxfqxgrid()
            Call CDMdraw_xfqx_Click(CDMdraw_xfqx, New System.EventArgs())
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的封隔器信封曲线参数！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"添加[&A]"按钮
    '*********************************************************************************************************************************************
    Private Sub fgqqx_save_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fgqqx_save.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        If IsNothing(Me.BindingSource1.Current) Or IsDBNull(Me.BindingSource1.Current("型号")) Then
            msg_prompt = "请选择封隔器。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Trim(Text13.Text) = "" Then
            msg_prompt = "封隔器型号是关键参数，请输入。"
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
        SQL_command = "select * from 封隔器 " & " where [最大外径(mm)]=" & CStr(Text2.Text) & " and [最小通径(mm)]= " & CStr(Text3.Text) _
                & " and [长度(m)]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            RECreader.Close()
            cn_basedb.Close()
            cn_basedb.Dispose()
            msg_prompt = "请选择封隔器或者先保存封隔器技术参数后选择该封隔器，然后添加信封曲线数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        SQL_command = "select * from 封隔器信封曲线参数表 " & " where [最大外径mm]=" & CStr(Text2.Text) & " and [最小通径mm]= " & CStr(Text3.Text) _
                & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " & " and  [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            '插入新记录
            SQL_command = "insert into 封隔器信封曲线参数表(序号,轴力kN,压差MPa,型号,最大外径mm,最小通径mm,长度m,坐封方式)  values (" _
                    & CStr(Text23.Text) & "," & CStr(Text24.Text) & "," & CStr(Text25.Text) & ",'" & Text13.Text & "'," & CStr(Text2.Text) & "," _
                    & CStr(Text3.Text) & "," & CStr(Text8.Text) & ",'" & Combo1.Text & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            basedb_chg = True
        Else
            '修改老记录
            SQL_command = "update 封隔器信封曲线参数表 set " _
                    & " [轴力kN]=" & CStr(Text24.Text) & "," & " [压差MPa]=" & CStr(Text25.Text) & " where [最大外径mm]=" & CStr(Text2.Text) _
                    & " and [最小通径mm]= " & CStr(Text3.Text) & " and [长度m]=" & CStr(Text8.Text) & " and  [坐封方式]='" & Combo1.Text & "' " _
                    & " and  [型号]='" & Text13.Text & "' and 序号=" & CStr(Text23.Text)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            basedb_chg = True
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_fgqxfqxgrid()
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
        Dim fgqxfqx_row As DataRow
        ReDim zl(fgqxfqx_Table.Rows.Count - 1)
        ReDim yc(fgqxfqx_Table.Rows.Count - 1)
        i = 0
        For Each fgqxfqx_row In fgqxfqx_Table.Select
            zl(i) = Val(fgqxfqx_row.Item("轴力kN").ToString)
            yc(i) = Val(fgqxfqx_row.Item("压差MPa").ToString)
            i = i + 1
        Next
        Call draw_xfqx(iPlotX1, zl, yc, fgqxfqx_Table.Rows.Count, "压差(MPa)", "轴力(kN)") '
    End Sub
End Class