Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_ssyj
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库伸缩管数据输入、修改、管理窗体的说明
    '                                                                                                       秦彦斌 2022年1月14日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200201，基础数据表维护模块升级到VS2008的第8个。
    '    20220114，重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private SHSG_Table As DataTable = b_dst.Tables.Add("SHSG_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_ssyj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_ssyj_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text11.Text = ""
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text1.Text = ""
        Text8.Text = CStr(0.0#)
        Text5.Text = CStr(0.0#)
        Text12.Text = ""
        Text16.Text = CStr(0.0#)
        Text18.Text = CStr(0.0#)
        Text19.Text = CStr(0.0#)
        Text13.Text = ""
        Text17.Text = CStr(0.0#)
        Text4.Text = CStr(0.0#)
        Text14.Text = ""
        Text10.Text = CStr(0.0#)
        Text9.Text = CStr(0.0#)
        Text15.Text = ""
        Text6.Text = CStr(0.0#)
        Text7.Text = ""
        Call fill_SHSGgrid()
    End Sub
    '*********************************************************************************************************************************************
    '伸缩管数据列表DataGridView1设置并填充函数fill_SHSGgrid()
    '*********************************************************************************************************************************************
    Private Sub fill_SHSGgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [元件名称],[零件号],[外径(mm)],[内径(mm)],[全缩短长度m]," _
                & "[伸缩行程(m)],[重量(kg)],[生产厂家],[温度范围下℃],[温度范围上℃]," _
                & "[压力等级MPa],[主体材料],[主材屈服强度MPa],[上端扣型],[下端扣型]," _
                & "[抗内压强度(MPa)],[抗外挤强度(MPa)],[抗拉强度kN],[备注] from 伸缩元件 order by [外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        SHSG_Table.Clear()
        ad.Fill(SHSG_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = SHSG_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.Columns(0).Width = 180
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中伸缩管列表中的某行。
    '用选中的伸缩管列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("零件号"))) Then
            Text11.Text = BindingSource1.Current("零件号").ToString
            Text2.Text = IIf(BindingSource1.Current("外径(mm)").ToString = "", "0.0", BindingSource1.Current("外径(mm)").ToString)
            Text3.Text = IIf(BindingSource1.Current("内径(mm)").ToString = "", "0.0", BindingSource1.Current("内径(mm)").ToString)
            Text1.Text = BindingSource1.Current("元件名称").ToString
            Text8.Text = IIf(BindingSource1.Current("全缩短长度m").ToString = "", "0.0", BindingSource1.Current("全缩短长度m").ToString)
            Text5.Text = IIf(BindingSource1.Current("伸缩行程(m)").ToString = "", "0.0", BindingSource1.Current("伸缩行程(m)").ToString)
            Text12.Text = BindingSource1.Current("生产厂家").ToString
            Text16.Text = IIf(BindingSource1.Current("压力等级MPa").ToString = "", "0.0", BindingSource1.Current("压力等级MPa").ToString)
            Text18.Text = IIf(BindingSource1.Current("温度范围下℃").ToString = "", "0.0", BindingSource1.Current("温度范围下℃").ToString)
            Text19.Text = IIf(BindingSource1.Current("温度范围上℃").ToString = "", "0.0", BindingSource1.Current("温度范围上℃").ToString)
            Text13.Text = BindingSource1.Current("主体材料").ToString
            Text17.Text = IIf(BindingSource1.Current("主材屈服强度MPa").ToString = "", "0.0", BindingSource1.Current("主材屈服强度MPa").ToString)
            Text4.Text = IIf(BindingSource1.Current("重量(kg)").ToString = "", "0.0", BindingSource1.Current("重量(kg)").ToString)
            Text14.Text = BindingSource1.Current("上端扣型").ToString
            Text10.Text = IIf(BindingSource1.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗内压强度(MPa)").ToString)
            Text9.Text = IIf(BindingSource1.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗外挤强度(MPa)").ToString)
            Text15.Text = BindingSource1.Current("下端扣型").ToString
            Text6.Text = IIf(BindingSource1.Current("抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("抗拉强度kN").ToString)
            Text7.Text = BindingSource1.Current("备注").ToString
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"退出[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击"帮助[&H]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 伸缩元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [全缩短长度m]=" & CStr(Text8.Text) _
                & " and  [伸缩行程(m)]=" & CStr(Text5.Text) & " and [零件号]='" & Text11.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的伸缩管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 伸缩元件 " _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [全缩短长度m]=" & CStr(Text8.Text) _
                        & " and  [伸缩行程(m)]=" & CStr(Text5.Text) & " and [零件号]='" & Text11.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "伸缩管数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的伸缩管！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_SHSGgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的伸缩管！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        If Text11.Text = "" Then
            msg_prompt = "请输入伸缩元件型号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) <= 0 Or (Not IsNumeric(Text2.Text)) Then
            msg_prompt = "请输入伸缩元件外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) <= 0 Or (Not IsNumeric(Text3.Text)) Then
            msg_prompt = "请输入伸缩元件内径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) <= 0 Or (Not IsNumeric(Text8.Text)) Then
            msg_prompt = "请输入伸缩元件全缩短长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text5.Text) <= 0 Or (Not IsNumeric(Text5.Text)) Then
            msg_prompt = "请输入伸缩元件的伸缩行程。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text1.Text = "" Then
            msg_prompt = "请输入伸缩元件规格名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) <= 0 Or (Not IsNumeric(Text4.Text)) Then
            msg_prompt = "请输入伸缩元件重量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 伸缩元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [全缩短长度m]=" & CStr(Text8.Text) _
                & " and  [伸缩行程(m)]=" & CStr(Text5.Text) & " and [零件号]='" & Text11.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的伸缩管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into 伸缩元件([元件名称],[外径(mm)],[内径(mm)], [重量(kg)], [全缩短长度m]," _
                        & "[伸缩行程(m)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[压力等级MPa],[抗拉强度kN]," _
                        & "[上端扣型],[下端扣型],[主体材料],[主材屈服强度MPa],[温度范围下℃]," _
                        & "[温度范围上℃],[零件号],[生产厂家],[备注]) values (" _
                        & "'" & Text1.Text & "'," & CStr(Text2.Text) & "," & CStr(Text3.Text) & "," & CStr(Text4.Text) & "," & CStr(Text8.Text) & "," _
                        & CStr(Text5.Text) & "," & CStr(Text9.Text) & "," & CStr(Text10.Text) & "," & CStr(Text16.Text) & "," & CStr(Text6.Text) & "," _
                        & "'" & Trim(Text14.Text) & "','" & Trim(Text15.Text) & "','" & Trim(Text13.Text) & "'," & CStr(Text17.Text) & "," & CStr(Text18.Text) & "," _
                        & CStr(Text19.Text) & ",'" & Trim(Text11.Text) & "','" & Trim(Text12.Text) & "','" & Trim(Text7.Text) & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
            End If
        Else
            msg_prompt = "是否要保存对数据的修改？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "update 伸缩元件 set [元件名称]='" & Text1.Text & "',[重量(kg)]=" & CStr(Text4.Text) & "," & " [抗外挤强度(MPa)]=" & CStr(Text9.Text) _
                        & ",[抗内压强度(MPa)]=" & CStr(Text10.Text) & "," & " [生产厂家]='" & Trim(Text12.Text) & "',[压力等级MPa]=" & CStr(Text16.Text) & "," _
                        & " [温度范围下℃]=" & CStr(Text18.Text) & ",[温度范围上℃]=" & CStr(Text19.Text) & "," & " [主体材料]='" & Trim(Text13.Text) _
                        & "',[主材屈服强度MPa]=" & CStr(Text17.Text) & "," & " [上端扣型]='" & Trim(Text14.Text) & "',[下端扣型]='" & Trim(Text14.Text) & "'," _
                        & " [抗拉强度kN]=" & CStr(Text6.Text) & ",[备注]='" & Trim(Text7.Text) & "'" _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [全缩短长度m]=" & CStr(Text8.Text) _
                        & " and  [伸缩行程(m)]=" & CStr(Text5.Text) & " and [零件号]='" & Text11.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        If basedb_chg = True Then
            Call fill_SHSGgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class