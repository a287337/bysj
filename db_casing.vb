Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_casing
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库套管数据输入、修改、管理窗体的说明
    '                                                                                           秦彦斌 2022年01月14日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20191216基础数据表维护模块升级到VS2008的第2个，积累经验。
    '    20220114重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private casing_Table As DataTable = b_dst.Tables.Add("casing_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_casing_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_casing_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        Text3.Text = CStr(0.0#)
        Text2.Text = CStr(0.0#)
        Text8.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text4.Text = CStr(0.0#)
        Text5.Text = CStr(0.0#)
        Text7.Text = ""
        Text9.Text = ""
        Text10.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text12.Text = CStr(0.0#)
        Text13.Text = ""
        Text14.Text = CStr(0.0#)
        Text17.Text = CStr(0.0#)
        Text18.Text = CStr(0.0#)
        Call fill_casinggrid()
    End Sub
    '*********************************************************************************************************************************************
    '套管数据列表DataGridView1设置并填充函数fill_casinggrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_casinggrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [套管规格],[套管外径(mm)],[壁厚(mm)],[钢级],[扣型],[单位长质量(kg/m)],[屈服极限(MPa)],[抗挤强度(MPa)],[抗内压强度(MPa)],[接头抗内压强度(MPa)],[抗拉强度(kN)],[接头抗拉强度(kN)],[备注] from 套管 order by [套管外径(mm)] desc ,[壁厚(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        casing_Table.Clear()
        ad.Fill(casing_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = casing_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.Columns(0).Width = (DataGridView1.Width - 26) * 0.17
        DataGridView1.Columns(1).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(2).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(3).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(4).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(5).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(6).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(7).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(8).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(9).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(10).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(11).Width = (DataGridView1.Width - 26) * 0.06
        DataGridView1.Columns(12).Width = (DataGridView1.Width - 26) * 0.17
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中套管列表中的某行。
    '用选中的套管列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("套管规格"))) Then
            Text1.Text = BindingSource1.Current("套管规格").ToString
            Text2.Text = IIf(BindingSource1.Current("套管外径(mm)").ToString = "", "0.0", BindingSource1.Current("套管外径(mm)").ToString)
            Text3.Text = IIf(BindingSource1.Current("壁厚(mm)").ToString = "", "0.0", BindingSource1.Current("壁厚(mm)").ToString)
            Text4.Text = IIf(BindingSource1.Current("抗挤强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗挤强度(MPa)").ToString)
            Text5.Text = IIf(BindingSource1.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("抗拉强度(kN)").ToString)
            Text6.Text = IIf(BindingSource1.Current("屈服极限(MPa)").ToString = "", "0.0", BindingSource1.Current("屈服极限(MPa)").ToString)
            Text7.Text = BindingSource1.Current("备注").ToString
            Text8.Text = IIf(BindingSource1.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗内压强度(MPa)").ToString)
            Text9.Text = BindingSource1.Current("钢级").ToString
            Text13.Text = BindingSource1.Current("扣型").ToString
            Text14.Text = IIf(BindingSource1.Current("单位长质量(kg/m)").ToString = "", "0.0", BindingSource1.Current("单位长质量(kg/m)").ToString)
            Text17.Text = IIf(BindingSource1.Current("接头抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("接头抗内压强度(MPa)").ToString)
            Text18.Text = IIf(BindingSource1.Current("接头抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("接头抗拉强度(kN)").ToString)
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
    '点击"计算[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入套管外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入套管壁厚。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text6.Text) = 0 Then
            msg_prompt = "请输入套管屈服极限。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Text10.Text = CStr(cal_kangneiya_qd(Val(Text6.Text), Val(Text2.Text), Val(Text3.Text)))
        Text11.Text = CStr(cal_kangji_qd(Val(Text6.Text), Val(Text2.Text), Val(Text3.Text), 0, 0))
        Text12.Text = CStr(cal_kangla_qd(Val(Text6.Text), Val(Text2.Text), Val(Text3.Text)))
    End Sub
    '*********************************************************************************************************************************************
    '点击"使用计算结果"按钮
    '*********************************************************************************************************************************************
    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Text8.Text = Text10.Text
        Text4.Text = Text11.Text
        Text5.Text = Text12.Text
        Text7.Text = CStr(Now) & "利用API公式计算"
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 套管 where [钢级]='" & Text9.Text & "' and [扣型]='" & Text13.Text & "' and [套管外径(mm)]=" & CStr(Text2.Text) & "  and [壁厚(mm)]=" & CStr(Text3.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的套管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 套管 where [钢级]='" & Text9.Text & "' and  [扣型]='" & Text13.Text & "' and [套管外径(mm)]=" & CStr(Text2.Text) & "  and [壁厚(mm)]=" & CStr(Text3.Text)
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "套管数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的套管！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_casinggrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "关键数据有错，请选择好要删除的套管！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        On Error GoTo errhandler
        If Text1.Text = "" Then
            msg_prompt = "请输入套管规格。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入套管外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入套管壁厚。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) = 0 Then
            msg_prompt = "请输入套管抗挤强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text5.Text) = 0 Then
            msg_prompt = "请输入套管抗拉强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text6.Text) = 0 Then
            msg_prompt = "请输入套管屈服极限。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入套管抗内压强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text9.Text = "" Then
            msg_prompt = "请输入套管钢级。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text13.Text = "" Then
            msg_prompt = "请输入套管扣型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text14.Text = "" Then
            msg_prompt = "请输入套管线重。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 套管 where [钢级]='" & Text9.Text & "' and [扣型]='" & Text13.Text & "' and [套管外径(mm)]=" & CStr(Text2.Text) & "  and [壁厚(mm)]=" & CStr(Text3.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新套管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into  套管 ([套管规格],[套管外径(mm)],[壁厚(mm)],[抗拉强度(kN)],[抗挤强度(MPa)]," _
                    & "[抗内压强度(MPa)],[屈服极限(MPa)],[备注],[钢级],[扣型]," _
                    & "[单位长质量(kg/m)],[接头抗内压强度(MPa)],[接头抗拉强度(kN)]) values (" _
                    & "'" & Text1.Text & "'," & Str(Val(Text2.Text)) & "," & Str(Val(Text3.Text)) & "," & Str(Val(Text5.Text)) & "," & Str(Val(Text4.Text)) & "," _
                    & Str(Val(Text8.Text)) & "," & Str(Val(Text6.Text)) & ",'" & Text7.Text & "','" & Text9.Text & "','" & Text13.Text & "'," _
                    & CStr(Text14.Text) & "," & Str(Val(Text17.Text)) & "," & Str(Val(Text18.Text)) & ")"
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
                SQL_command = "update 套管 set " & " [抗拉强度(kN)]=" & Str(Val(Text5.Text)) & ",[屈服极限(MPa)]=" & Str(Val(Text6.Text)) & "," _
                    & " [抗挤强度(MPa)]=" & Str(Val(Text4.Text)) & ",[抗内压强度(MPa)]=" & Str(Val(Text8.Text)) & "," _
                    & " [备注]='" & Text7.Text & "',[扣型]='" & Text13.Text & "'," _
                    & " [单位长质量(kg/m)]=" & CStr(Text14.Text) & ",[接头抗内压强度(MPa)]=" & Str(Val(Text17.Text)) & "," _
                    & " [接头抗拉强度(kN)]=" & CStr(Text18.Text) & ",[套管规格]='" & Text1.Text & "' " _
                    & " where [钢级]='" & Text9.Text & "' and [扣型]='" & Text13.Text & "'and [套管外径(mm)]=" & CStr(Text2.Text) & " and [壁厚(mm)]=" & CStr(Text3.Text)
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
            Call fill_casinggrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class