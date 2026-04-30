Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_youguan
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库油管数据输入、修改、管理窗体的说明
    '                                                                                                        秦彦斌 2022年01月14日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200128,基础数据表维护模块升级到VS2008的第3个。
    '    20220114重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private youguan_Table As DataTable = b_dst.Tables.Add("youguan_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_youguan_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_youguan_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        Text2.Text = ""
        Text8.Text = CStr(0.0#)
        Text4.Text = CStr(0.0#)
        Text5.Text = CStr(0.0#)
        Text7.Text = CStr(206842.72#)
        Text6.Text = CStr(0.3)
        Text10.Text = CStr(0.0#)
        Text9.Text = CStr(0.0#)
        Text11.Text = CStr(0.0#)
        Text12.Text = CStr(0.0#)
        Text3.Text = ""
        Text13.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text15.Text = ""
        Text19.Text = CStr(0.0000124)
        Call fill_youguangrid()
    End Sub
    '*********************************************************************************************************************************************
    '油管数据列表DataGridView1设置并填充函数fill_youguangrid()。
    '*********************************************************************************************************************************************
    Private Sub fill_youguangrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select [油管规格],[油管外径mm],[油管壁厚mm],[材料],[扣型]," _
            & "[单位长度质量kg/m],[屈服应力MPa],[抗外挤强度MPa],[抗内压强度MPa],[抗拉强度kN]," _
            & "[接头抗内压强度MPa],[接头抗拉强度kN],[备注],[弹性模量MPa],[泊松比], " _
            & "[热膨胀系数] from 油管 order by [油管外径mm] desc,[油管壁厚mm] desc,[材料] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        youguan_Table.Clear()
        ad.Fill(youguan_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = youguan_Table
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = BindingSource1
        DataGridView1.ResetBindings()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.MultiSelect = False
        DataGridView1.RowHeadersWidth = 24
        DataGridView1.Columns(0).Width = 250
        DataGridView1.Columns(12).Width = 400
        DataGridView1.ReadOnly = True
        DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Refresh()
        DataGridView1.Show()
    End Sub
    '*********************************************************************************************************************************************
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中油管列表中的某行。
    '用选中的油管列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("油管规格"))) Then
            Text1.Text = BindingSource1.Current("油管规格").ToString
            Text2.Text = BindingSource1.Current("材料").ToString
            Text8.Text = IIf(BindingSource1.Current("油管外径mm").ToString = "", "0.0", BindingSource1.Current("油管外径mm").ToString)
            Text4.Text = IIf(BindingSource1.Current("油管壁厚mm").ToString = "", "0.0", BindingSource1.Current("油管壁厚mm").ToString)
            Text5.Text = IIf(BindingSource1.Current("单位长度质量kg/m").ToString = "", "0.0", BindingSource1.Current("单位长度质量kg/m").ToString)
            Text7.Text = IIf(BindingSource1.Current("弹性模量MPa").ToString = "", "0.0", BindingSource1.Current("弹性模量MPa").ToString)
            Text6.Text = IIf(BindingSource1.Current("泊松比").ToString = "", "0.0", BindingSource1.Current("泊松比").ToString)
            Text10.Text = IIf(BindingSource1.Current("屈服应力MPa").ToString = "", "0.0", BindingSource1.Current("屈服应力MPa").ToString)
            Text9.Text = IIf(BindingSource1.Current("抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("抗拉强度kN").ToString)
            Text11.Text = IIf(BindingSource1.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗内压强度MPa").ToString)
            Text12.Text = IIf(BindingSource1.Current("抗外挤强度MPa").ToString = "", "0.0", BindingSource1.Current("抗外挤强度MPa").ToString)
            Text3.Text = BindingSource1.Current("扣型").ToString
            Text13.Text = IIf(BindingSource1.Current("接头抗内压强度MPa").ToString = "", "0.0", BindingSource1.Current("接头抗内压强度MPa").ToString)
            Text14.Text = IIf(BindingSource1.Current("接头抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("接头抗拉强度kN").ToString)
            Text15.Text = BindingSource1.Current("备注").ToString
            Text19.Text = IIf(BindingSource1.Current("热膨胀系数").ToString = "", "0.0", BindingSource1.Current("热膨胀系数").ToString)
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
    '点击"计算[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入油管外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) = 0 Then
            msg_prompt = "请输入油管壁厚。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text10.Text) = 0 Then
            msg_prompt = "请输入油管屈服极限。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        Text18.Text = CStr(cal_kangneiya_qd(Val(Text10.Text), Val(Text8.Text), Val(Text4.Text)))
        Text17.Text = CStr(cal_kangji_qd(Val(Text10.Text), Val(Text8.Text), Val(Text4.Text), 0, 0))
        Text16.Text = CStr(cal_kangla_qd(Val(Text10.Text), Val(Text8.Text), Val(Text4.Text)))
        '管端开口厚壁管   TextBox1.Text = CStr(cal_kangneiya_qd1(Val(Text10.Text), Val(Text8.Text), Val(Text4.Text)))
    End Sub
    '*********************************************************************************************************************************************
    '点击"使用计算结果"按钮
    '*********************************************************************************************************************************************
    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Text9.Text = Text16.Text
        Text12.Text = Text17.Text
        Text11.Text = Text18.Text
        Text15.Text = CStr(Now) & "利用API公式计算"
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
        SQL_command = "select * from 油管 " & " where [材料]='" & Text2.Text & "' and [油管外径mm]= " & CStr(Text8.Text) & " and [油管壁厚mm]=" & CStr(Text4.Text) _
                        & " and  [单位长度质量kg/m]=" & CStr(Text5.Text) & " and 扣型='" & Text3.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的油管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 油管 " & " where [材料]='" & Text2.Text & "' and [油管外径mm]= " & CStr(Text8.Text) & " and [油管壁厚mm]=" & CStr(Text4.Text) _
                                & " and  [单位长度质量kg/m]=" & CStr(Text5.Text) & " and 扣型='" & Text3.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "油管数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的油管！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_youguangrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "关键数据有错，请选择好要删除的油管！"
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
            msg_prompt = "请输入油管规格。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text2.Text = "" Then
            msg_prompt = "请输入油管钢级。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text3.Text = "" Then
            msg_prompt = "请输入油管扣型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入油管外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) = 0 Then
            msg_prompt = "请输入油管壁厚。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text5.Text) = 0 Then
            msg_prompt = "请输入油管单位长度质量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text7.Text) = 0 Then
            msg_prompt = "请输入油管材料弹性模量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text6.Text) = 0 Then
            msg_prompt = "请输入油管材料泊松比。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text10.Text) = 0 Then
            msg_prompt = "请输入油管材料屈服极限。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text9.Text) = 0 Then
            msg_prompt = "请输入油管抗拉强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text11.Text) = 0 Then
            msg_prompt = "请输入油管抗内压强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text12.Text) = 0 Then
            msg_prompt = "请输入油管抗外挤强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text19.Text) = 0 Then
            msg_prompt = "请输入油管的热膨胀系数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 油管 " & " where [材料]='" & Text2.Text & "' and [油管外径mm]= " & CStr(Text8.Text) & " and [油管壁厚mm]=" & CStr(Text4.Text) _
                    & " and  [单位长度质量kg/m]=" & CStr(Text5.Text) & " and 扣型='" & Text3.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的油管数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into  油管 ([油管规格],[材料],[油管外径mm],[油管壁厚mm],[单位长度质量kg/m]," _
                            & "[弹性模量MPa],[泊松比],[屈服应力MPa],[抗拉强度kN],[抗内压强度MPa]," _
                            & "[抗外挤强度MPa],[扣型],[接头抗内压强度MPa],[接头抗拉强度kN],[备注]," _
                            & "[热膨胀系数]) values (" & "'" _
                            & Text1.Text & "','" & Text2.Text & "'," & CStr(Text8.Text) & "," & CStr(Text4.Text) & "," & CStr(Text5.Text) & "," _
                            & CStr(Text7.Text) & "," & CStr(Text6.Text) & "," & CStr(Text10.Text) & "," & CStr(Text9.Text) & "," & CStr(Text11.Text) & "," _
                            & CStr(Text12.Text) & ",'" & Text3.Text & "'," & CStr(Text13.Text) & "," & CStr(Text14.Text) & ",'" & Text15.Text & "'," _
                            & CStr(Text19.Text) & ")"
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
                SQL_command = "update 油管 set " _
                        & " [油管规格]='" & Text1.Text & "',[材料]='" & Text2.Text & "'," & " [单位长度质量kg/m]=" & CStr(Text5.Text) & "," & " [弹性模量MPa]=" & CStr(Text7.Text) _
                        & ",[泊松比]=" & CStr(Text6.Text) & "," & " [屈服应力MPa]=" & CStr(Text10.Text) & ",[抗拉强度kN]=" & CStr(Text9.Text) & "," & " [扣型]='" & Trim(Text3.Text) _
                        & "',[接头抗内压强度MPa]=" & CStr(Text13.Text) & "," & " [接头抗拉强度kN]=" & CStr(Text14.Text) & ",[备注]='" & Trim(Text15.Text) & "'," _
                        & " [抗内压强度MPa]=" & CStr(Text11.Text) & ",[抗外挤强度MPa]=" & CStr(Text12.Text) & ",[热膨胀系数]=" & CStr(Text19.Text) _
                        & " where [材料]='" & Text2.Text & "' and [油管外径mm]= " & CStr(Text8.Text) & " and [油管壁厚mm]=" & CStr(Text4.Text) _
                        & " and  [单位长度质量kg/m]=" & CStr(Text5.Text) & " and 扣型='" & Text3.Text & "'"
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
            Call fill_youguangrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class