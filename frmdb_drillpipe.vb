Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_drillpipe
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库钻杆数据输入、修改、管理窗体的说明
    '                                                                                           秦彦斌 2022年01月13日最后整理
    ' 说明：
    '   2018年3月18日增加字段
    '   (01)钢级              TEXT(50)
    '   (02)扣型              TEXT(50)
    '   (03)屈服强度MPa       FLOAT
    '   (04)弹性模量MPa       FLOAT                   钢制和铝合金的不同
    '   (05)泊松比            FLOAT                   钢制和铝合金的不同
    '   (06)管体抗扭强度Nm    FLOAT
    '   (07)接头抗扭强度Nm    FLOAT
    '   (08)管体抗拉强度kN    FLOAT
    '   (09)接头抗拉强度kN    FLOAT
    '   (10)抗内压强度MPa     FLOAT
    '   (11)抗挤强度MPa       FLOAT
    '   (12)接头外径mm        FLOAT
    ' 将“最小抗拉强度(kN)”值赋给“管体抗拉强度kN”后删除
    '
    ' 程序升级记事：
    '    20191215基础数据表维护模块升级到VS2008的第1个，积累经验。
    '    20220113重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private DP_Table As DataTable = b_dst.Tables.Add("DP_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_drillpipe_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_drillpipe_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text4.Text = CStr(0.0#)
        Text5.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text7.Text = ""
        Text8.Text = ""
        Text9.Text = CStr(0.0#)
        Text10.Text = CStr(0.0#)
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = CStr(0.0#)
        Text14.Text = CStr(0.0#)
        Text15.Text = CStr(0.0#)
        Text16.Text = CStr(0.0#)
        Text17.Text = CStr(206000.0#)
        Text18.Text = CStr(0.3)
        Text19.Text = CStr(0.0#)
        Text20.Text = CStr(0.0000124)
        Call fill_DPgrid()
    End Sub
    '*********************************************************************************************************************************************
    '钻杆数据列表DataGridView1设置并填充函数fill_DPgrid()，原来用 DataGrid1和Adodc4
    '*********************************************************************************************************************************************
    Private Sub fill_DPgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command =  "select * from 钻杆 order by [钻杆外径(mm)]"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        DP_Table.Clear()
        ad.Fill(DP_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = DP_Table
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
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中钻杆列表中的某行。原来用DataGrid1和Adodc4
    '用选中的钻杆列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("钻杆规格"))) Then
            Text1.Text = BindingSource1.Current("钻杆规格").ToString
            Text2.Text = IIf(BindingSource1.Current("钻杆外径(mm)").ToString = "", "0.0", BindingSource1.Current("钻杆外径(mm)").ToString)
            Text3.Text = IIf(BindingSource1.Current("钻杆壁厚(mm)").ToString = "", "0.0", BindingSource1.Current("钻杆壁厚(mm)").ToString)
            Text4.Text = IIf(BindingSource1.Current("管体抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("管体抗拉强度kN").ToString)
            Text5.Text = IIf(BindingSource1.Current("单根长度(m)").ToString = "", "0.0", BindingSource1.Current("单根长度(m)").ToString)
            Text6.Text = IIf(BindingSource1.Current("单位长度质量(Kg/m)").ToString = "", "0.0", BindingSource1.Current("单位长度质量(Kg/m)").ToString)
            Text7.Text = BindingSource1.Current("加厚型式").ToString
            Text8.Text = BindingSource1.Current("备注").ToString
            Text9.Text = IIf(BindingSource1.Current("接头外径mm").ToString = "", "0.0", BindingSource1.Current("接头外径mm").ToString)
            Text10.Text = IIf(BindingSource1.Current("屈服强度MPa").ToString = "", "0.0", BindingSource1.Current("屈服强度MPa").ToString)
            Text11.Text = BindingSource1.Current("扣型").ToString
            Text12.Text = BindingSource1.Current("钢级").ToString
            Text13.Text = IIf(BindingSource1.Current("管体抗扭强度Nm").ToString = "", "0.0", BindingSource1.Current("管体抗扭强度Nm").ToString)
            Text14.Text = IIf(BindingSource1.Current("接头抗扭强度Nm").ToString = "", "0.0", BindingSource1.Current("接头抗扭强度Nm").ToString)
            Text15.Text = IIf(BindingSource1.Current("抗内压强度MPa").ToString = "", "0.0", BindingSource1.Current("抗内压强度MPa").ToString)
            Text16.Text = IIf(BindingSource1.Current("抗挤强度MPa").ToString = "", "0.0", BindingSource1.Current("抗挤强度MPa").ToString)
            Text17.Text = IIf(BindingSource1.Current("弹性模量MPa").ToString = "", "206000.0", BindingSource1.Current("弹性模量MPa").ToString)
            Text18.Text = IIf(BindingSource1.Current("泊松比").ToString = "", "0.3", BindingSource1.Current("泊松比").ToString)
            Text19.Text = IIf(BindingSource1.Current("接头抗拉强度kN").ToString = "", "0.0", BindingSource1.Current("接头抗拉强度kN").ToString)
            Text20.Text = IIf(BindingSource1.Current("热膨胀系数").ToString = "", "0.0", BindingSource1.Current("热膨胀系数").ToString)
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
    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
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
        SQL_command = "select * from 钻杆 " & " where [钻杆规格]='" & Text1.Text & "' and [钻杆外径(mm)]= " & CStr(Text2.Text) _
            & " and [钻杆壁厚(mm)]=" & CStr(Text3.Text) & " and  [单根长度(m)]=" & CStr(Text5.Text) & " and  [单位长度质量(Kg/m)]=" & CStr(Text6.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的钻杆数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 钻杆 " & " where [钻杆规格]='" & Text1.Text & "' and [钻杆外径(mm)]= " & CStr(Text2.Text) _
                    & " and [钻杆壁厚(mm)]=" & CStr(Text3.Text) & " and  [单根长度(m)]=" & CStr(Text5.Text) & " and  [单位长度质量(Kg/m)]=" & CStr(Text6.Text)
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "钻杆数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的钻杆！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_DPgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的钻杆！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim basedb_chg As Boolean
        On Error GoTo ErrHandler
        If Text1.Text = "" Then
            msg_prompt = "请输入钻杆规格。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text2.Text) Or IsDBNull(Text2.Text) Or Val(Text2.Text) = 0 Then
            Text2.Focus()
            msg_prompt = "请输入合法的钻杆外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text3.Text) Or IsDBNull(Text3.Text) Or Val(Text3.Text) = 0 Then
            Text3.Focus()
            msg_prompt = "请输入合法的钻杆壁厚。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text5.Text) Or IsDBNull(Text5.Text) Or Val(Text5.Text) = 0 Then
            Text5.Focus()
            msg_prompt = "请输入合法的单根长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text9.Text) Or IsDBNull(Text9.Text) Or Val(Text9.Text) = 0 Then
            Text9.Focus()
            msg_prompt = "请输入合法的接头外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text7.Text = "" Then
            msg_prompt = "请输入加厚型式。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text6.Text) Or IsDBNull(Text6.Text) Or Val(Text6.Text) = 0 Then
            Text6.Focus()
            msg_prompt = "请输入合法的单位长度质量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text11.Text = "" Then
            msg_prompt = "请输入扣型。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Text12.Text = "" Then
            msg_prompt = "请输入钢级。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text10.Text) Or IsDBNull(Text10.Text) Or Val(Text10.Text) = 0 Then
            Text10.Focus()
            msg_prompt = "请输入合法的屈服强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text4.Text) Or IsDBNull(Text4.Text) Or Val(Text4.Text) = 0 Then
            Text4.Focus()
            msg_prompt = "请输入合法的管体抗拉强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text19.Text) Or IsDBNull(Text19.Text) Or Val(Text19.Text) = 0 Then
            Text19.Focus()
            msg_prompt = "请输入合法的接头抗拉强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text15.Text) Or IsDBNull(Text15.Text) Or Val(Text15.Text) = 0 Then
            Text15.Focus()
            msg_prompt = "请输入合法的抗内压强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text17.Text) Or IsDBNull(Text17.Text) Or Val(Text17.Text) = 0 Then
            Text17.Focus()
            msg_prompt = "请输入合法的弹性模量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text16.Text) Or IsDBNull(Text16.Text) Or Val(Text16.Text) = 0 Then
            Text16.Focus()
            msg_prompt = "请输入合法的抗挤强度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text18.Text) Or IsDBNull(Text18.Text) Or Val(Text18.Text) = 0 Then
            Text18.Focus()
            msg_prompt = "请输入合法的泊松比。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsNumeric(Text20.Text) Or IsDBNull(Text20.Text) Or Val(Text20.Text) = 0 Then
            Text20.Focus()
            msg_prompt = "请输入合法的热膨胀系数。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 钻杆 " _
            & " where [钻杆规格]='" & Text1.Text & "' and [钻杆外径(mm)]= " & CStr(Text2.Text) _
            & " and [钻杆壁厚(mm)]=" & CStr(Text3.Text) & " and  [单根长度(m)]=" & CStr(Text5.Text) & " and  [单位长度质量(Kg/m)]=" & CStr(Text6.Text)
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的钻杆数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into  钻杆 ([钻杆规格],[钻杆外径(mm)],[钻杆壁厚(mm)],[管体抗拉强度kN],[单根长度(m)]," _
                    & "[单位长度质量(Kg/m)],[加厚型式],[备注],[钢级],[扣型]," _
                    & "[屈服强度MPa],[弹性模量MPa],[泊松比],[管体抗扭强度Nm],[接头抗扭强度Nm]," _
                    & "[接头抗拉强度kN],[抗内压强度MPa],[抗挤强度MPa],[接头外径mm],[热膨胀系数]) values (" _
                    & "'" & Trim(Text1.Text) & "'," & Str(Val(Text2.Text)) & "," & Str(Val(Text3.Text)) & "," & Str(Val(Text4.Text)) & "," & Str(Val(Text5.Text)) & "," _
                    & Str(Val(Text6.Text)) & ",'" & Trim(Text7.Text) & "','" & Trim(Text8.Text) & "','" & Trim(Text12.Text) & " ','" & Trim(Text11.Text) & "'," _
                    & Str(Val(Text10.Text)) & "," & Str(Val(Text17.Text)) & "," & Str(Val(Text18.Text)) & "," & Str(Val(Text13.Text)) & "," & Str(Val(Text14.Text)) & "," _
                    & Str(Val(Text19.Text)) & "," & Str(Val(Text15.Text)) & "," & Str(Val(Text16.Text)) & "," & Str(Val(Text9.Text)) & "," & Str(Val(Text20.Text)) & ")"
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
                SQL_command = "update 钻杆 set " & " [钻杆外径(mm)]=" & Str(Val(Text2.Text)) & ",[钻杆壁厚(mm)]=" & Str(Val(Text3.Text)) & "," _
                    & " [管体抗拉强度kN]=" & Str(Val(Text4.Text)) & ",[单根长度(m)]=" & Str(Val(Text5.Text)) & "," _
                    & " [单位长度质量(Kg/m)]=" & Str(Val(Text6.Text)) & ",[加厚型式]='" & Text7.Text & "'," _
                    & " [备注]='" & Text8.Text & "',[钢级]='" & Text12.Text & "'," _
                    & " [扣型]='" & Text11.Text & "',[屈服强度MPa]=" & Str(Val(Text10.Text)) & "," _
                    & " [弹性模量MPa]=" & Str(Val(Text17.Text)) & ",[泊松比]=" & Str(Val(Text18.Text)) & "," _
                    & " [管体抗扭强度Nm]=" & Str(Val(Text13.Text)) & ",[接头抗扭强度Nm]=" & Str(Val(Text14.Text)) & "," _
                    & " [接头抗拉强度kN]=" & Str(Val(Text19.Text)) & ",[抗内压强度MPa]=" & Str(Val(Text15.Text)) & "," _
                    & " [抗挤强度MPa]=" & Str(Val(Text16.Text)) & ",[接头外径mm]=" & Str(Val(Text9.Text)) & "," _
                    & " [热膨胀系数]=" & Str(Val(Text20.Text)) _
                    & " where [钻杆规格]='" & Text1.Text & "' and [钻杆外径(mm)]= " & CStr(Text2.Text) _
                    & " and [钻杆壁厚(mm)]=" & CStr(Text3.Text) & " and  [单根长度(m)]=" & CStr(Text5.Text) & " and  [单位长度质量(Kg/m)]=" & CStr(Text6.Text)
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        If basedb_chg = True Then
            Call fill_DPgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class