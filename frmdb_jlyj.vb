Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmdb_jlyj
    '*********************************************************************************************************************************************
    '                                                   关于基础数据库节流工具数据输入、修改、管理窗体的说明
    '                                                                                                            秦彦斌 2022年1月14日最后整理
    ' 说明：
    '
    ' 程序升级记事：
    '    20200201，基础数据表维护模块升级到VS2008的第7个。
    '    20220114，重新布置界面控件排列，实现可最大化（缩放）。
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private b_dst As New DataSet("base_dst")
    Private JLGJ_Table As DataTable = b_dst.Tables.Add("JLGJ_Table")
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub frmdb_jlyj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub frmdb_jlyj_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        Text2.Text = CStr(0.0#)
        Text3.Text = CStr(0.0#)
        Text8.Text = CStr(0.0#)
        Text4.Text = CStr(0.0#)
        Text5.Text = CStr(0.0#)
        Text7.Text = CStr(0.0#)
        Text6.Text = CStr(0.0#)
        Text9.Text = ""
        Text10.Text = CStr(0.0#)
        Text11.Text = ""
        Call fill_JLGJgrid()
    End Sub
    '*********************************************************************************************************************************************
    '节流工具数据列表DataGridView1设置并填充函数fill_JLGJgrid()
    '*********************************************************************************************************************************************
    Private Sub fill_JLGJgrid()
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 节流元件 order by [外径(mm)] desc"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        JLGJ_Table.Clear()
        ad.Fill(JLGJ_Table)
        ad.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
        BindingSource1.DataSource = JLGJ_Table
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
    '单击DataGridView1，在单元格的任何部分被单击时事件-选中节流工具列表中的某行。
    '用选中的节流工具列表中的数据填写界面中各文本框
    '*********************************************************************************************************************************************
    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If (Not IsNothing(Me.BindingSource1.Current)) And (Not IsDBNull(Me.BindingSource1.Current("名称"))) Then
            Text1.Text = BindingSource1.Current("名称").ToString
            Text2.Text = IIf(BindingSource1.Current("外径(mm)").ToString = "", "0.0", BindingSource1.Current("外径(mm)").ToString)
            Text3.Text = IIf(BindingSource1.Current("内径(mm)").ToString = "", "0.0", BindingSource1.Current("内径(mm)").ToString)
            Text8.Text = IIf(BindingSource1.Current("长度(m)").ToString = "", "0.0", BindingSource1.Current("长度(m)").ToString)
            Text4.Text = IIf(BindingSource1.Current("重量(kg)").ToString = "", "0.0", BindingSource1.Current("重量(kg)").ToString)
            Text5.Text = IIf(BindingSource1.Current("抗拉强度(kN)").ToString = "", "0.0", BindingSource1.Current("抗拉强度(kN)").ToString)
            Text6.Text = IIf(BindingSource1.Current("抗内压强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗内压强度(MPa)").ToString)
            Text7.Text = IIf(BindingSource1.Current("抗外挤强度(MPa)").ToString = "", "0.0", BindingSource1.Current("抗外挤强度(MPa)").ToString)
            Text9.Text = IIf(BindingSource1.Current("节流孔内径(mm)").ToString = "", "0.0", BindingSource1.Current("节流孔内径(mm)").ToString)
            Text10.Text = IIf(BindingSource1.Current("节流孔长度(mm)").ToString = "", "0.0", BindingSource1.Current("节流孔长度(mm)").ToString)
            Text11.Text = BindingSource1.Current("节流流向").ToString
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
        SQL_command = "select * from 节流元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [节流流向]='" & Text11.Text & "' and  [名称]='" & Text1.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            msg_prompt = "是否确定要删除所选的节流工具数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "DELETE * from 节流元件 " _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                        & " and  [节流流向]='" & Text11.Text & "' and  [名称]='" & Text1.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
                msg_prompt = "节流工具数据删除完成。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Else
            msg_prompt = "请选择好要删除的节流工具！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()
        If basedb_chg = True Then
            Call fill_JLGJgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "关键数据有错，请选择好要删除的节流工具！"
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
            msg_prompt = "请输入节流元件型号规格名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text2.Text) = 0 Then
            msg_prompt = "请输入节流元件外径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text3.Text) = 0 Then
            msg_prompt = "请输入节流元件内径。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text8.Text) = 0 Then
            msg_prompt = "请输入节流元件长度。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Val(Text4.Text) = 0 Then
            msg_prompt = "请输入节流元件重量。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not (Text11.Text = "上下流通" Or Text11.Text = "内外流通") Then
            msg_prompt = "请选择合适的开关流向。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        basedb_chg = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 节流元件 " _
                & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                & " and  [节流流向]='" & Text11.Text & "' and  [名称]='" & Text1.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If Not RECreader.HasRows Then
            msg_prompt = "是否要建立新的节流工具数据？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return = 6 Then
                SQL_command = "insert into 节流元件(" _
                        & "[名称],[外径(mm)],[内径(mm)], [重量(kg)], [长度(m)]," _
                        & "[抗拉强度(kN)],[抗外挤强度(MPa)],[抗内压强度(MPa)],[节流孔内径(mm)],[节流孔长度(mm)]," _
                        & "[节流流向]) values (" _
                        & "'" & Text1.Text & "'," & CStr(Text2.Text) & "," & CStr(Text3.Text) & "," & CStr(Text4.Text) & "," & CStr(Text8.Text) & "," _
                        & CStr(Text5.Text) & "," & CStr(Text7.Text) & "," & CStr(Text6.Text) & "," & CStr(Text9.Text) & "," & CStr(Text10.Text) & ",'" _
                        & Text11.Text & "')"
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
                SQL_command = "update 节流元件 set " _
                        & " [重量(kg)]=" & CStr(Text4.Text) & "," & " [抗拉强度(kN)]=" & CStr(Text5.Text) & "," & " [抗外挤强度(MPa)]=" & CStr(Text7.Text) & "," _
                        & " [抗内压强度(MPa)]=" & CStr(Text6.Text) & "," & " [节流孔内径(mm)]=" & CStr(Text9.Text) & "," & " [节流孔长度(mm)]=" & CStr(Text9.Text) _
                        & " where [外径(mm)]=" & CStr(Text2.Text) & " and [内径(mm)]= " & CStr(Text3.Text) & " and [长度(m)]=" & CStr(Text8.Text) _
                        & " and  [节流流向]='" & Text11.Text & "' and  [名称]='" & Text1.Text & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                basedb_chg = True
            End If
        End If
        RECreader.Close()
        cn_basedb.Close()
        If basedb_chg = True Then
            Call fill_JLGJgrid()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "保存数据出错,请输入正确合理的数据！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class