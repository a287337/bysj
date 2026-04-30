Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class selt_zuoye
    '*********************************************************************************************************************************************
    '                                                   关于作业选择及管理界面的说明
    ' 参考资料
    '   VB怎么获取一个ACCESS数据库中的所有用户数据表的表名列表
    '   https://wenwen.sogou.com/z/q788487209.htm
    '   Data Type Mappings
    '   https://docs.microsoft.com/en-us/previous-versions/office/dd583373(v=office.11)?redirectedfrom=MSDN
    ' 程序升级记事：
    '                                                                                           秦彦斌 2019年10月22日最后整理
    '  （1）摸索掌握了VS中遍历库中所有表，遍历表中所有列的方法。
    '
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn As System.Data.OleDb.OleDbConnection
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub selt_zuoye_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
        zct_main.my_refresh()
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub selt_zuoye_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2 - VB6.PixelsToTwipsX(Me.Width) / 2), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2 - VB6.PixelsToTwipsY(Me.Height) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        Me.Text2.Text = ""
        If zuoye_name <> "" Then
            Text2.Text = zuoye_name
            Call textboxs_fuzhi()
        End If
        Call opt_change()
    End Sub
    '*********************************************************************************************************************************************
    'UPGRADE_WARNING: 初始化窗体时可能激发事件 Option1.CheckedChanged。 
    '*********************************************************************************************************************************************
    Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
        If eventSender.Checked Then
            Call opt_change()
        End If
    End Sub
    '*********************************************************************************************************************************************
    'UPGRADE_WARNING: 初始化窗体时可能激发事件 Option2.CheckedChanged。 
    '*********************************************************************************************************************************************
    Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
        If eventSender.Checked Then
            Call opt_change()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '双击ListBox1，给Text2.Text、zuoye_name赋值并退出
    '*********************************************************************************************************************************************
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.Items.Count > 0 Then
            Text2.Text = ListBox1.Text
            zuoye_name = Text2.Text
            '李润洲2025年10月22日增加
            If Not IsNothing(gz_yj_table) Then
                gz_yj_table.Clear()
            End If
            Me.Close()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '单击ListBox1，给Text2.Text赋值
    '*********************************************************************************************************************************************
    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.Items.Count > 0 Then
            Text2.Text = ListBox1.Text
            textboxs_fuzhi()
        End If
    End Sub
    '*********************************************************************************************************************************************
    '在“作业地层参数表”中找作业名称=Text2.Text的记录，给各输入框赋值
    '*********************************************************************************************************************************************
    Private Sub textboxs_fuzhi()
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select top 1 * from 作业地层参数表 where 井号='" & well_name & " ' and 作业名称='" & Text2.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            TextBox1.Text = RECreader.Item("压力系数").ToString
            TextBox2.Text = RECreader.Item("地破压力MPa").ToString
            TextBox3.Text = RECreader.Item("地破垂深m").ToString
            TextBox4.Text = RECreader.Item("目的层温度℃").ToString
            TextBox5.Text = RECreader.Item("目的层垂深m").ToString
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
        End If
        RECreader.Close()
        cn.Close()
    End Sub
    '*********************************************************************************************************************************************
    '将文本框中的数据保存到“作业地层参数表”中
    '*********************************************************************************************************************************************
    Private Function save_dccsh() As Boolean
        On Error GoTo ErrHandler
        save_dccsh = True
        '*********************************************************************************************************************************************
        '检查文本框数据合法性
        '*********************************************************************************************************************************************
        If Not IsNumeric(TextBox1.Text) Or IsDBNull(TextBox1.Text) Or Val(TextBox1.Text) = 0 Or Val(TextBox1.Text) < 0 Or Val(TextBox1.Text) > 3 Then
            TextBox1.Focus()
            msg_prompt = "请输入合法的压力系数值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            save_dccsh = False
            Exit Function
        End If
        If Not IsNumeric(TextBox2.Text) Or IsDBNull(TextBox2.Text) Or Val(TextBox2.Text) = 0 Or Val(TextBox2.Text) < 0 Or Val(TextBox2.Text) > 300 Then
            TextBox1.Focus()
            msg_prompt = "请输入合法的地层破裂压力值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            save_dccsh = False
            Exit Function
        End If
        If Not IsNumeric(TextBox3.Text) Or IsDBNull(TextBox3.Text) Or Val(TextBox3.Text) = 0 Or Val(TextBox3.Text) < 0 Or Val(TextBox3.Text) > 20000 Then
            TextBox1.Focus()
            msg_prompt = "请输入合法的地层破裂压力对应的垂深值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            save_dccsh = False
            Exit Function
        End If
        If Not IsNumeric(TextBox4.Text) Or IsDBNull(TextBox4.Text) Or Val(TextBox4.Text) = 0 Or Val(TextBox4.Text) < 0 Or Val(TextBox4.Text) > 300 Then
            TextBox1.Focus()
            msg_prompt = "请输入合法的目的层温度值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            save_dccsh = False
            Exit Function
        End If
        If Not IsNumeric(TextBox5.Text) Or IsDBNull(TextBox5.Text) Or Val(TextBox5.Text) = 0 Or Val(TextBox5.Text) < 0 Or Val(TextBox5.Text) > 20000 Then
            TextBox1.Focus()
            msg_prompt = "请输入合法的目的层温度对应的垂深值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            save_dccsh = False
            Exit Function
        End If
        '*********************************************************************************************************************************************
        '保存
        '*********************************************************************************************************************************************
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select * from 作业地层参数表 where 井号='" & well_name & " ' and 作业名称='" & Text2.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.Read Then
            '压力系数 FLOAT,地破压力MPa FLOAT,地破垂深m FLOAT,目的层温度℃ FLOAT,"                     & " 目的层垂深m FLOAT,井号
            SQL_command = "insert into 作业地层参数表(作业名称,压力系数,地破压力MPa,地破垂深m,目的层温度℃,目的层垂深m,井号) values (" _
                & "'" & Trim(Text2.Text) & "'," & CStr(TextBox1.Text) & "," & CStr(TextBox2.Text) & "," & CStr(TextBox3.Text) & "," & CStr(TextBox4.Text) & "," _
                & CStr(TextBox5.Text) & ",'" & well_name & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn.Close()
        Else
            SQL_command = "update 作业地层参数表 set " _
                        & " 压力系数=" & CStr(TextBox1.Text) & "," _
                        & " 地破压力MPa=" & CStr(TextBox2.Text) & "," _
                        & " 地破垂深m=" & CStr(TextBox3.Text) & "," _
                        & " 目的层温度℃=" & CStr(TextBox4.Text) & "," _
                        & " 目的层垂深m=" & CStr(TextBox5.Text) _
                        & " where 井号='" & well_name & " ' and 作业名称='" & Text2.Text & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
            RECreader.Close()
            cn.Close()
        End If
        save_dccsh = True
        Exit Function
ErrHandler:
        save_dccsh = False
        msg_prompt = "往[作业地层参数表]中写数据时出错。可能是数据文件在别处打开了，否则，请联系软件作者完善软件。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Function

    '*********************************************************************************************************************************************
    '待选名称来源选择radiobutton发生变化
    '*********************************************************************************************************************************************
    Private Sub opt_change()
        ListBox1.Items.Clear()
        If Option2.Checked = True Then
            cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn.Open()
            SQL_command = "select distinct 作业名称 from 管柱数据表 where 井号='" & well_name & " '"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            RECreader = EXECOleDbCommand.ExecuteReader()
            While RECreader.Read
                ListBox1.Items.Add(RECreader.Item("作业名称").ToString)
            End While
            RECreader.Close()
            cn.Close()
            CmdSaveAs.Enabled = True
            Cmddelete.Enabled = True
            Text1.Enabled = True
        End If
        If Option1.Checked = True Then
            cn = New System.Data.OleDb.OleDbConnection(AdoConString)
            cn.Open()
            SQL_command = "select distinct 作业名称 from 管柱作业名称"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            RECreader = EXECOleDbCommand.ExecuteReader()
            While RECreader.Read
                ListBox1.Items.Add(RECreader.Item("作业名称").ToString)
            End While
            RECreader.Close()
            cn.Close()
            CmdSaveAs.Enabled = False
            Cmddelete.Enabled = False
            Text1.Enabled = False
        End If
        ListBox1.Refresh()
    End Sub
    '*********************************************************************************************************************************************
    '点击"删除所选作业的所有数据[&D]"按钮
    '*********************************************************************************************************************************************
    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmddelete.Click
        Dim table_count As Integer
        table_count = 0
        If Text2.Text = "" Then
            msg_prompt = "请选择要删除的管柱作业名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        msg_prompt = "是否确定要删除‘" & Text2.Text & "’作业的所有数据？"
        msg_buttons = 4 + 32
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        If msg_return <> 6 Then
            Exit Sub
        End If
        '***********************************************************************************
        ' 2025年7月22日，李润洲修改为直接调用ipt_mdceng.vb中的delete_zy函数,避免在多个类中完成同一功能所带来的维护麻烦
        ' 传入待删除的作业名，返回删除涉及的表的个数
        ' Public Shared Function delete_zy(ByVal zyname As String) As Integer
        ' 20251112,秦彦斌将ipt_mdceng.vb中的delete_zy函数移至data_str.vb中
        ' Public Function delete_zy(ByVal zyname As String) As Integer
        '***********************************************************************************
        '***********************************************************************************
        ' 遍历库中所有表，删除 井号=well_name、作业名称=Text2.text 的所有记录
        '***********************************************************************************
        table_count = delete_zy(Trim(Text2.Text))
        If table_count > 0 Then
            msg_prompt = "共在" & table_count & "个表中删除了记录。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Call opt_change()
        Text2.Text = ""
    End Sub
    '*********************************************************************************************************************************************
    '点击"取消退出[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '点击"选定作业，保存地层参数，退出[&X]"按钮
    '*********************************************************************************************************************************************
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        If Text2.Text <> "" Then
            If save_dccsh() = True Then
                zuoye_name = Text2.Text
                '李润洲 2025年10月22日增加
                If Not IsNothing(gz_yj_table) Then
                    gz_yj_table.Clear()
                End If
                Me.Close()
            End If
        Else
            msg_prompt = "请选择或输入管柱作业名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存地层参数[&B]"按钮
    '*********************************************************************************************************************************************
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Text2.Text = "" Then
            msg_prompt = "请选择或输入管柱作业名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If save_dccsh() = True Then
            msg_prompt = "地层参数保存完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub

    '*********************************************************************************************************************************************
    '点击"将所选作业的所有数据另存给所输的作业名称[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub CmdSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaveAs.Click
        Dim dbSchema As DataTable
        Dim columnTable As DataTable
        Dim dbschema_row As DataRow
        Dim columnTab_row As DataRow
        Dim table_name As String
        Dim col_name As String
        Dim col_type As String
        Dim col_value As String
        Dim Insert_str As String
        Dim Values_str As String
        Dim i As Short
        Dim j As Short
        Dim hasZYField As Boolean
        Dim hasJHField As Boolean
        Dim oldZuoye_name As String
        Dim arrSussTableName() As String '已经插入了另存记录的表名字，用于回滚操作时删除已插入的记录
        Dim save_secsess As Boolean '有另存行为,另存成功
        Dim have_save As Boolean

        On Error GoTo ErrHandler
        oldZuoye_name = ""
        table_name = ""
        If CBool(Trim(CStr(Text2.Text = ""))) Then
            msg_prompt = "请选择要存储的作业名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        Else
            oldZuoye_name = Trim(Text2.Text)
        End If
        If Text1.Text = "" Then
            msg_prompt = "请输入有效的新作业名称。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '判断新的作业名称是否已有？
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "Select top 1 * from 管柱数据表 where 井号='" & well_name & "' And 作业名称='" & Trim(Text1.Text) & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            RECreader.Close()
            cn.Close()
            msg_prompt = "所输入新作业名称已存在。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        '*********************************************************************************************************************************************
        ' 遍历库中所有表，对存在“作业名称”和“井号”的表复制作业名称=Text2.text的记录，复制时将新记录的“作业名称”字段填为Text1.text
        '*********************************************************************************************************************************************
        save_secsess = False
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        dbSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        ReDim arrSussTableName(dbSchema.Rows.Count)
        i = 0
        For Each dbschema_row In dbSchema.Rows
            hasZYField = False
            hasJHField = False
            have_save = False
            table_name = dbschema_row.Item("TABLE_NAME").ToString
            columnTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "作业名称" Then
                    hasZYField = True
                ElseIf col_name = "井号" Then
                    hasJHField = True
                End If
            Next
            columnTable.Dispose()
            If hasZYField And hasJHField Then
                SQL_command = "Select * from " & table_name & " where 井号='" & well_name & "' And 作业名称='" & oldZuoye_name & "'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                RECreader = EXECOleDbCommand.ExecuteReader()
                While RECreader.Read()
                    have_save = True
                    Insert_str = "insert into " & table_name & "("
                    Values_str = " VALUES ("
                    For j = 0 To RECreader.FieldCount - 2
                        col_name = RECreader.GetName(j)
                        Insert_str = Insert_str & col_name & ","
                        If RECreader.GetName(j) = "作业名称" Then
                            '如果是("作业名称", 则修改为另存的名称)
                            Values_str = Values_str & "'" & Trim(Text1.Text) & "',"
                        Else
                            col_type = RECreader.GetDataTypeName(j)
                            col_value = Trim(RECreader.Item(j).ToString)
                            Select Case col_type
                                Case "DBTYPE_WVARCHAR", "DBTYPE_STR"
                                    '字符串
                                    Values_str = Values_str & "'" & col_value & "',"
                                Case "DBTYPE_R4", "DBTYPE_R8"
                                    '实型数字
                                    If col_value = "" Then col_value = "0.0"
                                    Values_str = Values_str & col_value & ","
                                Case "DBTYPE_I4", "DBTYPE_I8", "DBTYPE_I1", "DBTYPE_I2", "DBTYPE_UI1"
                                    '整型数字
                                    If col_value = "" Then col_value = "0"
                                    Values_str = Values_str & col_value & ","
                                Case "DBTYPE_FILETIME", "DBTYPE_DATE"
                                    '“日期”型，或“日期/时间”型
                                    If col_value = "" Then col_value = "2010-1-1"
                                    Values_str = Values_str & "#" & col_value & "#,"
                                Case Else
                                    msg_prompt = "另存作业时数据有遗漏，请联系软件作者完善软件。"
                                    msg_buttons = 0 + 48
                                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                                    Exit Sub
                            End Select
                        End If
                    Next j
                    '给Insert into子句加上结束的“）”，给Values子句加上结束的“）”。
                    col_name = RECreader.GetName(RECreader.FieldCount - 1)
                    Insert_str = Insert_str & col_name & ")"
                    If RECreader.GetName(RECreader.FieldCount - 1) = "作业名称" Then
                        '如果是("作业名称", 则修改为另存的名称)
                        Values_str = Values_str & "'" & Trim(Text1.Text) & "')"
                    Else
                        col_type = RECreader.GetDataTypeName(RECreader.FieldCount - 1)
                        col_value = Trim(RECreader.Item(RECreader.FieldCount - 1).ToString)
                        Select Case col_type
                            Case "DBTYPE_WVARCHAR", "DBTYPE_STR"
                                '字符串
                                Values_str = Values_str & "'" & col_value & "')"
                            Case "DBTYPE_R4", "DBTYPE_R8"
                                '实型数字
                                If col_value = "" Then col_value = "0.0"
                                Values_str = Values_str & col_value & ")"
                            Case "DBTYPE_I4", "DBTYPE_I8", "DBTYPE_I1", "DBTYPE_I2", "DBTYPE_UI1"
                                '整型数字
                                If col_value = "" Then col_value = "0"
                                Values_str = Values_str & col_value & ")"
                            Case "DBTYPE_FILETIME", "DBTYPE_DATE"
                                '“日期”型，或“日期/时间”型
                                If col_value = "" Then col_value = "0"
                                Values_str = Values_str & "#" & col_value & "#)"
                            Case Else
                                msg_prompt = "另存作业时数据有遗漏，请联系软件作者完善软件。"
                                msg_buttons = 0 + 48
                                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                                Exit Sub
                        End Select
                    End If
                    SQL_command = Insert_str & Values_str
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                    save_secsess = True
                End While
            End If 'hasZYField And hasJHField 判断结束
            If have_save Then
                '记录插入的表名字
                arrSussTableName(i) = table_name
                i = i + 1
                'msg_prompt = "在" & table_name & "表中另存数据成功。"
                'msg_buttons = 0 + 48
                'msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            End If
        Next
        dbSchema.Dispose()
        cn.Close()
        If save_secsess = True Then
            msg_prompt = "作业另存完成，共在" & i & "个表中插入记录。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Call opt_change()
        Else
            msg_prompt = "井数据库中没有‘" & Text2.Text & "’数据，无法另存。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "另存‘" & table_name & "’数据时出错，请联系软件作者完善软件，所有数据表将回滚到未插入前状态。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        '数据记录插入的回滚操作
        For j = 0 To i
            SQL_command = "delete * from " & arrSussTableName(j) & " where 井号='" & well_name & "' And 作业名称='" & Trim(Text1.Text) & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            EXECOleDbCommand.ExecuteNonQuery()
        Next j
        dbSchema.Dispose()
        cn.Close()
    End Sub
End Class
