Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class save_as
    '************************************************************************************************************************************************
    '                     另存为界面说明
    ' 程序升级记事：
    '                                                                                                           秦彦斌 2020年9月20日最后整理更新
    '   20190717-从VB6升级到VS2008基本完成，尚有几个UPGRADE_???未处理，数据库技术用ADODB。
    '   2020年9月20日开始数据库用ADO.NET往VS2008迁移，2020年9月20日，数据库用ADO.NET操作迁移完成。
    '   UPGRADE_WARNING: Dir 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
    '
    '************************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    '*********************************************************************************************************************************************
    '点击"退出[&C]"按钮
    '*********************************************************************************************************************************************
    Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub save_as_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    '*********************************************************************************************************************************************
    '界面LOAD
    '*********************************************************************************************************************************************
    Private Sub save_as_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Text1.Text = well_name
    End Sub
    '*********************************************************************************************************************************************
    '点击"保存[&S]"按钮
    '*********************************************************************************************************************************************
    Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
        Dim new_well_Renamed As String
        Dim new_path As String
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim dbSchema As DataTable
        Dim foundRow As DataRow
        Dim EXECOleDbCommand As OleDbCommand
        Dim columnTable As DataTable
        Dim columnTab_row As DataRow
        Dim table_name As String
        Dim col_name As String
        Dim SQL_command As String
        Dim i As Integer
        Dim j As Integer

        'On Error GoTo ErrHandler
        If Text2.Text = "" Then
            msg_prompt = "请输入新的井号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        '下面为分离油井名称
        CommonDialog1Save.InitialDirectory = My.Application.Info.DirectoryPath & "data\"
        CommonDialog1Save.Filter = "Access(*.mdb)|*.mdb"
        CommonDialog1Save.FileName = Text2.Text
        CommonDialog1Save.ShowDialog()
        new_path = CommonDialog1Save.FileName '获得井名路径
        If new_path <> "" Then

            i = InStrRev(new_path, ".")
            j = InStrRev(new_path, "\")

            new_well_Renamed = Mid(new_path, j + 1, i - j - 1)
            If Not Dir(CommonDialog1Save.FileName) = "" Then
                msg_prompt = "所选文件夹下有同名文件存在，是否覆盖？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return <> 6 Then
                    Exit Sub
                End If
            End If
            FileCopy(use_dbname, new_path)
            '将老井数据中的井号换做新的井号
            '全局变量连接字符串在此赋值
            use_AdoConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & new_path & ";Persist Security Info=False"
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            For Each foundRow In dbSchema.Select
                table_name = foundRow.Item("TABLE_NAME").ToString
                columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
                For Each columnTab_row In columnTable.Rows
                    col_name = columnTab_row.Item("COLUMN_NAME").ToString
                    If col_name = "井号" Then
                        SQL_command = "update " & table_name & " set 井号='" & new_well_Renamed & "' where 井号='" & well_name & "'"
                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                        EXECOleDbCommand.ExecuteNonQuery()
                        EXECOleDbCommand.Dispose()
                    End If
                Next
            Next
            cn_userdb.Close()
            well_name = new_well_Renamed
            use_dbname = new_path
            msg_prompt = "油气井数据另存完成。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Me.Close()
        Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
        msg_prompt = "                     另存数据出错" & Chr(13) & Chr(10) & "(1)请检查数据文件是否已由其他软件打开，若打开，请关闭之。" & Chr(13) & Chr(10) & "(2)文件名中不得有系统保留字符如" - "、" / "、" * "、" = "、"" & chr(37) & ""等。" & Chr(13) & Chr(10) & "(3)请检查数据库，是否有非程序生成的库表，若有，请删除。" & Chr(13) & Chr(10) & "(4)井号数据文件中的表名中不得前述系统保留字符，若有请删除或改名。"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub
End Class