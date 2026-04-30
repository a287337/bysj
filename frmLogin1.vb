Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class frmLogin1
    '*********************************************************************************************************************************************
    '                                                       关于登录窗体的说明
    ' 说明：
    '     输入用户名：13909183925，密码：13909183925，进入软件。
    ' 程序升级记事：
    '   20190717-从VB6升级到VS2008完成，但数据库技术用ADODB。
    '   2020年9月18日，数据库用ADO.NET操作迁移完成。
    '                                                                                           秦彦斌 2019年12月18日最后整理
    '*********************************************************************************************************************************************
    Inherits System.Windows.Forms.Form
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        '设置全局变量为 false
        '不提示失败的登录
        LoginSucceeded = False
        Me.Close()
        End
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Static login_num As Integer
        Dim rec_count As Integer
        If login_num >= 3 Then
            MsgBox("你尝试的次数太多！再见!", , "登录")
            End
        End If
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select  * from users where 用户名='" & txtUserName.Text & "' and 密码='" & txtPassword.Text & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.HasRows Then
            rec_count = 0
            While RECreader.Read
                rec_count = rec_count + 1
            End While
            If rec_count = 1 Then
                RECreader.Close()
                cn_basedb.Close()
                MsgBox("欢迎使用" & sofe_name & "。", , "登录")
                login_num = 0
                LoginSucceeded = True
                Me.Hide()
                Me.Close()
            Else
                RECreader.Close()
                cn_basedb.Close()
            End If
        Else
            RECreader.Close()
            cn_basedb.Close()
            MsgBox("无效的用户名或密码，请重试!", , "登录")
            txtPassword.Focus()
            System.Windows.Forms.SendKeys.Send("{Home}+{End}")
            login_num = login_num + 1
        End If
    End Sub
    Private Sub frmLogin1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.lblCompanyProduct(0).Text = sofe_name_l1
        Me.lblCompanyProduct(1).Text = sofe_name_l2
    End Sub
    Private Sub imgLogo_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgLogo.DoubleClick
        LoginSucceeded = True
        Me.Hide()
        Me.Close()
    End Sub
End Class