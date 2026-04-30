'**********************************************************************************************************************************
' 20190717-从VB6升级到VS2008完成
' 20200116舍弃ADODC修改
' 2025年7月21日增加“完钻钻头尺寸mm FLOAT”字段，界面中为text2控件输入。设计界面改为ToolLayoutPanel控件  李润洲    
'**********************************************************************************************************************************
Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Friend Class well_manage
    Inherits System.Windows.Forms.Form
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        On Error GoTo ErrHandler
        If Not IsDate(Text6.Text) Then
            msg_prompt = "无效的开钻日期！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
        If Not IsDate(Text9.Text) Then
            msg_prompt = "无效的完钻日期！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If

        If well_name <> "" Then
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select * from 油气井表 where 井号='" & well_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            If Not RECreader.Read Then
                msg_prompt = "是否保存油气井数据？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return = 6 Then
                    '增加完钻钻头尺寸mm 李润洲
                    SQL_command = "insert into 油气井表(井号,井别,地理位置,构造位置,完钻层位,设计井深m,完钻井深m,开钻日期,完钻日期,完钻钻头尺寸mm) values( " _
                        & "'" & well_name & "'," & "'" & Text5.Text & "'," & "'" & Text3.Text & "'," & "'" & Text4.Text & "'," & "'" & Text1.Text & "'," & CStr(Text7.Text) & "," _
                        & CStr(Text8.Text) & "," & "'" & Text6.Text & "'," & "'" & Text9.Text & "'," & Text2.Text & ")"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Else
                msg_prompt = "是否更新油气井数据？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return = 6 Then
                    '增加完钻钻头尺寸mm 李润洲
                    SQL_command = "update 油气井表 set " _
                        & "井别='" & Text5.Text & "'," & "地理位置='" & Text3.Text & "'," & "构造位置='" & Text4.Text & "'," & "完钻层位='" & Text1.Text & "'," _
                        & "设计井深m=" & CStr(Text7.Text) & "," & "完钻井深m=" & Str(Val(Text8.Text)) & "," & "开钻日期='" & Text6.Text & "'," & "完钻日期='" & Text9.Text & "'" _
                        & ",完钻钻头尺寸mm=" & Text2.Text _
                        & " where " & "井号='" & well_name & "'"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            End If
            cn_userdb.Close()
            cn_userdb.Dispose()
        Else
            msg_prompt = "请选择打开油气井文件！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
            Exit Sub ' 退出程序，以避免进入错误处理程序。
ErrHandler:
            msg_prompt = "保存数据出错,请输入正确合理的数据！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Me.Close()
    End Sub
    '*********************************************************************************************************************************************
    '界面Closed
    '*********************************************************************************************************************************************
    Private Sub well_manage_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
    Private Sub well_manage_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Text = well_name & "井基本参数输入"
        Text1.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = CStr(0.0#)
        Text8.Text = CStr(0.0#)
        Text9.Text = ""
        '增加完钻钻头尺寸mm  李润洲
        Text2.Text = "0"
        'Text11.text = 0#
        'Text12.text = 0#
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            Text1.Text = RECreader.Item("完钻层位").ToString()
            Text3.Text = RECreader.Item("地理位置").ToString()
            Text4.Text = RECreader.Item("构造位置").ToString()
            Text5.Text = RECreader.Item("井别").ToString()
            Text6.Text = Format(RECreader.Item("开钻日期"), "D")
            Text7.Text = RECreader.Item("设计井深m").ToString()
            Text8.Text = RECreader.Item("完钻井深m").ToString()
            Text9.Text = Format(RECreader.Item("完钻日期"), "D")
            '增加完钻钻头尺寸mm  李润洲
            Text2.Text = RECreader.Item("完钻钻头尺寸mm").ToString()
            'Combo1.text =  RECreader.Item("井型").ToString()
            'Text11.text =  RECreader.Item("地压梯度MPapbm").ToString()
            'Text12.text =  RECreader.Item("地温梯度℃pbm").ToString()
        End If
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Sub
End Class