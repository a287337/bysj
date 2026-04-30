'**********************************************************************************************************************************
' 20190717-从VB6升级到VS2008完成
' 20211111-将Splash_TSM和Splash_WCA合并，放到common files文件夹下共用
'**********************************************************************************************************************************
Option Strict Off
Option Explicit On
Friend Class frmSplash
    Inherits System.Windows.Forms.Form
    Private Sub frmSplash_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar) '单击键盘后，启动主界面
        Me.Close()
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Label1.Text = sofe_name_l1
        Me.Label2.Text = sofe_name_l2
        Me.lblCompany.Text = sofe_danwei_l1
        Me.Label3.Text = sofe_danwei_l2
        Me.lblWarning.Text = sofe_time & sofe_ver
        Timer1.Interval = 3000
        Timer1.Start()
    End Sub
    Private Sub imgLogo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgLogo.Click '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub Label1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label1.Click '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub lblCompany_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblCompany.Click '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub lblCompanyProduct_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)  '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub lblCopyright_Click() '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub lblVersion_Click() '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub lblWarning_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblWarning.Click '单击后，启动主界面
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick '定时3秒，启动主界面
        Me.Close()
    End Sub
End Class