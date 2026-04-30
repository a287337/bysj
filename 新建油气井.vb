'**********************************************************************************************************************************
' 20190717-从VB6升级到VS2008基本完成，尚有几个UPGRADE_???
'**********************************************************************************************************************************
Option Strict Off
Option Explicit On
Friend Class new_well
	Inherits System.Windows.Forms.Form
	
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		Me.Close()
	End Sub
	
	Private Sub new_well_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2 - VB6.PixelsToTwipsX(Me.Width) / 2), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2 - VB6.PixelsToTwipsY(Me.Height) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
	
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
        'UPGRADE_NOTE: new_well 已升级到 new_well_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
		Dim new_well_Renamed As String
		Dim new_path As String
        Dim cat As New ADOX.Catalog
		Dim i As Integer
        Dim j As Integer

        'On Error GoTo err_handle
        If Text2.Text = "" Then
            msg_prompt = "请输入新的井号。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            Exit Sub
        End If
		
		'下面为分离油井名称
		'UPGRADE_WARNING: 在 Visual Basic .NET 中不支持 CommonDialog CancelError 属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"”
        'CommonDialog1.CancelError = False
        'CommonDialog1Save.InitialDirectory = My.Application.Info.DirectoryPath & "data\"
		'UPGRADE_WARNING: Filter 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
		CommonDialog1Save.Filter = "Access(*.mdb)|*.mdb"
		CommonDialog1Save.FileName = Text2.Text
		CommonDialog1Save.ShowDialog()
		new_path = CommonDialog1Save.FileName '获得井名路径
        If new_path <> "" And new_path <> Text2.Text Then
            i = InStrRev(new_path, ".")
            j = InStrRev(new_path, "\")
            new_well_Renamed = Mid(new_path, j + 1, i - j - 1)
            'UPGRADE_WARNING: Dir 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
            If Not Dir(CommonDialog1Save.FileName) = "" Then
                msg_prompt = "所选文件夹下有同名文件存在，是否覆盖？"
                msg_buttons = 4 + 32
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                If msg_return <> 6 Then
                    Exit Sub
                End If
                Kill((new_path))
            End If

            'Set mydatabase = CreateDatabase(new_path, dbLangGeneral)
            cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & new_path & "")
            well_name = new_well_Renamed
            use_dbname = new_path
            cat = Nothing
            msg_prompt = "油气井数据建立完成，请在相应的功能模块中输入数据并进行计算。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
        Me.Close()
    End Sub
End Class