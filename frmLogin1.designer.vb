<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLogin1
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
	End Sub
	'Form 重写 Dispose，以清理组件列表。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows 窗体设计器所必需的
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents imgLogo As System.Windows.Forms.PictureBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtUserName As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents _lblCompanyProduct_0 As System.Windows.Forms.Label
	Public WithEvents _lblCompanyProduct_1 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
	Public WithEvents lblCompanyProduct As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.imgLogo = New System.Windows.Forms.PictureBox
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me._lblCompanyProduct_0 = New System.Windows.Forms.Label
        Me._lblCompanyProduct_1 = New System.Windows.Forms.Label
        Me._lblLabels_0 = New System.Windows.Forms.Label
        Me._lblLabels_1 = New System.Windows.Forms.Label
        Me.lblCompanyProduct = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Frame1.SuspendLayout()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCompanyProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.imgLogo)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(6, 4)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(161, 261)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        '
        'imgLogo
        '
        Me.imgLogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.imgLogo.Image = CType(resources.GetObject("imgLogo.Image"), System.Drawing.Image)
        Me.imgLogo.Location = New System.Drawing.Point(4, 10)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(151, 245)
        Me.imgLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLogo.TabIndex = 0
        Me.imgLogo.TabStop = False
        '
        'txtUserName
        '
        Me.txtUserName.AcceptsReturn = True
        Me.txtUserName.BackColor = System.Drawing.SystemColors.Window
        Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUserName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUserName.Location = New System.Drawing.Point(321, 127)
        Me.txtUserName.MaxLength = 0
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUserName.Size = New System.Drawing.Size(155, 21)
        Me.txtUserName.TabIndex = 1
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(204, 216)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(98, 38)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "登录[&L]"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(362, 216)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(98, 38)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "取消[&C]"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'txtPassword
        '
        Me.txtPassword.AcceptsReturn = True
        Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassword.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtPassword.Location = New System.Drawing.Point(321, 175)
        Me.txtPassword.MaxLength = 0
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPassword.Size = New System.Drawing.Size(155, 21)
        Me.txtPassword.TabIndex = 3
        '
        '_lblCompanyProduct_0
        '
        Me._lblCompanyProduct_0.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me._lblCompanyProduct_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblCompanyProduct_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCompanyProduct_0.Font = New System.Drawing.Font("楷体_GB2312", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._lblCompanyProduct_0.ForeColor = System.Drawing.Color.Blue
        Me.lblCompanyProduct.SetIndex(Me._lblCompanyProduct_0, CType(0, Short))
        Me._lblCompanyProduct_0.Location = New System.Drawing.Point(167, 38)
        Me._lblCompanyProduct_0.Name = "_lblCompanyProduct_0"
        Me._lblCompanyProduct_0.Size = New System.Drawing.Size(339, 28)
        Me._lblCompanyProduct_0.TabIndex = 8
        Me._lblCompanyProduct_0.Text = "井下管柱力学分析软件"
        Me._lblCompanyProduct_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_lblCompanyProduct_1
        '
        Me._lblCompanyProduct_1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me._lblCompanyProduct_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblCompanyProduct_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCompanyProduct_1.Font = New System.Drawing.Font("楷体_GB2312", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._lblCompanyProduct_1.ForeColor = System.Drawing.Color.Blue
        Me.lblCompanyProduct.SetIndex(Me._lblCompanyProduct_1, CType(1, Short))
        Me._lblCompanyProduct_1.Location = New System.Drawing.Point(167, 69)
        Me._lblCompanyProduct_1.Name = "_lblCompanyProduct_1"
        Me._lblCompanyProduct_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCompanyProduct_1.Size = New System.Drawing.Size(339, 28)
        Me._lblCompanyProduct_1.TabIndex = 7
        Me._lblCompanyProduct_1.Text = "井下管柱力学分析软件"
        Me._lblCompanyProduct_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.AutoSize = True
        Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(209, 132)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(77, 12)
        Me._lblLabels_0.TabIndex = 0
        Me._lblLabels_0.Text = "用户名称(&U):"
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.AutoSize = True
        Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(221, 180)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(53, 12)
        Me._lblLabels_1.TabIndex = 2
        Me._lblLabels_1.Text = "密码(&P):"
        '
        'frmLogin1
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(506, 270)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me._lblCompanyProduct_0)
        Me.Controls.Add(Me._lblCompanyProduct_1)
        Me.Controls.Add(Me._lblLabels_0)
        Me.Controls.Add(Me._lblLabels_1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(189, 232)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "登录"
        Me.Frame1.ResumeLayout(False)
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCompanyProduct, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class