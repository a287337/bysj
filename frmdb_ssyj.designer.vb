<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmdb_ssyj
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
	Private ADOBind_Adodc1 As VB6.MBindingCollection
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents cmdUpdate As System.Windows.Forms.Button
	Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Text19 As System.Windows.Forms.TextBox
	Public WithEvents Text18 As System.Windows.Forms.TextBox
	Public WithEvents Text17 As System.Windows.Forms.TextBox
	Public WithEvents Text16 As System.Windows.Forms.TextBox
	Public WithEvents Text15 As System.Windows.Forms.TextBox
	Public WithEvents Text14 As System.Windows.Forms.TextBox
	Public WithEvents Text13 As System.Windows.Forms.TextBox
	Public WithEvents Text12 As System.Windows.Forms.TextBox
	Public WithEvents Text11 As System.Windows.Forms.TextBox
    Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text10 As System.Windows.Forms.TextBox
	Public WithEvents Text9 As System.Windows.Forms.TextBox
    Public WithEvents Text7 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text8 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents _Label_4 As System.Windows.Forms.Label
	Public WithEvents _Label_15 As System.Windows.Forms.Label
    Public WithEvents _Label_1 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents _Label_5 As System.Windows.Forms.Label
	Public WithEvents _Label_17 As System.Windows.Forms.Label
	Public WithEvents _Label_16 As System.Windows.Forms.Label
	Public WithEvents _Label_14 As System.Windows.Forms.Label
	Public WithEvents _Label_13 As System.Windows.Forms.Label
	Public WithEvents _Label_12 As System.Windows.Forms.Label
	Public WithEvents _Label_11 As System.Windows.Forms.Label
	Public WithEvents _Label_9 As System.Windows.Forms.Label
	Public WithEvents _Label_8 As System.Windows.Forms.Label
	Public WithEvents _Label_2 As System.Windows.Forms.Label
	Public WithEvents _Label_3 As System.Windows.Forms.Label
	Public WithEvents _Label_7 As System.Windows.Forms.Label
	Public WithEvents _Label_6 As System.Windows.Forms.Label
	Public WithEvents _Label_0 As System.Windows.Forms.Label
    Public WithEvents Frame As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents Label As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command1 = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.Text19 = New System.Windows.Forms.TextBox
        Me.Text18 = New System.Windows.Forms.TextBox
        Me.Text17 = New System.Windows.Forms.TextBox
        Me.Text16 = New System.Windows.Forms.TextBox
        Me.Text15 = New System.Windows.Forms.TextBox
        Me.Text14 = New System.Windows.Forms.TextBox
        Me.Text13 = New System.Windows.Forms.TextBox
        Me.Text12 = New System.Windows.Forms.TextBox
        Me.Text11 = New System.Windows.Forms.TextBox
        Me.Text3 = New System.Windows.Forms.TextBox
        Me.Text10 = New System.Windows.Forms.TextBox
        Me.Text9 = New System.Windows.Forms.TextBox
        Me.Text7 = New System.Windows.Forms.TextBox
        Me.Text6 = New System.Windows.Forms.TextBox
        Me.Text5 = New System.Windows.Forms.TextBox
        Me.Text4 = New System.Windows.Forms.TextBox
        Me.Text8 = New System.Windows.Forms.TextBox
        Me.Text2 = New System.Windows.Forms.TextBox
        Me.Text1 = New System.Windows.Forms.TextBox
        Me._Label_4 = New System.Windows.Forms.Label
        Me._Label_15 = New System.Windows.Forms.Label
        Me._Label_1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me._Label_5 = New System.Windows.Forms.Label
        Me._Label_17 = New System.Windows.Forms.Label
        Me._Label_16 = New System.Windows.Forms.Label
        Me._Label_14 = New System.Windows.Forms.Label
        Me._Label_13 = New System.Windows.Forms.Label
        Me._Label_12 = New System.Windows.Forms.Label
        Me._Label_11 = New System.Windows.Forms.Label
        Me._Label_9 = New System.Windows.Forms.Label
        Me._Label_8 = New System.Windows.Forms.Label
        Me._Label_2 = New System.Windows.Forms.Label
        Me._Label_3 = New System.Windows.Forms.Label
        Me._Label_7 = New System.Windows.Forms.Label
        Me._Label_6 = New System.Windows.Forms.Label
        Me._Label_0 = New System.Windows.Forms.Label
        Me.Frame = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Frame1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command1
        '
        Me.Command1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(943, 588)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(126, 47)
        Me.Command1.TabIndex = 3
        Me.Command1.Text = "帮助[&H]"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(943, 645)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(126, 47)
        Me.cmdClose.TabIndex = 2
        Me.cmdClose.Text = "退出[&X]"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(943, 474)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(126, 47)
        Me.cmdUpdate.TabIndex = 1
        Me.cmdUpdate.Text = "保存[&S]"
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(943, 531)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(126, 47)
        Me.cmdDelete.TabIndex = 0
        Me.cmdDelete.Text = "删除(&D)"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.DataGridView1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(5, 9)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(1064, 459)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "伸缩元件参数信息"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Control
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(5, 26)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(1053, 426)
        Me.DataGridView1.TabIndex = 0
        '
        'Text19
        '
        Me.Text19.AcceptsReturn = True
        Me.Text19.BackColor = System.Drawing.SystemColors.Window
        Me.Text19.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text19.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text19.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text19.Location = New System.Drawing.Point(811, 102)
        Me.Text19.MaxLength = 0
        Me.Text19.Name = "Text19"
        Me.Text19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text19.Size = New System.Drawing.Size(115, 23)
        Me.Text19.TabIndex = 52
        '
        'Text18
        '
        Me.Text18.AcceptsReturn = True
        Me.Text18.BackColor = System.Drawing.SystemColors.Window
        Me.Text18.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text18.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text18.Location = New System.Drawing.Point(560, 102)
        Me.Text18.MaxLength = 0
        Me.Text18.Name = "Text18"
        Me.Text18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text18.Size = New System.Drawing.Size(112, 23)
        Me.Text18.TabIndex = 51
        '
        'Text17
        '
        Me.Text17.AcceptsReturn = True
        Me.Text17.BackColor = System.Drawing.SystemColors.Window
        Me.Text17.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text17.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text17.Location = New System.Drawing.Point(560, 134)
        Me.Text17.MaxLength = 0
        Me.Text17.Name = "Text17"
        Me.Text17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text17.Size = New System.Drawing.Size(112, 23)
        Me.Text17.TabIndex = 50
        '
        'Text16
        '
        Me.Text16.AcceptsReturn = True
        Me.Text16.BackColor = System.Drawing.SystemColors.Window
        Me.Text16.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text16.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text16.Location = New System.Drawing.Point(560, 70)
        Me.Text16.MaxLength = 0
        Me.Text16.Name = "Text16"
        Me.Text16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text16.Size = New System.Drawing.Size(112, 23)
        Me.Text16.TabIndex = 49
        '
        'Text15
        '
        Me.Text15.AcceptsReturn = True
        Me.Text15.BackColor = System.Drawing.SystemColors.Window
        Me.Text15.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text15.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text15.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text15.Location = New System.Drawing.Point(136, 166)
        Me.Text15.MaxLength = 0
        Me.Text15.Name = "Text15"
        Me.Text15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text15.Size = New System.Drawing.Size(276, 23)
        Me.Text15.TabIndex = 48
        '
        'Text14
        '
        Me.Text14.AcceptsReturn = True
        Me.Text14.BackColor = System.Drawing.SystemColors.Window
        Me.Text14.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text14.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text14.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text14.Location = New System.Drawing.Point(136, 134)
        Me.Text14.MaxLength = 0
        Me.Text14.Name = "Text14"
        Me.Text14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text14.Size = New System.Drawing.Size(276, 23)
        Me.Text14.TabIndex = 47
        '
        'Text13
        '
        Me.Text13.AcceptsReturn = True
        Me.Text13.BackColor = System.Drawing.SystemColors.Window
        Me.Text13.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text13.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text13.Location = New System.Drawing.Point(136, 102)
        Me.Text13.MaxLength = 0
        Me.Text13.Name = "Text13"
        Me.Text13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text13.Size = New System.Drawing.Size(276, 23)
        Me.Text13.TabIndex = 46
        '
        'Text12
        '
        Me.Text12.AcceptsReturn = True
        Me.Text12.BackColor = System.Drawing.SystemColors.Window
        Me.Text12.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text12.Location = New System.Drawing.Point(136, 70)
        Me.Text12.MaxLength = 0
        Me.Text12.Name = "Text12"
        Me.Text12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text12.Size = New System.Drawing.Size(276, 23)
        Me.Text12.TabIndex = 45
        '
        'Text11
        '
        Me.Text11.AcceptsReturn = True
        Me.Text11.BackColor = System.Drawing.SystemColors.Window
        Me.Text11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text11.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text11.Location = New System.Drawing.Point(136, 6)
        Me.Text11.MaxLength = 0
        Me.Text11.Name = "Text11"
        Me.Text11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text11.Size = New System.Drawing.Size(276, 23)
        Me.Text11.TabIndex = 44
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.Location = New System.Drawing.Point(811, 6)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(115, 23)
        Me.Text3.TabIndex = 28
        '
        'Text10
        '
        Me.Text10.AcceptsReturn = True
        Me.Text10.BackColor = System.Drawing.SystemColors.Window
        Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text10.Location = New System.Drawing.Point(560, 166)
        Me.Text10.MaxLength = 0
        Me.Text10.Name = "Text10"
        Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text10.Size = New System.Drawing.Size(112, 23)
        Me.Text10.TabIndex = 27
        '
        'Text9
        '
        Me.Text9.AcceptsReturn = True
        Me.Text9.BackColor = System.Drawing.SystemColors.Window
        Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text9.Location = New System.Drawing.Point(811, 134)
        Me.Text9.MaxLength = 0
        Me.Text9.Name = "Text9"
        Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text9.Size = New System.Drawing.Size(115, 23)
        Me.Text9.TabIndex = 26
        '
        'Text7
        '
        Me.Text7.AcceptsReturn = True
        Me.Text7.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel1.SetColumnSpan(Me.Text7, 5)
        Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text7.Location = New System.Drawing.Point(136, 198)
        Me.Text7.MaxLength = 0
        Me.Text7.Name = "Text7"
        Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text7.Size = New System.Drawing.Size(790, 23)
        Me.Text7.TabIndex = 14
        '
        'Text6
        '
        Me.Text6.AcceptsReturn = True
        Me.Text6.BackColor = System.Drawing.SystemColors.Window
        Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text6.Location = New System.Drawing.Point(811, 166)
        Me.Text6.MaxLength = 0
        Me.Text6.Name = "Text6"
        Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text6.Size = New System.Drawing.Size(115, 23)
        Me.Text6.TabIndex = 12
        '
        'Text5
        '
        Me.Text5.AcceptsReturn = True
        Me.Text5.BackColor = System.Drawing.SystemColors.Window
        Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text5.Location = New System.Drawing.Point(811, 38)
        Me.Text5.MaxLength = 0
        Me.Text5.Name = "Text5"
        Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text5.Size = New System.Drawing.Size(115, 23)
        Me.Text5.TabIndex = 11
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(811, 70)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(115, 23)
        Me.Text4.TabIndex = 10
        '
        'Text8
        '
        Me.Text8.AcceptsReturn = True
        Me.Text8.BackColor = System.Drawing.SystemColors.Window
        Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text8.Location = New System.Drawing.Point(560, 38)
        Me.Text8.MaxLength = 0
        Me.Text8.Name = "Text8"
        Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text8.Size = New System.Drawing.Size(112, 23)
        Me.Text8.TabIndex = 9
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(560, 6)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(112, 23)
        Me.Text2.TabIndex = 8
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(136, 38)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(276, 23)
        Me.Text1.TabIndex = 6
        '
        '_Label_4
        '
        Me._Label_4.AutoSize = True
        Me._Label_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_4, CType(4, Short))
        Me._Label_4.Location = New System.Drawing.Point(681, 131)
        Me._Label_4.Name = "_Label_4"
        Me._Label_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_4.Size = New System.Drawing.Size(121, 29)
        Me._Label_4.TabIndex = 55
        Me._Label_4.Text = "抗外挤强度(MPa)"
        Me._Label_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_15
        '
        Me._Label_15.AutoSize = True
        Me._Label_15.BackColor = System.Drawing.SystemColors.Control
        Me._Label_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_15.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_15, CType(15, Short))
        Me._Label_15.Location = New System.Drawing.Point(421, 99)
        Me._Label_15.Name = "_Label_15"
        Me._Label_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_15.Size = New System.Drawing.Size(130, 29)
        Me._Label_15.TabIndex = 54
        Me._Label_15.Text = "温度范围低(℃)"
        Me._Label_15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_1
        '
        Me._Label_1.AutoSize = True
        Me._Label_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_1.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_1, CType(1, Short))
        Me._Label_1.Location = New System.Drawing.Point(681, 3)
        Me._Label_1.Name = "_Label_1"
        Me._Label_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_1.Size = New System.Drawing.Size(121, 29)
        Me._Label_1.TabIndex = 43
        Me._Label_1.Text = "内径(mm)"
        Me._Label_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(681, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(121, 29)
        Me.Label4.TabIndex = 42
        Me.Label4.Text = "伸缩行程(m)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_5
        '
        Me._Label_5.AutoSize = True
        Me._Label_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_5, CType(5, Short))
        Me._Label_5.Location = New System.Drawing.Point(681, 67)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(121, 29)
        Me._Label_5.TabIndex = 41
        Me._Label_5.Text = "重量(kg)"
        Me._Label_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_17
        '
        Me._Label_17.AutoSize = True
        Me._Label_17.BackColor = System.Drawing.SystemColors.Control
        Me._Label_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_17.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_17, CType(17, Short))
        Me._Label_17.Location = New System.Drawing.Point(681, 163)
        Me._Label_17.Name = "_Label_17"
        Me._Label_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_17.Size = New System.Drawing.Size(121, 29)
        Me._Label_17.TabIndex = 37
        Me._Label_17.Text = "抗拉强度(kN)"
        Me._Label_17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_16
        '
        Me._Label_16.AutoSize = True
        Me._Label_16.BackColor = System.Drawing.SystemColors.Control
        Me._Label_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_16.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_16, CType(16, Short))
        Me._Label_16.Location = New System.Drawing.Point(421, 131)
        Me._Label_16.Name = "_Label_16"
        Me._Label_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_16.Size = New System.Drawing.Size(130, 29)
        Me._Label_16.TabIndex = 36
        Me._Label_16.Text = "主材屈服强度(MPa)"
        Me._Label_16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_14
        '
        Me._Label_14.AutoSize = True
        Me._Label_14.BackColor = System.Drawing.SystemColors.Control
        Me._Label_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_14.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_14, CType(14, Short))
        Me._Label_14.Location = New System.Drawing.Point(421, 67)
        Me._Label_14.Name = "_Label_14"
        Me._Label_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_14.Size = New System.Drawing.Size(130, 29)
        Me._Label_14.TabIndex = 35
        Me._Label_14.Text = "压力等级(MPa)"
        Me._Label_14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_13
        '
        Me._Label_13.AutoSize = True
        Me._Label_13.BackColor = System.Drawing.SystemColors.Control
        Me._Label_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_13.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_13, CType(13, Short))
        Me._Label_13.Location = New System.Drawing.Point(6, 67)
        Me._Label_13.Name = "_Label_13"
        Me._Label_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_13.Size = New System.Drawing.Size(121, 29)
        Me._Label_13.TabIndex = 34
        Me._Label_13.Text = "生产厂家"
        Me._Label_13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_12
        '
        Me._Label_12.AutoSize = True
        Me._Label_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_12.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_12, CType(12, Short))
        Me._Label_12.Location = New System.Drawing.Point(6, 99)
        Me._Label_12.Name = "_Label_12"
        Me._Label_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_12.Size = New System.Drawing.Size(121, 29)
        Me._Label_12.TabIndex = 33
        Me._Label_12.Text = "主体材料"
        Me._Label_12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_11
        '
        Me._Label_11.AutoSize = True
        Me._Label_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_11.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_11, CType(11, Short))
        Me._Label_11.Location = New System.Drawing.Point(6, 163)
        Me._Label_11.Name = "_Label_11"
        Me._Label_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_11.Size = New System.Drawing.Size(121, 29)
        Me._Label_11.TabIndex = 32
        Me._Label_11.Text = "下端扣型"
        Me._Label_11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_9
        '
        Me._Label_9.AutoSize = True
        Me._Label_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_9.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_9, CType(9, Short))
        Me._Label_9.Location = New System.Drawing.Point(6, 131)
        Me._Label_9.Name = "_Label_9"
        Me._Label_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_9.Size = New System.Drawing.Size(121, 29)
        Me._Label_9.TabIndex = 31
        Me._Label_9.Text = "上端扣型"
        Me._Label_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_8
        '
        Me._Label_8.AutoSize = True
        Me._Label_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_8.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_8.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_8, CType(8, Short))
        Me._Label_8.Location = New System.Drawing.Point(6, 3)
        Me._Label_8.Name = "_Label_8"
        Me._Label_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_8.Size = New System.Drawing.Size(121, 29)
        Me._Label_8.TabIndex = 30
        Me._Label_8.Text = "型号"
        Me._Label_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_2
        '
        Me._Label_2.AutoSize = True
        Me._Label_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_2.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_2, CType(2, Short))
        Me._Label_2.Location = New System.Drawing.Point(6, 35)
        Me._Label_2.Name = "_Label_2"
        Me._Label_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_2.Size = New System.Drawing.Size(121, 29)
        Me._Label_2.TabIndex = 24
        Me._Label_2.Text = "伸缩元件名称"
        Me._Label_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_3
        '
        Me._Label_3.AutoSize = True
        Me._Label_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_3.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_3.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_3, CType(3, Short))
        Me._Label_3.Location = New System.Drawing.Point(421, 35)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(130, 29)
        Me._Label_3.TabIndex = 23
        Me._Label_3.Text = "全缩短长度(m)"
        Me._Label_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_7
        '
        Me._Label_7.AutoSize = True
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_7, CType(7, Short))
        Me._Label_7.Location = New System.Drawing.Point(6, 195)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(121, 31)
        Me._Label_7.TabIndex = 22
        Me._Label_7.Text = "备注"
        Me._Label_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_6
        '
        Me._Label_6.AutoSize = True
        Me._Label_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_6.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_6, CType(6, Short))
        Me._Label_6.Location = New System.Drawing.Point(421, 163)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(130, 29)
        Me._Label_6.TabIndex = 13
        Me._Label_6.Text = "抗内压强度(MPa)"
        Me._Label_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_0
        '
        Me._Label_0.AutoSize = True
        Me._Label_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_0.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_0.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_0, CType(0, Short))
        Me._Label_0.Location = New System.Drawing.Point(421, 3)
        Me._Label_0.Name = "_Label_0"
        Me._Label_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_0.Size = New System.Drawing.Size(130, 29)
        Me._Label_0.TabIndex = 7
        Me._Label_0.Text = "外径(mm)"
        Me._Label_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(681, 99)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(121, 29)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "温度范围高(℃)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 6
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Text19, 5, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 4, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_8, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text3, 5, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text11, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text16, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text18, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_7, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Text17, 3, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Text9, 5, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Text7, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_11, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Text13, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_1, 4, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text5, 5, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text14, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Text4, 5, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 4, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text15, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Text10, 3, 5)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_9, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_5, 4, 2)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_12, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_6, 2, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Text6, 5, 5)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_16, 2, 4)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_4, 4, 4)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_15, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_14, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text2, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text8, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_13, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text12, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_17, 4, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Text1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_3, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_0, 2, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(5, 468)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 7
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(932, 229)
        Me.TableLayoutPanel1.TabIndex = 54
        '
        'frmdb_ssyj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1075, 703)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(73, 23)
        Me.MinimizeBox = False
        Me.Name = "frmdb_ssyj"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "伸缩元件数据维护"
        Me.Frame1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support"
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
#End Region 
End Class