<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmdb_youguan
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
    Public WithEvents Text18 As System.Windows.Forms.TextBox
	Public WithEvents Text17 As System.Windows.Forms.TextBox
	Public WithEvents Text16 As System.Windows.Forms.TextBox
	Public WithEvents _Label_16 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _Label_15 As System.Windows.Forms.Label
    Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents cmdUpdate As System.Windows.Forms.Button
	Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Text15 As System.Windows.Forms.TextBox
	Public WithEvents Text14 As System.Windows.Forms.TextBox
	Public WithEvents Text13 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
    Public WithEvents Text12 As System.Windows.Forms.TextBox
	Public WithEvents Text11 As System.Windows.Forms.TextBox
	Public WithEvents Text10 As System.Windows.Forms.TextBox
	Public WithEvents Text9 As System.Windows.Forms.TextBox
    Public WithEvents Text7 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text8 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents _Label_14 As System.Windows.Forms.Label
	Public WithEvents _Label_13 As System.Windows.Forms.Label
	Public WithEvents _Label_12 As System.Windows.Forms.Label
	Public WithEvents _Label_11 As System.Windows.Forms.Label
	Public WithEvents _Label_2 As System.Windows.Forms.Label
	Public WithEvents _Label_3 As System.Windows.Forms.Label
	Public WithEvents _Label_4 As System.Windows.Forms.Label
	Public WithEvents _Label_10 As System.Windows.Forms.Label
	Public WithEvents _Label_9 As System.Windows.Forms.Label
	Public WithEvents _Label_8 As System.Windows.Forms.Label
	Public WithEvents _Label_7 As System.Windows.Forms.Label
	Public WithEvents _Label_6 As System.Windows.Forms.Label
	Public WithEvents _Label_5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents _Label_0 As System.Windows.Forms.Label
    Public WithEvents _Label_17 As System.Windows.Forms.Label
	Public WithEvents Frame As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents Label As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text18 = New System.Windows.Forms.TextBox
        Me.Text17 = New System.Windows.Forms.TextBox
        Me.Text16 = New System.Windows.Forms.TextBox
        Me._Label_16 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me._Label_15 = New System.Windows.Forms.Label
        Me.Command2 = New System.Windows.Forms.Button
        Me.Command3 = New System.Windows.Forms.Button
        Me.Command1 = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.Text19 = New System.Windows.Forms.TextBox
        Me.Text15 = New System.Windows.Forms.TextBox
        Me.Text14 = New System.Windows.Forms.TextBox
        Me.Text13 = New System.Windows.Forms.TextBox
        Me.Text3 = New System.Windows.Forms.TextBox
        Me.Text12 = New System.Windows.Forms.TextBox
        Me.Text11 = New System.Windows.Forms.TextBox
        Me.Text10 = New System.Windows.Forms.TextBox
        Me.Text9 = New System.Windows.Forms.TextBox
        Me.Text7 = New System.Windows.Forms.TextBox
        Me.Text6 = New System.Windows.Forms.TextBox
        Me.Text5 = New System.Windows.Forms.TextBox
        Me.Text4 = New System.Windows.Forms.TextBox
        Me.Text8 = New System.Windows.Forms.TextBox
        Me.Text2 = New System.Windows.Forms.TextBox
        Me.Text1 = New System.Windows.Forms.TextBox
        Me._Label_14 = New System.Windows.Forms.Label
        Me._Label_13 = New System.Windows.Forms.Label
        Me._Label_12 = New System.Windows.Forms.Label
        Me._Label_11 = New System.Windows.Forms.Label
        Me._Label_2 = New System.Windows.Forms.Label
        Me._Label_3 = New System.Windows.Forms.Label
        Me._Label_4 = New System.Windows.Forms.Label
        Me._Label_10 = New System.Windows.Forms.Label
        Me._Label_9 = New System.Windows.Forms.Label
        Me._Label_8 = New System.Windows.Forms.Label
        Me._Label_7 = New System.Windows.Forms.Label
        Me._Label_6 = New System.Windows.Forms.Label
        Me._Label_5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me._Label_0 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me._Label_17 = New System.Windows.Forms.Label
        Me.Frame = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel
        Me.Frame1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Text18
        '
        Me.Text18.AcceptsReturn = True
        Me.Text18.BackColor = System.Drawing.SystemColors.Window
        Me.Text18.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text18.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text18.Location = New System.Drawing.Point(431, 6)
        Me.Text18.MaxLength = 0
        Me.Text18.Name = "Text18"
        Me.Text18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text18.Size = New System.Drawing.Size(80, 23)
        Me.Text18.TabIndex = 53
        '
        'Text17
        '
        Me.Text17.AcceptsReturn = True
        Me.Text17.BackColor = System.Drawing.SystemColors.Window
        Me.Text17.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text17.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text17.Location = New System.Drawing.Point(138, 6)
        Me.Text17.MaxLength = 0
        Me.Text17.Name = "Text17"
        Me.Text17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text17.Size = New System.Drawing.Size(80, 23)
        Me.Text17.TabIndex = 52
        '
        'Text16
        '
        Me.Text16.AcceptsReturn = True
        Me.Text16.BackColor = System.Drawing.SystemColors.Window
        Me.Text16.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text16.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text16.Location = New System.Drawing.Point(652, 6)
        Me.Text16.MaxLength = 0
        Me.Text16.Name = "Text16"
        Me.Text16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text16.Size = New System.Drawing.Size(81, 23)
        Me.Text16.TabIndex = 51
        '
        '_Label_16
        '
        Me._Label_16.AutoSize = True
        Me._Label_16.BackColor = System.Drawing.SystemColors.Control
        Me._Label_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_16.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_16, CType(16, Short))
        Me._Label_16.Location = New System.Drawing.Point(520, 3)
        Me._Label_16.Name = "_Label_16"
        Me._Label_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_16.Size = New System.Drawing.Size(123, 31)
        Me._Label_16.TabIndex = 61
        Me._Label_16.Text = "抗拉强度(kN)"
        Me._Label_16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(6, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(123, 31)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "抗挤强度(MPa)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_15
        '
        Me._Label_15.AutoSize = True
        Me._Label_15.BackColor = System.Drawing.SystemColors.Control
        Me._Label_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_15.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_15, CType(15, Short))
        Me._Label_15.Location = New System.Drawing.Point(227, 3)
        Me._Label_15.Name = "_Label_15"
        Me._Label_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_15.Size = New System.Drawing.Size(195, 31)
        Me._Label_15.TabIndex = 59
        Me._Label_15.Text = "抗内压强度(MPa)巴洛公式"
        Me._Label_15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Command2
        '
        Me.Command2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(755, 678)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(111, 41)
        Me.Command2.TabIndex = 49
        Me.Command2.Text = "计算[&C]"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(875, 678)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(111, 41)
        Me.Command3.TabIndex = 48
        Me.Command3.Text = "使用计算结果"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command1
        '
        Me.Command1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(993, 616)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(111, 41)
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
        Me.cmdClose.Location = New System.Drawing.Point(993, 678)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(111, 41)
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
        Me.cmdUpdate.Location = New System.Drawing.Point(993, 554)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(111, 41)
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
        Me.cmdDelete.Location = New System.Drawing.Point(993, 492)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(111, 41)
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
        Me.Frame1.Location = New System.Drawing.Point(6, 9)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(1101, 477)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "油管参数信息"
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
        Me.DataGridView1.Location = New System.Drawing.Point(5, 23)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(1091, 447)
        Me.DataGridView1.TabIndex = 0
        '
        'Text19
        '
        Me.Text19.AcceptsReturn = True
        Me.Text19.BackColor = System.Drawing.SystemColors.Window
        Me.Text19.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text19.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text19.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text19.Location = New System.Drawing.Point(160, 134)
        Me.Text19.MaxLength = 0
        Me.Text19.Name = "Text19"
        Me.Text19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text19.Size = New System.Drawing.Size(79, 23)
        Me.Text19.TabIndex = 49
        Me.Text19.Text = "Text19"
        Me.Text19.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text15
        '
        Me.Text15.AcceptsReturn = True
        Me.Text15.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text15, 5)
        Me.Text15.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text15.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text15.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text15.Location = New System.Drawing.Point(402, 134)
        Me.Text15.MaxLength = 0
        Me.Text15.Name = "Text15"
        Me.Text15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text15.Size = New System.Drawing.Size(568, 23)
        Me.Text15.TabIndex = 47
        Me.Text15.Text = "Text15"
        '
        'Text14
        '
        Me.Text14.AcceptsReturn = True
        Me.Text14.BackColor = System.Drawing.SystemColors.Window
        Me.Text14.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text14.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text14.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text14.Location = New System.Drawing.Point(886, 102)
        Me.Text14.MaxLength = 0
        Me.Text14.Name = "Text14"
        Me.Text14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text14.Size = New System.Drawing.Size(84, 23)
        Me.Text14.TabIndex = 46
        Me.Text14.Text = "Text14"
        Me.Text14.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text13
        '
        Me.Text13.AcceptsReturn = True
        Me.Text13.BackColor = System.Drawing.SystemColors.Window
        Me.Text13.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text13.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text13.Location = New System.Drawing.Point(644, 102)
        Me.Text13.MaxLength = 0
        Me.Text13.Name = "Text13"
        Me.Text13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text13.Size = New System.Drawing.Size(79, 23)
        Me.Text13.TabIndex = 45
        Me.Text13.Text = "Text13"
        Me.Text13.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.Location = New System.Drawing.Point(644, 38)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(79, 23)
        Me.Text3.TabIndex = 44
        Me.Text3.Text = "Text3"
        '
        'Text12
        '
        Me.Text12.AcceptsReturn = True
        Me.Text12.BackColor = System.Drawing.SystemColors.Window
        Me.Text12.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text12.Location = New System.Drawing.Point(402, 70)
        Me.Text12.MaxLength = 0
        Me.Text12.Name = "Text12"
        Me.Text12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text12.Size = New System.Drawing.Size(79, 23)
        Me.Text12.TabIndex = 36
        Me.Text12.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text11
        '
        Me.Text11.AcceptsReturn = True
        Me.Text11.BackColor = System.Drawing.SystemColors.Window
        Me.Text11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text11.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text11.Location = New System.Drawing.Point(644, 70)
        Me.Text11.MaxLength = 0
        Me.Text11.Name = "Text11"
        Me.Text11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text11.Size = New System.Drawing.Size(79, 23)
        Me.Text11.TabIndex = 35
        Me.Text11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text10
        '
        Me.Text10.AcceptsReturn = True
        Me.Text10.BackColor = System.Drawing.SystemColors.Window
        Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text10.Location = New System.Drawing.Point(160, 70)
        Me.Text10.MaxLength = 0
        Me.Text10.Name = "Text10"
        Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text10.Size = New System.Drawing.Size(79, 23)
        Me.Text10.TabIndex = 34
        Me.Text10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text9
        '
        Me.Text9.AcceptsReturn = True
        Me.Text9.BackColor = System.Drawing.SystemColors.Window
        Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text9.Location = New System.Drawing.Point(886, 70)
        Me.Text9.MaxLength = 0
        Me.Text9.Name = "Text9"
        Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text9.Size = New System.Drawing.Size(84, 23)
        Me.Text9.TabIndex = 33
        Me.Text9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text7
        '
        Me.Text7.AcceptsReturn = True
        Me.Text7.BackColor = System.Drawing.SystemColors.Window
        Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text7.Location = New System.Drawing.Point(160, 102)
        Me.Text7.MaxLength = 0
        Me.Text7.Name = "Text7"
        Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text7.Size = New System.Drawing.Size(79, 23)
        Me.Text7.TabIndex = 16
        Me.Text7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text6
        '
        Me.Text6.AcceptsReturn = True
        Me.Text6.BackColor = System.Drawing.SystemColors.Window
        Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text6.Location = New System.Drawing.Point(402, 102)
        Me.Text6.MaxLength = 0
        Me.Text6.Name = "Text6"
        Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text6.Size = New System.Drawing.Size(79, 23)
        Me.Text6.TabIndex = 14
        Me.Text6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text5
        '
        Me.Text5.AcceptsReturn = True
        Me.Text5.BackColor = System.Drawing.SystemColors.Window
        Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text5.Location = New System.Drawing.Point(886, 38)
        Me.Text5.MaxLength = 0
        Me.Text5.Name = "Text5"
        Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text5.Size = New System.Drawing.Size(84, 23)
        Me.Text5.TabIndex = 12
        Me.Text5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(160, 38)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(79, 23)
        Me.Text4.TabIndex = 11
        Me.Text4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text8
        '
        Me.Text8.AcceptsReturn = True
        Me.Text8.BackColor = System.Drawing.SystemColors.Window
        Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text8.Location = New System.Drawing.Point(886, 6)
        Me.Text8.MaxLength = 0
        Me.Text8.Name = "Text8"
        Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text8.Size = New System.Drawing.Size(84, 23)
        Me.Text8.TabIndex = 9
        Me.Text8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(402, 38)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(79, 23)
        Me.Text2.TabIndex = 8
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text1, 5)
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(160, 6)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(563, 23)
        Me.Text1.TabIndex = 6
        '
        '_Label_14
        '
        Me._Label_14.AutoSize = True
        Me._Label_14.BackColor = System.Drawing.SystemColors.Control
        Me._Label_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_14.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_14, CType(14, Short))
        Me._Label_14.Location = New System.Drawing.Point(248, 131)
        Me._Label_14.Name = "_Label_14"
        Me._Label_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_14.Size = New System.Drawing.Size(145, 30)
        Me._Label_14.TabIndex = 40
        Me._Label_14.Text = "备注"
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
        Me._Label_13.Location = New System.Drawing.Point(732, 99)
        Me._Label_13.Name = "_Label_13"
        Me._Label_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_13.Size = New System.Drawing.Size(145, 29)
        Me._Label_13.TabIndex = 39
        Me._Label_13.Text = "接头抗拉强度(kN)"
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
        Me._Label_12.Location = New System.Drawing.Point(490, 99)
        Me._Label_12.Name = "_Label_12"
        Me._Label_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_12.Size = New System.Drawing.Size(145, 29)
        Me._Label_12.TabIndex = 38
        Me._Label_12.Text = "接头抗内压强度(MPa)"
        Me._Label_12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_11
        '
        Me._Label_11.AutoSize = True
        Me._Label_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_11.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_11.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_11, CType(11, Short))
        Me._Label_11.Location = New System.Drawing.Point(490, 35)
        Me._Label_11.Name = "_Label_11"
        Me._Label_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_11.Size = New System.Drawing.Size(145, 29)
        Me._Label_11.TabIndex = 37
        Me._Label_11.Text = "扣型"
        Me._Label_11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_2
        '
        Me._Label_2.AutoSize = True
        Me._Label_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_2.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_2, CType(2, Short))
        Me._Label_2.Location = New System.Drawing.Point(6, 3)
        Me._Label_2.Name = "_Label_2"
        Me._Label_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_2.Size = New System.Drawing.Size(145, 29)
        Me._Label_2.TabIndex = 31
        Me._Label_2.Text = "油管规格"
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
        Me._Label_3.Location = New System.Drawing.Point(732, 3)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(145, 29)
        Me._Label_3.TabIndex = 30
        Me._Label_3.Text = "油管外径(mm)"
        Me._Label_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_4
        '
        Me._Label_4.AutoSize = True
        Me._Label_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_4, CType(4, Short))
        Me._Label_4.Location = New System.Drawing.Point(6, 99)
        Me._Label_4.Name = "_Label_4"
        Me._Label_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_4.Size = New System.Drawing.Size(145, 29)
        Me._Label_4.TabIndex = 29
        Me._Label_4.Text = "弹性模量(MPa)"
        Me._Label_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_10
        '
        Me._Label_10.AutoSize = True
        Me._Label_10.BackColor = System.Drawing.SystemColors.Control
        Me._Label_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_10.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_10, CType(10, Short))
        Me._Label_10.Location = New System.Drawing.Point(6, 67)
        Me._Label_10.Name = "_Label_10"
        Me._Label_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_10.Size = New System.Drawing.Size(145, 29)
        Me._Label_10.TabIndex = 28
        Me._Label_10.Text = "屈服强度(MPa)"
        Me._Label_10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_9
        '
        Me._Label_9.AutoSize = True
        Me._Label_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_9.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_9, CType(9, Short))
        Me._Label_9.Location = New System.Drawing.Point(248, 67)
        Me._Label_9.Name = "_Label_9"
        Me._Label_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_9.Size = New System.Drawing.Size(145, 29)
        Me._Label_9.TabIndex = 27
        Me._Label_9.Text = "抗外挤强度(MPa)"
        Me._Label_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_8
        '
        Me._Label_8.AutoSize = True
        Me._Label_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_8.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_8, CType(8, Short))
        Me._Label_8.Location = New System.Drawing.Point(490, 67)
        Me._Label_8.Name = "_Label_8"
        Me._Label_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_8.Size = New System.Drawing.Size(145, 29)
        Me._Label_8.TabIndex = 26
        Me._Label_8.Text = "管体抗内压强度(MPa)"
        Me._Label_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_7
        '
        Me._Label_7.AutoSize = True
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_7, CType(7, Short))
        Me._Label_7.Location = New System.Drawing.Point(732, 67)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(145, 29)
        Me._Label_7.TabIndex = 25
        Me._Label_7.Text = "管体抗拉强度(kN)"
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
        Me._Label_6.Location = New System.Drawing.Point(248, 99)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(145, 29)
        Me._Label_6.TabIndex = 15
        Me._Label_6.Text = "泊松比"
        Me._Label_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_5
        '
        Me._Label_5.AutoSize = True
        Me._Label_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_5.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_5, CType(5, Short))
        Me._Label_5.Location = New System.Drawing.Point(732, 35)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(145, 29)
        Me._Label_5.TabIndex = 13
        Me._Label_5.Text = "单位长度质量(kg/m)"
        Me._Label_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(6, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(145, 29)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "油管壁厚(mm)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me._Label_0.Location = New System.Drawing.Point(248, 35)
        Me._Label_0.Name = "_Label_0"
        Me._Label_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_0.Size = New System.Drawing.Size(145, 29)
        Me._Label_0.TabIndex = 7
        Me._Label_0.Text = "钢级"
        Me._Label_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(6, 131)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(145, 30)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "热膨胀系数(∕℃)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_17
        '
        Me._Label_17.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me._Label_17.AutoSize = True
        Me._Label_17.BackColor = System.Drawing.SystemColors.Control
        Me._Label_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_17, CType(17, Short))
        Me._Label_17.Location = New System.Drawing.Point(10, 661)
        Me._Label_17.Name = "_Label_17"
        Me._Label_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_17.Size = New System.Drawing.Size(182, 14)
        Me._Label_17.TabIndex = 62
        Me._Label_17.Text = "用API公式计算油管管体强度"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 6
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text17, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_15, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text18, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_16, 4, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text16, 5, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(10, 682)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(739, 37)
        Me.TableLayoutPanel1.TabIndex = 63
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel2.ColumnCount = 8
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Text15, 3, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Text19, 1, 4)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text1, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text14, 7, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_3, 6, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text13, 5, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Text8, 7, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text3, 5, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Text9, 7, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text11, 5, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text12, 3, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label4, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_14, 2, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Text4, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Text10, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text6, 3, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Text7, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_0, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_13, 6, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Text2, 3, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_11, 4, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_5, 6, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_12, 4, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Text5, 7, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_10, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_9, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_6, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_8, 4, 2)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_7, 6, 2)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_4, 0, 3)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(6, 492)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 5
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(976, 164)
        Me.TableLayoutPanel2.TabIndex = 64
        '
        'frmdb_youguan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1111, 723)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._Label_17)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(73, 23)
        Me.MinimizeBox = False
        Me.Name = "frmdb_youguan"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "油管数据维护"
        Me.Frame1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support"
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Text19 As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
#End Region
End Class