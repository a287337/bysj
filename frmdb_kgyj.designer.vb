<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmdb_kgyj
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
	Private ADOBind_Adodc2 As VB6.MBindingCollection
	Private ADOBind_Adodc3 As VB6.MBindingCollection
    Public WithEvents _Frame_12 As System.Windows.Forms.GroupBox
    Public WithEvents Text18 As System.Windows.Forms.TextBox
    Public WithEvents Text17 As System.Windows.Forms.TextBox
	Public WithEvents Text16 As System.Windows.Forms.TextBox
	Public WithEvents Text15 As System.Windows.Forms.TextBox
	Public WithEvents Text14 As System.Windows.Forms.TextBox
	Public WithEvents Text13 As System.Windows.Forms.TextBox
    Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Text7 As System.Windows.Forms.TextBox
    Public WithEvents Text10 As System.Windows.Forms.TextBox
	Public WithEvents Text12 As System.Windows.Forms.TextBox
    Public WithEvents _Label_18 As System.Windows.Forms.Label
	Public WithEvents _Label_0 As System.Windows.Forms.Label
	Public WithEvents _Label_1 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents _Label_5 As System.Windows.Forms.Label
	Public WithEvents _Label_6 As System.Windows.Forms.Label
	Public WithEvents _Label_7 As System.Windows.Forms.Label
	Public WithEvents _Label_8 As System.Windows.Forms.Label
	Public WithEvents _Label_9 As System.Windows.Forms.Label
	Public WithEvents _Label_10 As System.Windows.Forms.Label
	Public WithEvents _Label_4 As System.Windows.Forms.Label
	Public WithEvents _Label_3 As System.Windows.Forms.Label
	Public WithEvents _Label_2 As System.Windows.Forms.Label
    Public WithEvents CDMdraw_xfqx As System.Windows.Forms.Button
	Public WithEvents zhxnqx_delete As System.Windows.Forms.Button
	Public WithEvents zhxnqx_save As System.Windows.Forms.Button
	Public WithEvents _Frame_23 As System.Windows.Forms.GroupBox
	Public WithEvents _Frame_24 As System.Windows.Forms.GroupBox
	Public WithEvents _Frame_25 As System.Windows.Forms.GroupBox
	Public WithEvents Text23 As System.Windows.Forms.TextBox
	Public WithEvents Text24 As System.Windows.Forms.TextBox
	Public WithEvents Text25 As System.Windows.Forms.TextBox
	Public WithEvents _Label_23 As System.Windows.Forms.Label
	Public WithEvents _Label_25 As System.Windows.Forms.Label
	Public WithEvents _Label_26 As System.Windows.Forms.Label
	Public WithEvents _Frame_22 As System.Windows.Forms.GroupBox
	Public WithEvents iPlotX1 As AxiPlotLibrary.AxiPlotX
	Public WithEvents _Frame_20 As System.Windows.Forms.GroupBox
    Public WithEvents _Frame_21 As System.Windows.Forms.GroupBox
    Public WithEvents Text9 As System.Windows.Forms.TextBox
	Public WithEvents Text8 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
    Public WithEvents Text11 As System.Windows.Forms.TextBox
    Public WithEvents _Label_17 As System.Windows.Forms.Label
	Public WithEvents _Label_16 As System.Windows.Forms.Label
	Public WithEvents _Label_15 As System.Windows.Forms.Label
	Public WithEvents _Label_14 As System.Windows.Forms.Label
	Public WithEvents _Label_13 As System.Windows.Forms.Label
    Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents cmdUpdate As System.Windows.Forms.Button
	Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents _Label_19 As System.Windows.Forms.Label
	Public WithEvents _Label_12 As System.Windows.Forms.Label
	Public WithEvents _Label_11 As System.Windows.Forms.Label
	Public WithEvents Frame As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents Label As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmdb_kgyj))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Frame_12 = New System.Windows.Forms.GroupBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Text18 = New System.Windows.Forms.TextBox
        Me.Text17 = New System.Windows.Forms.TextBox
        Me.Text16 = New System.Windows.Forms.TextBox
        Me.Text15 = New System.Windows.Forms.TextBox
        Me.Text14 = New System.Windows.Forms.TextBox
        Me.Text13 = New System.Windows.Forms.TextBox
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.Text4 = New System.Windows.Forms.TextBox
        Me.Text5 = New System.Windows.Forms.TextBox
        Me.Text6 = New System.Windows.Forms.TextBox
        Me.Text7 = New System.Windows.Forms.TextBox
        Me.Text10 = New System.Windows.Forms.TextBox
        Me.Text12 = New System.Windows.Forms.TextBox
        Me._Label_18 = New System.Windows.Forms.Label
        Me._Label_0 = New System.Windows.Forms.Label
        Me._Label_1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me._Label_5 = New System.Windows.Forms.Label
        Me._Label_6 = New System.Windows.Forms.Label
        Me._Label_7 = New System.Windows.Forms.Label
        Me._Label_8 = New System.Windows.Forms.Label
        Me._Label_9 = New System.Windows.Forms.Label
        Me._Label_10 = New System.Windows.Forms.Label
        Me._Label_4 = New System.Windows.Forms.Label
        Me._Label_3 = New System.Windows.Forms.Label
        Me._Label_2 = New System.Windows.Forms.Label
        Me.CDMdraw_xfqx = New System.Windows.Forms.Button
        Me.zhxnqx_delete = New System.Windows.Forms.Button
        Me.zhxnqx_save = New System.Windows.Forms.Button
        Me._Frame_22 = New System.Windows.Forms.GroupBox
        Me._Frame_23 = New System.Windows.Forms.GroupBox
        Me._Frame_25 = New System.Windows.Forms.GroupBox
        Me.Text23 = New System.Windows.Forms.TextBox
        Me.Text24 = New System.Windows.Forms.TextBox
        Me.Text25 = New System.Windows.Forms.TextBox
        Me._Label_23 = New System.Windows.Forms.Label
        Me._Label_25 = New System.Windows.Forms.Label
        Me._Label_26 = New System.Windows.Forms.Label
        Me._Frame_24 = New System.Windows.Forms.GroupBox
        Me._Frame_20 = New System.Windows.Forms.GroupBox
        Me.iPlotX1 = New AxiPlotLibrary.AxiPlotX
        Me._Frame_21 = New System.Windows.Forms.GroupBox
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.Text9 = New System.Windows.Forms.TextBox
        Me.Text8 = New System.Windows.Forms.TextBox
        Me.Text3 = New System.Windows.Forms.TextBox
        Me.Text2 = New System.Windows.Forms.TextBox
        Me.Text11 = New System.Windows.Forms.TextBox
        Me._Label_17 = New System.Windows.Forms.Label
        Me._Label_16 = New System.Windows.Forms.Label
        Me._Label_15 = New System.Windows.Forms.Label
        Me._Label_14 = New System.Windows.Forms.Label
        Me._Label_13 = New System.Windows.Forms.Label
        Me.Command1 = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me._Label_19 = New System.Windows.Forms.Label
        Me._Label_12 = New System.Windows.Forms.Label
        Me._Label_11 = New System.Windows.Forms.Label
        Me.Frame = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.BindingSource2 = New System.Windows.Forms.BindingSource(Me.components)
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel
        Me._Frame_12.SuspendLayout()
        Me._Frame_22.SuspendLayout()
        Me._Frame_20.SuspendLayout()
        CType(Me.iPlotX1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._Frame_21.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        '_Frame_12
        '
        Me._Frame_12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me._Frame_12.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_12.Controls.Add(Me.ListBox1)
        Me._Frame_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_12, CType(12, Short))
        Me._Frame_12.Location = New System.Drawing.Point(6, 393)
        Me._Frame_12.Name = "_Frame_12"
        Me._Frame_12.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_12.Size = New System.Drawing.Size(160, 133)
        Me._Frame_12.TabIndex = 82
        Me._Frame_12.TabStop = False
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Location = New System.Drawing.Point(5, 12)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(150, 116)
        Me.ListBox1.TabIndex = 0
        '
        'Text18
        '
        Me.Text18.AcceptsReturn = True
        Me.Text18.BackColor = System.Drawing.SystemColors.Window
        Me.Text18.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text18.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text18.Location = New System.Drawing.Point(350, 66)
        Me.Text18.MaxLength = 0
        Me.Text18.Name = "Text18"
        Me.Text18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text18.Size = New System.Drawing.Size(73, 23)
        Me.Text18.TabIndex = 80
        Me.Text18.Text = "Text18"
        '
        'Text17
        '
        Me.Text17.AcceptsReturn = True
        Me.Text17.BackColor = System.Drawing.SystemColors.Window
        Me.Text17.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text17.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text17.Location = New System.Drawing.Point(551, 36)
        Me.Text17.MaxLength = 0
        Me.Text17.Name = "Text17"
        Me.Text17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text17.Size = New System.Drawing.Size(78, 23)
        Me.Text17.TabIndex = 77
        Me.Text17.Text = "Text17"
        '
        'Text16
        '
        Me.Text16.AcceptsReturn = True
        Me.Text16.BackColor = System.Drawing.SystemColors.Window
        Me.Text16.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text16.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text16.Location = New System.Drawing.Point(137, 66)
        Me.Text16.MaxLength = 0
        Me.Text16.Name = "Text16"
        Me.Text16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text16.Size = New System.Drawing.Size(73, 23)
        Me.Text16.TabIndex = 76
        Me.Text16.Text = "Text16"
        '
        'Text15
        '
        Me.Text15.AcceptsReturn = True
        Me.Text15.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text15, 2)
        Me.Text15.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text15.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text15.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text15.Location = New System.Drawing.Point(432, 126)
        Me.Text15.MaxLength = 0
        Me.Text15.Name = "Text15"
        Me.Text15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text15.Size = New System.Drawing.Size(197, 23)
        Me.Text15.TabIndex = 75
        Me.Text15.Text = "Text15"
        '
        'Text14
        '
        Me.Text14.AcceptsReturn = True
        Me.Text14.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text14, 2)
        Me.Text14.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text14.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text14.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text14.Location = New System.Drawing.Point(137, 126)
        Me.Text14.MaxLength = 0
        Me.Text14.Name = "Text14"
        Me.Text14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text14.Size = New System.Drawing.Size(204, 23)
        Me.Text14.TabIndex = 74
        Me.Text14.Text = "Text14"
        '
        'Text13
        '
        Me.Text13.AcceptsReturn = True
        Me.Text13.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text13, 3)
        Me.Text13.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text13.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text13.Location = New System.Drawing.Point(137, 36)
        Me.Text13.MaxLength = 0
        Me.Text13.Name = "Text13"
        Me.Text13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text13.Size = New System.Drawing.Size(286, 23)
        Me.Text13.TabIndex = 73
        Me.Text13.Text = "Text13"
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text1, 3)
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(137, 6)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(286, 23)
        Me.Text1.TabIndex = 42
        Me.Text1.Text = "Text1"
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(551, 6)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(78, 23)
        Me.Text4.TabIndex = 41
        Me.Text4.Text = "Text4"
        '
        'Text5
        '
        Me.Text5.AcceptsReturn = True
        Me.Text5.BackColor = System.Drawing.SystemColors.Window
        Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text5.Location = New System.Drawing.Point(551, 96)
        Me.Text5.MaxLength = 0
        Me.Text5.Name = "Text5"
        Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text5.Size = New System.Drawing.Size(78, 23)
        Me.Text5.TabIndex = 40
        Me.Text5.Text = "Text5"
        '
        'Text6
        '
        Me.Text6.AcceptsReturn = True
        Me.Text6.BackColor = System.Drawing.SystemColors.Window
        Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text6.Location = New System.Drawing.Point(350, 96)
        Me.Text6.MaxLength = 0
        Me.Text6.Name = "Text6"
        Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text6.Size = New System.Drawing.Size(73, 23)
        Me.Text6.TabIndex = 39
        Me.Text6.Text = "Text6"
        '
        'Text7
        '
        Me.Text7.AcceptsReturn = True
        Me.Text7.BackColor = System.Drawing.SystemColors.Window
        Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text7.Location = New System.Drawing.Point(137, 96)
        Me.Text7.MaxLength = 0
        Me.Text7.Name = "Text7"
        Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text7.Size = New System.Drawing.Size(73, 23)
        Me.Text7.TabIndex = 38
        Me.Text7.Text = "Text7"
        '
        'Text10
        '
        Me.Text10.AcceptsReturn = True
        Me.Text10.BackColor = System.Drawing.SystemColors.Window
        Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text10.Location = New System.Drawing.Point(551, 66)
        Me.Text10.MaxLength = 0
        Me.Text10.Name = "Text10"
        Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text10.Size = New System.Drawing.Size(78, 23)
        Me.Text10.TabIndex = 29
        Me.Text10.Text = "Text10"
        '
        'Text12
        '
        Me.Text12.AcceptsReturn = True
        Me.Text12.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel2.SetColumnSpan(Me.Text12, 5)
        Me.Text12.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text12.Location = New System.Drawing.Point(137, 156)
        Me.Text12.MaxLength = 0
        Me.Text12.Name = "Text12"
        Me.Text12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text12.Size = New System.Drawing.Size(492, 23)
        Me.Text12.TabIndex = 28
        Me.Text12.Text = "Text12"
        '
        '_Label_18
        '
        Me._Label_18.AutoSize = True
        Me._Label_18.BackColor = System.Drawing.SystemColors.Control
        Me._Label_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_18.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_18, CType(18, Short))
        Me._Label_18.Location = New System.Drawing.Point(219, 63)
        Me._Label_18.Name = "_Label_18"
        Me._Label_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_18.Size = New System.Drawing.Size(122, 27)
        Me._Label_18.TabIndex = 79
        Me._Label_18.Text = "温度范围高(℃)"
        Me._Label_18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_0
        '
        Me._Label_0.AutoSize = True
        Me._Label_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_0.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_0, CType(0, Short))
        Me._Label_0.Location = New System.Drawing.Point(6, 123)
        Me._Label_0.Name = "_Label_0"
        Me._Label_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_0.Size = New System.Drawing.Size(122, 27)
        Me._Label_0.TabIndex = 54
        Me._Label_0.Text = "上端扣型"
        Me._Label_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_1
        '
        Me._Label_1.AutoSize = True
        Me._Label_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_1.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_1, CType(1, Short))
        Me._Label_1.Location = New System.Drawing.Point(432, 33)
        Me._Label_1.Name = "_Label_1"
        Me._Label_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_1.Size = New System.Drawing.Size(110, 27)
        Me._Label_1.TabIndex = 53
        Me._Label_1.Text = "压力等级(MPa)"
        Me._Label_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(432, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(110, 27)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "抗拉强度(kN)"
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
        Me._Label_5.Location = New System.Drawing.Point(432, 3)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(110, 27)
        Me._Label_5.TabIndex = 51
        Me._Label_5.Text = "质量(kg)"
        Me._Label_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_6
        '
        Me._Label_6.AutoSize = True
        Me._Label_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_6.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_6, CType(6, Short))
        Me._Label_6.Location = New System.Drawing.Point(219, 93)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(122, 27)
        Me._Label_6.TabIndex = 50
        Me._Label_6.Text = "抗内压强度(MPa)"
        Me._Label_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_7
        '
        Me._Label_7.AutoSize = True
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_7, CType(7, Short))
        Me._Label_7.Location = New System.Drawing.Point(350, 123)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(73, 27)
        Me._Label_7.TabIndex = 49
        Me._Label_7.Text = "下端扣型"
        Me._Label_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_8
        '
        Me._Label_8.AutoSize = True
        Me._Label_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_8.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_8, CType(8, Short))
        Me._Label_8.Location = New System.Drawing.Point(6, 63)
        Me._Label_8.Name = "_Label_8"
        Me._Label_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_8.Size = New System.Drawing.Size(122, 27)
        Me._Label_8.TabIndex = 48
        Me._Label_8.Text = "温度范围低(℃)"
        Me._Label_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_9
        '
        Me._Label_9.AutoSize = True
        Me._Label_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_9.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_9, CType(9, Short))
        Me._Label_9.Location = New System.Drawing.Point(6, 153)
        Me._Label_9.Name = "_Label_9"
        Me._Label_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_9.Size = New System.Drawing.Size(122, 32)
        Me._Label_9.TabIndex = 47
        Me._Label_9.Text = "备注"
        Me._Label_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_10
        '
        Me._Label_10.AutoSize = True
        Me._Label_10.BackColor = System.Drawing.SystemColors.Control
        Me._Label_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_10.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_10, CType(10, Short))
        Me._Label_10.Location = New System.Drawing.Point(432, 63)
        Me._Label_10.Name = "_Label_10"
        Me._Label_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_10.Size = New System.Drawing.Size(110, 27)
        Me._Label_10.TabIndex = 46
        Me._Label_10.Text = "极限压差(MPa)"
        Me._Label_10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_4
        '
        Me._Label_4.AutoSize = True
        Me._Label_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_4, CType(4, Short))
        Me._Label_4.Location = New System.Drawing.Point(6, 93)
        Me._Label_4.Name = "_Label_4"
        Me._Label_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_4.Size = New System.Drawing.Size(122, 27)
        Me._Label_4.TabIndex = 45
        Me._Label_4.Text = "抗外挤强度(MPa)"
        Me._Label_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_3
        '
        Me._Label_3.AutoSize = True
        Me._Label_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_3.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_3, CType(3, Short))
        Me._Label_3.Location = New System.Drawing.Point(6, 33)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(122, 27)
        Me._Label_3.TabIndex = 44
        Me._Label_3.Text = "生产厂家"
        Me._Label_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me._Label_2.Size = New System.Drawing.Size(122, 27)
        Me._Label_2.TabIndex = 43
        Me._Label_2.Text = "工具名称"
        Me._Label_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CDMdraw_xfqx
        '
        Me.CDMdraw_xfqx.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CDMdraw_xfqx.BackColor = System.Drawing.SystemColors.Control
        Me.CDMdraw_xfqx.Cursor = System.Windows.Forms.Cursors.Default
        Me.CDMdraw_xfqx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CDMdraw_xfqx.Location = New System.Drawing.Point(917, 387)
        Me.CDMdraw_xfqx.Name = "CDMdraw_xfqx"
        Me.CDMdraw_xfqx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CDMdraw_xfqx.Size = New System.Drawing.Size(180, 34)
        Me.CDMdraw_xfqx.TabIndex = 55
        Me.CDMdraw_xfqx.Text = "绘载荷性能曲线[&H]"
        Me.CDMdraw_xfqx.UseVisualStyleBackColor = False
        '
        'zhxnqx_delete
        '
        Me.zhxnqx_delete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.zhxnqx_delete.BackColor = System.Drawing.SystemColors.Control
        Me.zhxnqx_delete.Cursor = System.Windows.Forms.Cursors.Default
        Me.zhxnqx_delete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.zhxnqx_delete.Location = New System.Drawing.Point(917, 586)
        Me.zhxnqx_delete.Name = "zhxnqx_delete"
        Me.zhxnqx_delete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.zhxnqx_delete.Size = New System.Drawing.Size(180, 34)
        Me.zhxnqx_delete.TabIndex = 56
        Me.zhxnqx_delete.Text = "移除[&E]"
        Me.zhxnqx_delete.UseVisualStyleBackColor = False
        '
        'zhxnqx_save
        '
        Me.zhxnqx_save.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.zhxnqx_save.BackColor = System.Drawing.SystemColors.Control
        Me.zhxnqx_save.Cursor = System.Windows.Forms.Cursors.Default
        Me.zhxnqx_save.ForeColor = System.Drawing.SystemColors.ControlText
        Me.zhxnqx_save.Location = New System.Drawing.Point(917, 540)
        Me.zhxnqx_save.Name = "zhxnqx_save"
        Me.zhxnqx_save.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.zhxnqx_save.Size = New System.Drawing.Size(180, 34)
        Me.zhxnqx_save.TabIndex = 57
        Me.zhxnqx_save.Text = "添加[&A]"
        Me.zhxnqx_save.UseVisualStyleBackColor = False
        '
        '_Frame_22
        '
        Me._Frame_22.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Frame_22.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_22.Controls.Add(Me._Frame_23)
        Me._Frame_22.Controls.Add(Me._Frame_25)
        Me._Frame_22.Controls.Add(Me.Text23)
        Me._Frame_22.Controls.Add(Me.Text24)
        Me._Frame_22.Controls.Add(Me.Text25)
        Me._Frame_22.Controls.Add(Me._Label_23)
        Me._Frame_22.Controls.Add(Me._Label_25)
        Me._Frame_22.Controls.Add(Me._Label_26)
        Me._Frame_22.Controls.Add(Me._Frame_24)
        Me._Frame_22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_22, CType(22, Short))
        Me._Frame_22.Location = New System.Drawing.Point(917, 421)
        Me._Frame_22.Name = "_Frame_22"
        Me._Frame_22.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_22.Size = New System.Drawing.Size(180, 113)
        Me._Frame_22.TabIndex = 58
        Me._Frame_22.TabStop = False
        '
        '_Frame_23
        '
        Me._Frame_23.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_23, CType(23, Short))
        Me._Frame_23.Location = New System.Drawing.Point(0, 43)
        Me._Frame_23.Name = "_Frame_23"
        Me._Frame_23.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_23.Size = New System.Drawing.Size(178, 2)
        Me._Frame_23.TabIndex = 64
        Me._Frame_23.TabStop = False
        '
        '_Frame_25
        '
        Me._Frame_25.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_25, CType(25, Short))
        Me._Frame_25.Location = New System.Drawing.Point(0, 77)
        Me._Frame_25.Name = "_Frame_25"
        Me._Frame_25.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_25.Size = New System.Drawing.Size(178, 2)
        Me._Frame_25.TabIndex = 62
        Me._Frame_25.TabStop = False
        '
        'Text23
        '
        Me.Text23.AcceptsReturn = True
        Me.Text23.BackColor = System.Drawing.SystemColors.Window
        Me.Text23.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text23.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text23.Location = New System.Drawing.Point(84, 14)
        Me.Text23.MaxLength = 0
        Me.Text23.Name = "Text23"
        Me.Text23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text23.Size = New System.Drawing.Size(90, 23)
        Me.Text23.TabIndex = 61
        '
        'Text24
        '
        Me.Text24.AcceptsReturn = True
        Me.Text24.BackColor = System.Drawing.SystemColors.Window
        Me.Text24.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text24.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text24.Location = New System.Drawing.Point(84, 49)
        Me.Text24.MaxLength = 0
        Me.Text24.Name = "Text24"
        Me.Text24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text24.Size = New System.Drawing.Size(90, 23)
        Me.Text24.TabIndex = 60
        '
        'Text25
        '
        Me.Text25.AcceptsReturn = True
        Me.Text25.BackColor = System.Drawing.SystemColors.Window
        Me.Text25.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text25.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text25.Location = New System.Drawing.Point(84, 83)
        Me.Text25.MaxLength = 0
        Me.Text25.Name = "Text25"
        Me.Text25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text25.Size = New System.Drawing.Size(90, 23)
        Me.Text25.TabIndex = 59
        '
        '_Label_23
        '
        Me._Label_23.AutoSize = True
        Me._Label_23.BackColor = System.Drawing.SystemColors.Control
        Me._Label_23.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_23, CType(23, Short))
        Me._Label_23.Location = New System.Drawing.Point(23, 17)
        Me._Label_23.Name = "_Label_23"
        Me._Label_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_23.Size = New System.Drawing.Size(35, 14)
        Me._Label_23.TabIndex = 67
        Me._Label_23.Text = "序号"
        '
        '_Label_25
        '
        Me._Label_25.AutoSize = True
        Me._Label_25.BackColor = System.Drawing.SystemColors.Control
        Me._Label_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_25, CType(25, Short))
        Me._Label_25.Location = New System.Drawing.Point(9, 54)
        Me._Label_25.Name = "_Label_25"
        Me._Label_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_25.Size = New System.Drawing.Size(63, 14)
        Me._Label_25.TabIndex = 66
        Me._Label_25.Text = "轴力(kN)"
        '
        '_Label_26
        '
        Me._Label_26.AutoSize = True
        Me._Label_26.BackColor = System.Drawing.SystemColors.Control
        Me._Label_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_26, CType(26, Short))
        Me._Label_26.Location = New System.Drawing.Point(6, 89)
        Me._Label_26.Name = "_Label_26"
        Me._Label_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_26.Size = New System.Drawing.Size(70, 14)
        Me._Label_26.TabIndex = 65
        Me._Label_26.Text = "压差(MPa)"
        '
        '_Frame_24
        '
        Me._Frame_24.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_24, CType(24, Short))
        Me._Frame_24.Location = New System.Drawing.Point(78, 0)
        Me._Frame_24.Name = "_Frame_24"
        Me._Frame_24.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_24.Size = New System.Drawing.Size(2, 113)
        Me._Frame_24.TabIndex = 63
        Me._Frame_24.TabStop = False
        '
        '_Frame_20
        '
        Me._Frame_20.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Frame_20.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_20.Controls.Add(Me.iPlotX1)
        Me._Frame_20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_20, CType(20, Short))
        Me._Frame_20.Location = New System.Drawing.Point(527, 17)
        Me._Frame_20.Name = "_Frame_20"
        Me._Frame_20.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_20.Size = New System.Drawing.Size(567, 364)
        Me._Frame_20.TabIndex = 68
        Me._Frame_20.TabStop = False
        '
        'iPlotX1
        '
        Me.iPlotX1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.iPlotX1.Enabled = True
        Me.iPlotX1.Location = New System.Drawing.Point(6, 14)
        Me.iPlotX1.Name = "iPlotX1"
        Me.iPlotX1.OcxState = CType(resources.GetObject("iPlotX1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.iPlotX1.Size = New System.Drawing.Size(556, 347)
        Me.iPlotX1.TabIndex = 69
        '
        '_Frame_21
        '
        Me._Frame_21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Frame_21.BackColor = System.Drawing.SystemColors.Control
        Me._Frame_21.Controls.Add(Me.DataGridView2)
        Me._Frame_21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.SetIndex(Me._Frame_21, CType(21, Short))
        Me._Frame_21.Location = New System.Drawing.Point(647, 396)
        Me._Frame_21.Name = "_Frame_21"
        Me._Frame_21.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame_21.Size = New System.Drawing.Size(264, 324)
        Me._Frame_21.TabIndex = 70
        Me._Frame_21.TabStop = False
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.BackgroundColor = System.Drawing.SystemColors.Control
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(5, 13)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowTemplate.Height = 23
        Me.DataGridView2.Size = New System.Drawing.Size(254, 301)
        Me.DataGridView2.TabIndex = 0
        '
        'Text9
        '
        Me.Text9.AcceptsReturn = True
        Me.Text9.BackColor = System.Drawing.SystemColors.Window
        Me.TableLayoutPanel1.SetColumnSpan(Me.Text9, 3)
        Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text9.Location = New System.Drawing.Point(100, 6)
        Me.Text9.MaxLength = 0
        Me.Text9.Name = "Text9"
        Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text9.Size = New System.Drawing.Size(247, 23)
        Me.Text9.TabIndex = 24
        Me.Text9.Text = "Text9"
        '
        'Text8
        '
        Me.Text8.AcceptsReturn = True
        Me.Text8.BackColor = System.Drawing.SystemColors.Window
        Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text8.Location = New System.Drawing.Point(274, 38)
        Me.Text8.MaxLength = 0
        Me.Text8.Name = "Text8"
        Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text8.Size = New System.Drawing.Size(73, 23)
        Me.Text8.TabIndex = 23
        Me.Text8.Text = "Text8"
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.Location = New System.Drawing.Point(274, 70)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(73, 23)
        Me.Text3.TabIndex = 22
        Me.Text3.Text = "Text3"
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(100, 70)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(71, 23)
        Me.Text2.TabIndex = 21
        Me.Text2.Text = "Text2"
        '
        'Text11
        '
        Me.Text11.AcceptsReturn = True
        Me.Text11.BackColor = System.Drawing.SystemColors.Window
        Me.Text11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text11.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text11.Location = New System.Drawing.Point(100, 38)
        Me.Text11.MaxLength = 0
        Me.Text11.Name = "Text11"
        Me.Text11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text11.Size = New System.Drawing.Size(71, 23)
        Me.Text11.TabIndex = 11
        Me.Text11.Text = "Text11"
        '
        '_Label_17
        '
        Me._Label_17.AutoSize = True
        Me._Label_17.BackColor = System.Drawing.SystemColors.Control
        Me._Label_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_17.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_17, CType(17, Short))
        Me._Label_17.Location = New System.Drawing.Point(6, 3)
        Me._Label_17.Name = "_Label_17"
        Me._Label_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_17.Size = New System.Drawing.Size(85, 29)
        Me._Label_17.TabIndex = 20
        Me._Label_17.Text = "型号"
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
        Me._Label_16.Location = New System.Drawing.Point(180, 35)
        Me._Label_16.Name = "_Label_16"
        Me._Label_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_16.Size = New System.Drawing.Size(85, 29)
        Me._Label_16.TabIndex = 15
        Me._Label_16.Text = "长度(m)"
        Me._Label_16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_15
        '
        Me._Label_15.AutoSize = True
        Me._Label_15.BackColor = System.Drawing.SystemColors.Control
        Me._Label_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_15.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_15, CType(15, Short))
        Me._Label_15.Location = New System.Drawing.Point(180, 67)
        Me._Label_15.Name = "_Label_15"
        Me._Label_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_15.Size = New System.Drawing.Size(85, 30)
        Me._Label_15.TabIndex = 14
        Me._Label_15.Text = "最小通径(mm)"
        Me._Label_15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_14
        '
        Me._Label_14.AutoSize = True
        Me._Label_14.BackColor = System.Drawing.SystemColors.Control
        Me._Label_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_14.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_14, CType(14, Short))
        Me._Label_14.Location = New System.Drawing.Point(6, 67)
        Me._Label_14.Name = "_Label_14"
        Me._Label_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_14.Size = New System.Drawing.Size(85, 30)
        Me._Label_14.TabIndex = 13
        Me._Label_14.Text = "最大外径(mm)"
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
        Me._Label_13.Location = New System.Drawing.Point(6, 35)
        Me._Label_13.Name = "_Label_13"
        Me._Label_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_13.Size = New System.Drawing.Size(85, 29)
        Me._Label_13.TabIndex = 9
        Me._Label_13.Text = "工具类型"
        Me._Label_13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Command1
        '
        Me.Command1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(917, 632)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(180, 34)
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
        Me.cmdClose.Location = New System.Drawing.Point(917, 678)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(180, 34)
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
        Me.cmdUpdate.Location = New System.Drawing.Point(530, 460)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(111, 42)
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
        Me.cmdDelete.Location = New System.Drawing.Point(530, 404)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(111, 42)
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
        Me.Frame1.Location = New System.Drawing.Point(6, 17)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(517, 364)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
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
        Me.DataGridView1.Location = New System.Drawing.Point(6, 14)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(504, 346)
        Me.DataGridView1.TabIndex = 0
        '
        '_Label_19
        '
        Me._Label_19.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me._Label_19.AutoSize = True
        Me._Label_19.BackColor = System.Drawing.SystemColors.Control
        Me._Label_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_19, CType(19, Short))
        Me._Label_19.Location = New System.Drawing.Point(9, 384)
        Me._Label_19.Name = "_Label_19"
        Me._Label_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_19.Size = New System.Drawing.Size(119, 14)
        Me._Label_19.TabIndex = 81
        Me._Label_19.Text = "选择开关工具类型"
        '
        '_Label_12
        '
        Me._Label_12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me._Label_12.AutoSize = True
        Me._Label_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_12, CType(12, Short))
        Me._Label_12.Location = New System.Drawing.Point(170, 384)
        Me._Label_12.Name = "_Label_12"
        Me._Label_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_12.Size = New System.Drawing.Size(119, 14)
        Me._Label_12.TabIndex = 7
        Me._Label_12.Text = "开关工具关键参数"
        '
        '_Label_11
        '
        Me._Label_11.AutoSize = True
        Me._Label_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_11, CType(11, Short))
        Me._Label_11.Location = New System.Drawing.Point(9, 6)
        Me._Label_11.Name = "_Label_11"
        Me._Label_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_11.Size = New System.Drawing.Size(91, 14)
        Me._Label_11.TabIndex = 6
        Me._Label_11.Text = "开关工具选择"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(174, 509)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(119, 14)
        Me.Label1.TabIndex = 88
        Me.Label1.Text = "开关工具其它参数"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(569, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(91, 14)
        Me.Label2.TabIndex = 89
        Me.Label2.Text = "载荷性能曲线"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(647, 384)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(119, 14)
        Me.Label3.TabIndex = 95
        Me.Label3.Text = "载荷性能曲线参数"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.0!))
        Me.TableLayoutPanel1.Controls.Add(Me._Label_17, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text3, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text8, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text2, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text9, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_15, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_13, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text11, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_16, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_14, 0, 2)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(170, 404)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(353, 100)
        Me.TableLayoutPanel1.TabIndex = 96
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel2.ColumnCount = 6
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 19.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel2.Controls.Add(Me._Label_2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text12, 1, 5)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_9, 0, 5)
        Me.TableLayoutPanel2.Controls.Add(Me.Text17, 5, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Text1, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_5, 4, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text4, 5, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Text15, 4, 4)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_3, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Text13, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_1, 4, 1)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_8, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text14, 1, 4)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_4, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Text16, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text7, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_18, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text18, 3, 2)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_6, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Text6, 3, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_10, 4, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label4, 4, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_7, 3, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Text10, 5, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Text5, 5, 3)
        Me.TableLayoutPanel2.Controls.Add(Me._Label_0, 0, 4)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(6, 530)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 6
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(635, 188)
        Me.TableLayoutPanel2.TabIndex = 97
        '
        'frmdb_kgyj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1099, 725)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me._Label_19)
        Me.Controls.Add(Me._Frame_12)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CDMdraw_xfqx)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me._Label_11)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me._Label_12)
        Me.Controls.Add(Me._Frame_22)
        Me.Controls.Add(Me.zhxnqx_delete)
        Me.Controls.Add(Me.zhxnqx_save)
        Me.Controls.Add(Me._Frame_20)
        Me.Controls.Add(Me._Frame_21)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(73, 22)
        Me.MinimizeBox = False
        Me.Name = "frmdb_kgyj"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "开关工具数据维护"
        Me._Frame_12.ResumeLayout(False)
        Me._Frame_22.ResumeLayout(False)
        Me._Frame_22.PerformLayout()
        Me._Frame_20.ResumeLayout(False)
        CType(Me.iPlotX1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Frame_21.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
#Region "Upgrade Support"
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents BindingSource2 As System.Windows.Forms.BindingSource
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
#End Region 
End Class