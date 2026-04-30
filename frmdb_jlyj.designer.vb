<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmdb_jlyj
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
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
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
	Public WithEvents _Label_2 As System.Windows.Forms.Label
	Public WithEvents _Label_3 As System.Windows.Forms.Label
	Public WithEvents _Label_4 As System.Windows.Forms.Label
	Public WithEvents _Label_10 As System.Windows.Forms.Label
	Public WithEvents _Label_8 As System.Windows.Forms.Label
	Public WithEvents _Label_7 As System.Windows.Forms.Label
	Public WithEvents _Label_6 As System.Windows.Forms.Label
	Public WithEvents _Label_5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents _Label_1 As System.Windows.Forms.Label
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
        Me.List1 = New System.Windows.Forms.ListBox
        Me.Text3 = New System.Windows.Forms.TextBox
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
        Me._Label_2 = New System.Windows.Forms.Label
        Me._Label_3 = New System.Windows.Forms.Label
        Me._Label_4 = New System.Windows.Forms.Label
        Me._Label_10 = New System.Windows.Forms.Label
        Me._Label_8 = New System.Windows.Forms.Label
        Me._Label_7 = New System.Windows.Forms.Label
        Me._Label_6 = New System.Windows.Forms.Label
        Me._Label_5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me._Label_1 = New System.Windows.Forms.Label
        Me._Label_0 = New System.Windows.Forms.Label
        Me.Frame = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
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
        Me.Command1.Location = New System.Drawing.Point(686, 475)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(111, 31)
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
        Me.cmdClose.Location = New System.Drawing.Point(686, 515)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(111, 31)
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
        Me.cmdUpdate.Location = New System.Drawing.Point(686, 435)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(111, 31)
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
        Me.cmdDelete.Location = New System.Drawing.Point(686, 397)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(111, 29)
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
        Me.Frame1.Size = New System.Drawing.Size(791, 382)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "节流元件参数信息"
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
        Me.DataGridView1.Size = New System.Drawing.Size(781, 349)
        Me.DataGridView1.TabIndex = 0
        '
        'List1
        '
        Me.List1.BackColor = System.Drawing.SystemColors.Window
        Me.List1.Cursor = System.Windows.Forms.Cursors.Default
        Me.List1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.List1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.List1.ItemHeight = 14
        Me.List1.Items.AddRange(New Object() {"上下流通", "内外流通"})
        Me.List1.Location = New System.Drawing.Point(573, 108)
        Me.List1.Name = "List1"
        Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.List1.Size = New System.Drawing.Size(95, 32)
        Me.List1.TabIndex = 37
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.Location = New System.Drawing.Point(573, 6)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(95, 23)
        Me.Text3.TabIndex = 35
        '
        'Text11
        '
        Me.Text11.AcceptsReturn = True
        Me.Text11.BackColor = System.Drawing.SystemColors.Window
        Me.Text11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text11.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text11.Location = New System.Drawing.Point(573, 74)
        Me.Text11.MaxLength = 0
        Me.Text11.Name = "Text11"
        Me.Text11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text11.Size = New System.Drawing.Size(95, 23)
        Me.Text11.TabIndex = 34
        '
        'Text10
        '
        Me.Text10.AcceptsReturn = True
        Me.Text10.BackColor = System.Drawing.SystemColors.Window
        Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text10.Location = New System.Drawing.Point(355, 108)
        Me.Text10.MaxLength = 0
        Me.Text10.Name = "Text10"
        Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text10.Size = New System.Drawing.Size(90, 23)
        Me.Text10.TabIndex = 33
        '
        'Text9
        '
        Me.Text9.AcceptsReturn = True
        Me.Text9.BackColor = System.Drawing.SystemColors.Window
        Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text9.Location = New System.Drawing.Point(131, 108)
        Me.Text9.MaxLength = 0
        Me.Text9.Name = "Text9"
        Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text9.Size = New System.Drawing.Size(90, 23)
        Me.Text9.TabIndex = 32
        '
        'Text7
        '
        Me.Text7.AcceptsReturn = True
        Me.Text7.BackColor = System.Drawing.SystemColors.Window
        Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text7.Location = New System.Drawing.Point(131, 74)
        Me.Text7.MaxLength = 0
        Me.Text7.Name = "Text7"
        Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text7.Size = New System.Drawing.Size(90, 23)
        Me.Text7.TabIndex = 17
        '
        'Text6
        '
        Me.Text6.AcceptsReturn = True
        Me.Text6.BackColor = System.Drawing.SystemColors.Window
        Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text6.Location = New System.Drawing.Point(355, 74)
        Me.Text6.MaxLength = 0
        Me.Text6.Name = "Text6"
        Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text6.Size = New System.Drawing.Size(90, 23)
        Me.Text6.TabIndex = 15
        '
        'Text5
        '
        Me.Text5.AcceptsReturn = True
        Me.Text5.BackColor = System.Drawing.SystemColors.Window
        Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text5.Location = New System.Drawing.Point(573, 40)
        Me.Text5.MaxLength = 0
        Me.Text5.Name = "Text5"
        Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text5.Size = New System.Drawing.Size(95, 23)
        Me.Text5.TabIndex = 13
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(355, 40)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(90, 23)
        Me.Text4.TabIndex = 12
        '
        'Text8
        '
        Me.Text8.AcceptsReturn = True
        Me.Text8.BackColor = System.Drawing.SystemColors.Window
        Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text8.Location = New System.Drawing.Point(131, 40)
        Me.Text8.MaxLength = 0
        Me.Text8.Name = "Text8"
        Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text8.Size = New System.Drawing.Size(90, 23)
        Me.Text8.TabIndex = 10
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(355, 6)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(90, 23)
        Me.Text2.TabIndex = 8
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(131, 6)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(90, 23)
        Me.Text1.TabIndex = 6
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
        Me._Label_2.Location = New System.Drawing.Point(6, 3)
        Me._Label_2.Name = "_Label_2"
        Me._Label_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_2.Size = New System.Drawing.Size(116, 31)
        Me._Label_2.TabIndex = 30
        Me._Label_2.Text = "型号规格"
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
        Me._Label_3.Location = New System.Drawing.Point(6, 37)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(116, 31)
        Me._Label_3.TabIndex = 29
        Me._Label_3.Text = "长度(mm)"
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
        Me._Label_4.Location = New System.Drawing.Point(6, 71)
        Me._Label_4.Name = "_Label_4"
        Me._Label_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_4.Size = New System.Drawing.Size(116, 31)
        Me._Label_4.TabIndex = 28
        Me._Label_4.Text = "抗外挤强度(MPa)"
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
        Me._Label_10.Location = New System.Drawing.Point(230, 105)
        Me._Label_10.Name = "_Label_10"
        Me._Label_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_10.Size = New System.Drawing.Size(116, 39)
        Me._Label_10.TabIndex = 27
        Me._Label_10.Text = "节流孔长度(mm)"
        Me._Label_10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me._Label_8.Location = New System.Drawing.Point(454, 71)
        Me._Label_8.Name = "_Label_8"
        Me._Label_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TableLayoutPanel1.SetRowSpan(Me._Label_8, 2)
        Me._Label_8.Size = New System.Drawing.Size(110, 73)
        Me._Label_8.TabIndex = 26
        Me._Label_8.Text = "节流流向"
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
        Me._Label_7.Location = New System.Drawing.Point(6, 105)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(116, 39)
        Me._Label_7.TabIndex = 25
        Me._Label_7.Text = "节流孔内径(mm)"
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
        Me._Label_6.Location = New System.Drawing.Point(230, 71)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(116, 31)
        Me._Label_6.TabIndex = 16
        Me._Label_6.Text = "抗内压强度(MPa)"
        Me._Label_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_5
        '
        Me._Label_5.AutoSize = True
        Me._Label_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_5, CType(5, Short))
        Me._Label_5.Location = New System.Drawing.Point(230, 37)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(116, 31)
        Me._Label_5.TabIndex = 14
        Me._Label_5.Text = "重量(kg)"
        Me._Label_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(454, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(110, 31)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "抗拉强度(kN)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me._Label_1.Location = New System.Drawing.Point(454, 3)
        Me._Label_1.Name = "_Label_1"
        Me._Label_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_1.Size = New System.Drawing.Size(110, 31)
        Me._Label_1.TabIndex = 9
        Me._Label_1.Text = "内径(mm)"
        Me._Label_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me._Label_0.Location = New System.Drawing.Point(230, 3)
        Me._Label_0.Name = "_Label_0"
        Me._Label_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_0.Size = New System.Drawing.Size(116, 31)
        Me._Label_0.TabIndex = 7
        Me._Label_0.Text = "外径(mm)"
        Me._Label_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 6
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.81188!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.85149!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.81188!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.85149!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.82178!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.85149!))
        Me.TableLayoutPanel1.Controls.Add(Me.Text3, 5, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.List1, 5, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_7, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Text9, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Text11, 5, 2)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_10, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Text10, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_8, 4, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text2, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_4, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_0, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text7, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text5, 5, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text6, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text4, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_1, 4, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_6, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Text8, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 4, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_5, 2, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 397)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 24.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 24.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 24.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 28.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(674, 147)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'frmdb_jlyj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(804, 554)
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
        Me.Name = "frmdb_jlyj"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "节流元件数据维护"
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
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
#End Region 
End Class