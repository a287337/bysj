<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmipt_testwell
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
        'VB6_AddADODataBinding()
	End Sub
	'Form 重写 Dispose，以清理组件列表。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
            'VB6_RemoveADODataBinding()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows 窗体设计器所必需的
	Private components As System.ComponentModel.IContainer
    Private ADOBind_Adodc1 As VB6.MBindingCollection
	Private ADOBind_Adodc4 As VB6.MBindingCollection
    'Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
    Public WithEvents Frame As System.Windows.Forms.GroupBox
    Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents Text2 As System.Windows.Forms.TextBox
    Public WithEvents Text3 As System.Windows.Forms.TextBox
    Public WithEvents Text4 As System.Windows.Forms.TextBox
    Public WithEvents Text5 As System.Windows.Forms.TextBox
    Public WithEvents Text6 As System.Windows.Forms.TextBox
    Public WithEvents _Label_2 As System.Windows.Forms.Label
    Public WithEvents _Label_3 As System.Windows.Forms.Label
    Public WithEvents _Label_4 As System.Windows.Forms.Label
    Public WithEvents _Label_5 As System.Windows.Forms.Label
    Public WithEvents _Label_6 As System.Windows.Forms.Label
    Public WithEvents _Label_7 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器来修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmipt_testwell))
        Me.Frame = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.Text2 = New System.Windows.Forms.TextBox
        Me.Text3 = New System.Windows.Forms.TextBox
        Me.Text4 = New System.Windows.Forms.TextBox
        Me.Text5 = New System.Windows.Forms.TextBox
        Me.Text6 = New System.Windows.Forms.TextBox
        Me._Label_2 = New System.Windows.Forms.Label
        Me._Label_3 = New System.Windows.Forms.Label
        Me._Label_4 = New System.Windows.Forms.Label
        Me._Label_5 = New System.Windows.Forms.Label
        Me._Label_6 = New System.Windows.Forms.Label
        Me._Label_7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdImport2 = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdInput = New System.Windows.Forms.Button
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.Text7 = New System.Windows.Forms.TextBox
        Me.Text10 = New System.Windows.Forms.TextBox
        Me.Text9 = New System.Windows.Forms.TextBox
        Me.Text8 = New System.Windows.Forms.TextBox
        Me._Label_1 = New System.Windows.Forms.Label
        Me._Label_32 = New System.Windows.Forms.Label
        Me._Label_31 = New System.Windows.Forms.Label
        Me._Label_30 = New System.Windows.Forms.Label
        Me._Label_70 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Label2 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.CMDsave = New System.Windows.Forms.Button
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel
        Me.Frame.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame
        '
        Me.Frame.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Frame.BackColor = System.Drawing.SystemColors.Control
        Me.Frame.Controls.Add(Me.DataGridView1)
        Me.Frame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.Location = New System.Drawing.Point(3, 100)
        Me.Frame.Name = "Frame"
        Me.Frame.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame.Size = New System.Drawing.Size(730, 227)
        Me.Frame.TabIndex = 6
        Me.Frame.TabStop = False
        Me.Frame.Text = "井眼轨迹参数列表"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.DataSource = Me.BindingSource1
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.Location = New System.Drawing.Point(5, 20)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.DataGridView1.RowHeadersWidth = 24
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(720, 199)
        Me.DataGridView1.TabIndex = 56
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(6, 40)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(206, 23)
        Me.Text1.TabIndex = 29
        Me.Text1.Text = "Text1"
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(221, 40)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.ReadOnly = True
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(206, 23)
        Me.Text2.TabIndex = 28
        Me.Text2.Text = "Text2"
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.Location = New System.Drawing.Point(436, 40)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.ReadOnly = True
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(64, 23)
        Me.Text3.TabIndex = 27
        Me.Text3.Text = "Text3"
        '
        'Text4
        '
        Me.Text4.AcceptsReturn = True
        Me.Text4.BackColor = System.Drawing.SystemColors.Window
        Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text4.Location = New System.Drawing.Point(509, 40)
        Me.Text4.MaxLength = 0
        Me.Text4.Name = "Text4"
        Me.Text4.ReadOnly = True
        Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text4.Size = New System.Drawing.Size(64, 23)
        Me.Text4.TabIndex = 26
        Me.Text4.Text = "Text4"
        Me.Text4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text5
        '
        Me.Text5.AcceptsReturn = True
        Me.Text5.BackColor = System.Drawing.SystemColors.Window
        Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text5.Location = New System.Drawing.Point(582, 40)
        Me.Text5.MaxLength = 0
        Me.Text5.Name = "Text5"
        Me.Text5.ReadOnly = True
        Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text5.Size = New System.Drawing.Size(64, 23)
        Me.Text5.TabIndex = 25
        Me.Text5.Text = "Text5"
        Me.Text5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text6
        '
        Me.Text6.AcceptsReturn = True
        Me.Text6.BackColor = System.Drawing.SystemColors.Window
        Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text6.Location = New System.Drawing.Point(655, 40)
        Me.Text6.MaxLength = 0
        Me.Text6.Name = "Text6"
        Me.Text6.ReadOnly = True
        Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text6.Size = New System.Drawing.Size(69, 23)
        Me.Text6.TabIndex = 24
        Me.Text6.Text = "Text6"
        '
        '_Label_2
        '
        Me._Label_2.AutoSize = True
        Me._Label_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_2.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_2.Location = New System.Drawing.Point(6, 3)
        Me._Label_2.Name = "_Label_2"
        Me._Label_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_2.Size = New System.Drawing.Size(206, 31)
        Me._Label_2.TabIndex = 35
        Me._Label_2.Text = "地理位置"
        Me._Label_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_3
        '
        Me._Label_3.AutoSize = True
        Me._Label_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_3.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_3.Location = New System.Drawing.Point(509, 3)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(64, 31)
        Me._Label_3.TabIndex = 34
        Me._Label_3.Text = "设计井深"
        Me._Label_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_4
        '
        Me._Label_4.AutoSize = True
        Me._Label_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_4.Location = New System.Drawing.Point(582, 3)
        Me._Label_4.Name = "_Label_4"
        Me._Label_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_4.Size = New System.Drawing.Size(64, 31)
        Me._Label_4.TabIndex = 33
        Me._Label_4.Text = "完钻井深"
        Me._Label_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_5
        '
        Me._Label_5.AutoSize = True
        Me._Label_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_5.Location = New System.Drawing.Point(221, 3)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(206, 31)
        Me._Label_5.TabIndex = 32
        Me._Label_5.Text = "构造位置"
        Me._Label_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_6
        '
        Me._Label_6.AutoSize = True
        Me._Label_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_6.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_6.Location = New System.Drawing.Point(436, 3)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(64, 31)
        Me._Label_6.TabIndex = 31
        Me._Label_6.Text = "井别"
        Me._Label_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_7
        '
        Me._Label_7.AutoSize = True
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_7.Location = New System.Drawing.Point(655, 3)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(69, 31)
        Me._Label_7.TabIndex = 30
        Me._Label_7.Text = "完钻层位"
        Me._Label_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(8, 330)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(490, 14)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "注意：Excel导入时，须按照""井斜数据表导入样表""的格式设定列及表页名称！"
        '
        'cmdImport2
        '
        Me.cmdImport2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdImport2.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImport2.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImport2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImport2.Location = New System.Drawing.Point(458, 403)
        Me.cmdImport2.Name = "cmdImport2"
        Me.cmdImport2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImport2.Size = New System.Drawing.Size(268, 36)
        Me.cmdImport2.TabIndex = 49
        Me.cmdImport2.Text = "用打开EXcel对象的方法导入..."
        Me.cmdImport2.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(604, 541)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(125, 45)
        Me.cmdClose.TabIndex = 48
        Me.cmdClose.Text = "退出[&X]"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdHelp.Location = New System.Drawing.Point(454, 541)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdHelp.Size = New System.Drawing.Size(125, 45)
        Me.cmdHelp.TabIndex = 47
        Me.cmdHelp.Text = "帮助[&H]"
        Me.cmdHelp.UseVisualStyleBackColor = False
        '
        'cmdInput
        '
        Me.cmdInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdInput.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInput.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInput.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInput.Location = New System.Drawing.Point(458, 352)
        Me.cmdInput.Name = "cmdInput"
        Me.cmdInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInput.Size = New System.Drawing.Size(268, 36)
        Me.cmdInput.TabIndex = 46
        Me.cmdInput.Text = "用将Excel作为ODBC数据源的方法导入..."
        Me.cmdInput.UseVisualStyleBackColor = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(154, 541)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(125, 45)
        Me.cmdUpdate.TabIndex = 45
        Me.cmdUpdate.Text = "添加/更改[&A]"
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(5, 541)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(125, 45)
        Me.cmdDelete.TabIndex = 44
        Me.cmdDelete.Text = "删除[&D]"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'Text7
        '
        Me.Text7.AcceptsReturn = True
        Me.Text7.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Text7.BackColor = System.Drawing.SystemColors.Window
        Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text7.Location = New System.Drawing.Point(6, 40)
        Me.Text7.MaxLength = 0
        Me.Text7.Name = "Text7"
        Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text7.Size = New System.Drawing.Size(172, 23)
        Me.Text7.TabIndex = 19
        Me.Text7.Text = "Text7"
        Me.Text7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text10
        '
        Me.Text10.AcceptsReturn = True
        Me.Text10.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Text10.BackColor = System.Drawing.SystemColors.Window
        Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text10.Location = New System.Drawing.Point(549, 40)
        Me.Text10.MaxLength = 0
        Me.Text10.Name = "Text10"
        Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text10.Size = New System.Drawing.Size(175, 23)
        Me.Text10.TabIndex = 12
        Me.Text10.Text = "Text10"
        Me.Text10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text9
        '
        Me.Text9.AcceptsReturn = True
        Me.Text9.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Text9.BackColor = System.Drawing.SystemColors.Window
        Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text9.Location = New System.Drawing.Point(368, 40)
        Me.Text9.MaxLength = 0
        Me.Text9.Name = "Text9"
        Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text9.Size = New System.Drawing.Size(172, 23)
        Me.Text9.TabIndex = 11
        Me.Text9.Text = "Text9"
        Me.Text9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text8
        '
        Me.Text8.AcceptsReturn = True
        Me.Text8.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Text8.BackColor = System.Drawing.SystemColors.Window
        Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text8.Location = New System.Drawing.Point(187, 40)
        Me.Text8.MaxLength = 0
        Me.Text8.Name = "Text8"
        Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text8.Size = New System.Drawing.Size(172, 23)
        Me.Text8.TabIndex = 10
        Me.Text8.Text = "Text8"
        Me.Text8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_Label_1
        '
        Me._Label_1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Label_1.AutoSize = True
        Me._Label_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_1.Location = New System.Drawing.Point(6, 11)
        Me._Label_1.Name = "_Label_1"
        Me._Label_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_1.Size = New System.Drawing.Size(172, 14)
        Me._Label_1.TabIndex = 22
        Me._Label_1.Text = "序号"
        Me._Label_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_32
        '
        Me._Label_32.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Label_32.AutoSize = True
        Me._Label_32.BackColor = System.Drawing.SystemColors.Control
        Me._Label_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_32.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_32.Location = New System.Drawing.Point(549, 11)
        Me._Label_32.Name = "_Label_32"
        Me._Label_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_32.Size = New System.Drawing.Size(175, 14)
        Me._Label_32.TabIndex = 15
        Me._Label_32.Text = "方位角(°)"
        Me._Label_32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_31
        '
        Me._Label_31.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Label_31.AutoSize = True
        Me._Label_31.BackColor = System.Drawing.SystemColors.Control
        Me._Label_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_31.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_31.Location = New System.Drawing.Point(368, 11)
        Me._Label_31.Name = "_Label_31"
        Me._Label_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_31.Size = New System.Drawing.Size(172, 14)
        Me._Label_31.TabIndex = 14
        Me._Label_31.Text = "井斜角(°)"
        Me._Label_31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_30
        '
        Me._Label_30.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me._Label_30.AutoSize = True
        Me._Label_30.BackColor = System.Drawing.SystemColors.Control
        Me._Label_30.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_30.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me._Label_30.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_30.Location = New System.Drawing.Point(187, 12)
        Me._Label_30.Name = "_Label_30"
        Me._Label_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_30.Size = New System.Drawing.Size(172, 12)
        Me._Label_30.TabIndex = 13
        Me._Label_30.Text = "井深(m)"
        Me._Label_30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_Label_70
        '
        Me._Label_70.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me._Label_70.AutoSize = True
        Me._Label_70.BackColor = System.Drawing.SystemColors.Control
        Me._Label_70.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_70.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_70.Location = New System.Drawing.Point(10, 444)
        Me._Label_70.Name = "_Label_70"
        Me._Label_70.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_70.Size = New System.Drawing.Size(91, 14)
        Me._Label_70.TabIndex = 43
        Me._Label_70.Text = "井眼轨迹参数"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Text10, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text7, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Text9, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_30, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Text8, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_31, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me._Label_32, 3, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(4, 463)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(730, 70)
        Me.TableLayoutPanel1.TabIndex = 50
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 14)
        Me.Label2.TabIndex = 52
        Me.Label2.Text = "油气井参数"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CMDsave
        '
        Me.CMDsave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CMDsave.Location = New System.Drawing.Point(303, 541)
        Me.CMDsave.Name = "CMDsave"
        Me.CMDsave.Size = New System.Drawing.Size(125, 45)
        Me.CMDsave.TabIndex = 53
        Me.CMDsave.Text = "保存[&S]"
        Me.CMDsave.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(236, 411)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(207, 18)
        Me.RadioButton4.TabIndex = 65
        Me.RadioButton4.Text = "Microsoft Excel 3.0 工作簿"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(13, 411)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(207, 18)
        Me.RadioButton3.TabIndex = 64
        Me.RadioButton3.Text = "Microsoft Excel 5.0 工作簿"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.BackColor = System.Drawing.SystemColors.Control
        Me.RadioButton2.Location = New System.Drawing.Point(13, 387)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(291, 18)
        Me.RadioButton2.TabIndex = 63
        Me.RadioButton2.Text = "Microsoft Excel 5.0 和 7.0 (95) 工作簿"
        Me.RadioButton2.UseVisualStyleBackColor = False
        '
        'RadioButton1
        '
        Me.RadioButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(13, 361)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(431, 18)
        Me.RadioButton1.TabIndex = 62
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Microsoft Excel 8.0 (97)、9.0 (2000) 和 10.0 (2002) 工作簿"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(2, 343)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(450, 97)
        Me.GroupBox1.TabIndex = 66
        Me.GroupBox1.TabStop = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel3.ColumnCount = 6
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Text6, 5, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.Text5, 4, 1)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_5, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Text4, 3, 1)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Text3, 2, 1)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_6, 2, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Text2, 1, 1)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_3, 3, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Text1, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_4, 4, 0)
        Me.TableLayoutPanel3.Controls.Add(Me._Label_7, 5, 0)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 27)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(730, 71)
        Me.TableLayoutPanel3.TabIndex = 67
        '
        'frmipt_testwell
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(737, 594)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TableLayoutPanel3)
        Me.Controls.Add(Me.RadioButton4)
        Me.Controls.Add(Me.RadioButton3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.CMDsave)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.cmdImport2)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdInput)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me._Label_70)
        Me.Controls.Add(Me.Frame)
        Me.Controls.Add(Me.GroupBox1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MinimizeBox = False
        Me.Name = "frmipt_testwell"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "井眼轨迹数据输入与管理"
        Me.Frame.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support"
    Public WithEvents cmdImport2 As System.Windows.Forms.Button
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents cmdHelp As System.Windows.Forms.Button
    Public WithEvents cmdInput As System.Windows.Forms.Button
    Public WithEvents cmdUpdate As System.Windows.Forms.Button
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents Text7 As System.Windows.Forms.TextBox
    Public WithEvents Text10 As System.Windows.Forms.TextBox
    Public WithEvents Text9 As System.Windows.Forms.TextBox
    Public WithEvents Text8 As System.Windows.Forms.TextBox
    Public WithEvents _Label_1 As System.Windows.Forms.Label
    Public WithEvents _Label_32 As System.Windows.Forms.Label
    Public WithEvents _Label_31 As System.Windows.Forms.Label
    Public WithEvents _Label_30 As System.Windows.Forms.Label
    Public WithEvents _Label_70 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents CMDsave As System.Windows.Forms.Button
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
#End Region
End Class