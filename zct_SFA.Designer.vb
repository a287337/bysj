<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class zct_main
#Region "Windows 窗体设计器生成的代码 "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        '此调用是 Windows 窗体设计器所必需的。
        InitializeComponent()
        'MDIForm_Initialize_renamed()
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
    Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器来修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(zct_main))
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
        Me.wenjgl = New System.Windows.Forms.ToolStripMenuItem
        Me.newfile = New System.Windows.Forms.ToolStripMenuItem
        Me.open = New System.Windows.Forms.ToolStripMenuItem
        Me.saveas = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line11 = New System.Windows.Forms.ToolStripSeparator
        Me.tuichu = New System.Windows.Forms.ToolStripMenuItem
        Me.welldate = New System.Windows.Forms.ToolStripMenuItem
        Me.new_se_well = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line21 = New System.Windows.Forms.ToolStripSeparator
        Me.jing_xie = New System.Windows.Forms.ToolStripMenuItem
        Me.jing_shen = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line22 = New System.Windows.Forms.ToolStripSeparator
        Me.Menu_mdc = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line23 = New System.Windows.Forms.ToolStripSeparator
        Me.zy_selt = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line24 = New System.Windows.Forms.ToolStripSeparator
        Me.guan_zhu1 = New System.Windows.Forms.ToolStripMenuItem
        Me.div_line25 = New System.Windows.Forms.ToolStripSeparator
        Me.gongkuan_ipt = New System.Windows.Forms.ToolStripMenuItem
        Me.GH = New System.Windows.Forms.ToolStripMenuItem
        Me.ans_press = New System.Windows.Forms.ToolStripMenuItem
        Me.database = New System.Windows.Forms.ToolStripMenuItem
        Me.zuang = New System.Windows.Forms.ToolStripMenuItem
        Me.taog = New System.Windows.Forms.ToolStripMenuItem
        Me.youg = New System.Windows.Forms.ToolStripMenuItem
        Me.fenggq = New System.Windows.Forms.ToolStripMenuItem
        Me.msgj = New System.Windows.Forms.ToolStripMenuItem
        Me.openelem = New System.Windows.Forms.ToolStripMenuItem
        Me.jieliuyuanian = New System.Windows.Forms.ToolStripMenuItem
        Me.flexele = New System.Windows.Forms.ToolStripMenuItem
        Me.tuichu2 = New System.Windows.Forms.ToolStripMenuItem
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'wenjgl
        '
        Me.wenjgl.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.newfile, Me.open, Me.saveas, Me.div_line11, Me.tuichu})
        Me.wenjgl.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.wenjgl.Name = "wenjgl"
        Me.wenjgl.Size = New System.Drawing.Size(92, 21)
        Me.wenjgl.Text = "【文件管理】"
        '
        'newfile
        '
        Me.newfile.Name = "newfile"
        Me.newfile.Size = New System.Drawing.Size(127, 22)
        Me.newfile.Text = "新建[&N]"
        '
        'open
        '
        Me.open.Name = "open"
        Me.open.Size = New System.Drawing.Size(127, 22)
        Me.open.Text = "打开[&O]"
        '
        'saveas
        '
        Me.saveas.Name = "saveas"
        Me.saveas.Size = New System.Drawing.Size(127, 22)
        Me.saveas.Text = "另存为[&S]"
        '
        'div_line11
        '
        Me.div_line11.Name = "div_line11"
        Me.div_line11.Size = New System.Drawing.Size(124, 6)
        '
        'tuichu
        '
        Me.tuichu.Name = "tuichu"
        Me.tuichu.Size = New System.Drawing.Size(127, 22)
        Me.tuichu.Text = "退出[&X]"
        '
        'welldate
        '
        Me.welldate.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.new_se_well, Me.div_line21, Me.jing_xie, Me.jing_shen, Me.div_line22, Me.Menu_mdc, Me.div_line23, Me.zy_selt, Me.div_line24, Me.guan_zhu1, Me.div_line25, Me.gongkuan_ipt})
        Me.welldate.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.welldate.Name = "welldate"
        Me.welldate.Size = New System.Drawing.Size(128, 21)
        Me.welldate.Text = "【数据输入与管理】"
        '
        'new_se_well
        '
        Me.new_se_well.Name = "new_se_well"
        Me.new_se_well.Size = New System.Drawing.Size(178, 22)
        Me.new_se_well.Text = "油气井基本参数[&N]"
        '
        'div_line21
        '
        Me.div_line21.Name = "div_line21"
        Me.div_line21.Size = New System.Drawing.Size(175, 6)
        '
        'jing_xie
        '
        Me.jing_xie.Name = "jing_xie"
        Me.jing_xie.Size = New System.Drawing.Size(178, 22)
        Me.jing_xie.Text = "井眼轨迹[&J]"
        '
        'jing_shen
        '
        Me.jing_shen.Name = "jing_shen"
        Me.jing_shen.Size = New System.Drawing.Size(178, 22)
        Me.jing_shen.Text = "井身结构[&T]"
        '
        'div_line22
        '
        Me.div_line22.Name = "div_line22"
        Me.div_line22.Size = New System.Drawing.Size(175, 6)
        '
        'Menu_mdc
        '
        Me.Menu_mdc.Name = "Menu_mdc"
        Me.Menu_mdc.Size = New System.Drawing.Size(178, 22)
        Me.Menu_mdc.Text = "目的层参数[&D]"
        '
        'div_line23
        '
        Me.div_line23.Name = "div_line23"
        Me.div_line23.Size = New System.Drawing.Size(175, 6)
        Me.div_line23.Visible = False
        '
        'zy_selt
        '
        Me.zy_selt.Name = "zy_selt"
        Me.zy_selt.Size = New System.Drawing.Size(178, 22)
        Me.zy_selt.Text = "作业名称[&Y]"
        '
        'div_line24
        '
        Me.div_line24.Name = "div_line24"
        Me.div_line24.Size = New System.Drawing.Size(175, 6)
        '
        'guan_zhu1
        '
        Me.guan_zhu1.Name = "guan_zhu1"
        Me.guan_zhu1.Size = New System.Drawing.Size(178, 22)
        Me.guan_zhu1.Text = "作业管柱组合[&Z]"
        '
        'div_line25
        '
        Me.div_line25.Name = "div_line25"
        Me.div_line25.Size = New System.Drawing.Size(175, 6)
        '
        'gongkuan_ipt
        '
        Me.gongkuan_ipt.Name = "gongkuan_ipt"
        Me.gongkuan_ipt.Size = New System.Drawing.Size(178, 22)
        Me.gongkuan_ipt.Text = "作业工况参数[&G]"
        '
        'GH
        '
        Me.GH.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ans_press})
        Me.GH.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.GH.Name = "GH"
        Me.GH.Size = New System.Drawing.Size(92, 21)
        Me.GH.Text = "【分析运算】"
        '
        'ans_press
        '
        Me.ans_press.Name = "ans_press"
        Me.ans_press.Size = New System.Drawing.Size(206, 22)
        Me.ans_press.Text = "井下管柱压力场分析PR&S"
        Me.ans_press.Visible = False
        '
        'database
        '
        Me.database.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.zuang, Me.taog, Me.youg, Me.fenggq, Me.msgj, Me.openelem, Me.jieliuyuanian, Me.flexele})
        Me.database.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.database.Name = "database"
        Me.database.Size = New System.Drawing.Size(116, 21)
        Me.database.Text = "【基础数据管理】"
        '
        'zuang
        '
        Me.zuang.Name = "zuang"
        Me.zuang.Size = New System.Drawing.Size(152, 22)
        Me.zuang.Text = "钻 杆[&D]"
        '
        'taog
        '
        Me.taog.Name = "taog"
        Me.taog.Size = New System.Drawing.Size(152, 22)
        Me.taog.Text = "套 管[&O]"
        '
        'youg
        '
        Me.youg.Name = "youg"
        Me.youg.Size = New System.Drawing.Size(152, 22)
        Me.youg.Text = "油 管[&T]"
        '
        'fenggq
        '
        Me.fenggq.Name = "fenggq"
        Me.fenggq.Size = New System.Drawing.Size(152, 22)
        Me.fenggq.Text = "封隔器[&P]"
        '
        'msgj
        '
        Me.msgj.Name = "msgj"
        Me.msgj.Size = New System.Drawing.Size(152, 22)
        Me.msgj.Text = "锚定工具[&M]"
        Me.msgj.Visible = False
        '
        'openelem
        '
        Me.openelem.Name = "openelem"
        Me.openelem.Size = New System.Drawing.Size(152, 22)
        Me.openelem.Text = "开关元件[&W]"
        Me.openelem.Visible = False
        '
        'jieliuyuanian
        '
        Me.jieliuyuanian.Name = "jieliuyuanian"
        Me.jieliuyuanian.Size = New System.Drawing.Size(152, 22)
        Me.jieliuyuanian.Text = "节流元件[J]"
        Me.jieliuyuanian.Visible = False
        '
        'flexele
        '
        Me.flexele.Name = "flexele"
        Me.flexele.Size = New System.Drawing.Size(152, 22)
        Me.flexele.Text = "伸缩元件[&S]"
        Me.flexele.Visible = False
        '
        'tuichu2
        '
        Me.tuichu2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.tuichu2.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.tuichu2.Name = "tuichu2"
        Me.tuichu2.Size = New System.Drawing.Size(68, 21)
        Me.tuichu2.Text = "【退出】"
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.wenjgl, Me.welldate, Me.GH, Me.database, Me.tuichu2})
        Me.MainMenu1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.MdiWindowListItem = Me.wenjgl
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1245, 25)
        Me.MainMenu1.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1245, 550)
        Me.Panel1.TabIndex = 5
        '
        'zct_main
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1245, 575)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Location = New System.Drawing.Point(24, 184)
        Me.MainMenuStrip = Me.MainMenu1
        Me.Name = "zct_main"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "钻油管传输射孔力学分析软件"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents wenjgl As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents newfile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents open As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents saveas As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents div_line11 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents tuichu As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents welldate As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents new_se_well As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents div_line21 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents jing_xie As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents jing_shen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents div_line22 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents div_line23 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents zy_selt As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents div_line24 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents guan_zhu1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents div_line25 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents gongkuan_ipt As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents GH As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents ans_press As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents database As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents zuang As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents taog As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents youg As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents fenggq As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents msgj As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents openelem As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents jieliuyuanian As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents flexele As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents tuichu2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Menu_mdc As System.Windows.Forms.ToolStripMenuItem
#End Region
End Class