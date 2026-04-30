'**************************************************************************************************************************************************************
'                                                               管柱力学分析主菜单
'                                                                                                                   秦彦斌 2022年1月20日最后整理更新
'    升级情况说明：
'    由于VS新功能：模块结束后，所调用的表单也同时结束，不能使主菜单处于等待事件的状态，故项目起始从菜单“zct_TSM”开始。
'
'程序升级记事：
'    （1）2021年11月10日-11月14日，以管柱拉伸强度安全分析模块为例，探求界面最大化（缩放）的实现方法。
'    （2）在几个主要的界面最大化（缩放）已实现后，2021年12月19日，试图实现在本主菜单窗口缩放时，正在里边show的界面也能跟着缩放。但是，若利用Resize事件，在用鼠标
'拖动时，会频繁地触发Resize事件，事件响应慢且屏幕会频闪，效果不好；若利用ResizeEnd事件，经测试，点击“最大化”、“还原”按钮不会触发ResizeEnd事件，达不到预期目标。
'另外，如何识别“正在里边show的界面”的方（算）法未找到，故，此想法近期就放一放吧。                                          秦彦斌2021年12月19日备忘
'     上面（2）问题已于20220112前后解决：将窗体的FormBorderStyle属性设为Windows.Forms.FormBorderStyle.Fixed3D，将MaximizeBox属性设为True，用户操作时，利用最
'大化按钮在主菜单窗口内“最大化”和设计之初的大小间切换，实现上述目的。                                                      秦彦斌2022年1月20日备忘                            
'    （3）20220119，封隔器管柱力学分析各模块界面最大化（缩放）功能全部完成'
'
'**************************************************************************************************************************************************************
'知识点：
'                                                      窗口的显示的两种方式：模态显示（showdialog）和非模态显示（show）
'
' show和showdialog区别
'    窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show）。
'区别：
'    模态与非模态窗体的主要区别是窗体显示的时候是否可以操作其他窗体。模态窗体不允许操作其他窗体，非模态窗体可以操作其他窗体。
'
'模态显示后，
'    弹出窗口阻止调用窗口的所有消息响应。
'    只有在弹出窗口结束后调用窗口才能继续。
'    在模态窗口“关闭”后，可以读取模态窗口中信息，包括窗口的返回状态，窗口子控件的值。
'    在调用Form.ShowDialog方法后,直到关闭对话框后，才执行此方法后面的代码  
'    窗体显示为模式窗体时，单击“关闭”按钮会隐藏窗体，并将DialogResult属性设置为DialogResult.Cancel  
'    与无模式窗体不同，当用户单击对话框的关闭窗体按钮或设置DialogResult属性的值时,不调用窗体的Close方法  
'    实际上是把窗体的Visible属性赋值为false,隐藏窗体了  
'    这样隐藏的窗体是可以重新显示，而不用创建该对话框的新实例  
'    因为未关闭窗体,所以在应用程序不再需要该窗体时,请调用该窗体的Dispose方法  
'    所以模态窗口在关闭时，不会调用close方法，也不调用dispose方法，窗口仍然存在，占有资源，所以可以继续获得窗口相关信息。在窗口不再使用时，需要手动释放资源
'    testDialog.ShowDialog(); // 模态窗口关闭后，可以再次显示出来 
'    testDialog.Dispose(); // 当模态窗口不再使用时，应该调用dispose方法释放资源 
'
'非模态显示后，
'    可以在弹出窗口和调用窗口之间随意切换。
'    调用窗口调用show方法后，下面的代码可以立即执行。
'    在非模态窗口关闭后，窗口的所有资源被释放，窗口不存在，无法获取窗口的任何信息。
'
'怎么判断一个窗体是模式窗体呢？  
'    利用Form.Modal属性,如果该窗体是模式显示，则为true,否则为false  
'    根据通过Show和ShowDialog而显示出来的窗体的Modal属性分别对应false和true  
'
'特别注意：  
'    由于在窗体创建之前是无法得知显示方式的,所以在窗体构造函数中,Modal属性总是对应false,所以我们只能在Load事件中或者之后利用Modal属性值
'
'怎么确定窗体间的所有者关系?  
'    Form类的Owner属性:窗体的所有者  
'    当一个窗体归另一窗体所有时，它便随着所有者窗体最小化和关闭。  
'    例如，如果Form2归窗体Form1所有，则关闭或最小化Form1时，Form2也会关闭或最小化。
'**************************************************************************************************************************************************************
Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D ' 导入包含GraphicsPath的命名空间
Imports System.Drawing.Text

Friend Class zct_main
    Inherits System.Windows.Forms.Form
    Dim count_i As Integer
    '**********************************************************************************
    '取得焦点时更新菜单项。
    '      不行，会频闪、死循环的。-----20220628
    '**********************************************************************************
    'Private Sub zct_main_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
    '    'Call my_refresh()
    'End Sub

    '**********************************************************************************
    '窗体加载
    '**********************************************************************************
    Private Sub zct_TSM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '    '**********************************************************************************************************
        '    '(1)进行必要的全局设置
        '    '(2)调用登录模块，若不成功，结束程序。
        '    '**********************************************************************************************************
        '修改主界面背景显示内容            李润洲2025年7月27日
        Panel1.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\封面-PRS.jpg")
        Panel1.BackgroundImageLayout = ImageLayout.Stretch
        Call Main()
        frmSplash.ShowDialog()
        frmSplash.Dispose()
        count_i = 0
        Call my_refresh()
    End Sub

    Public Sub my_refresh()
        On Error Resume Next
        If Len(well_name) = 0 Then
            Me.Text = sofe_name
            jing_xie.Enabled = False
            jing_shen.Enabled = False
            zy_selt.Enabled = False
            guan_zhu1.Enabled = False
            gongkuan_ipt.Enabled = False
            new_se_well.Enabled = False
            saveas.Enabled = False
            ans_press.Enabled = False
            Menu_mdc.Enabled = False
            jing_xie.Enabled = False
            jing_shen.Enabled = False
            zy_selt.Enabled = False
            guan_zhu1.Enabled = False
            gongkuan_ipt.Visible = True
            ans_press.Visible = True
            'jygjmn.Visible = False
            '---------------------------------
        Else
            Me.Text = sofe_name & "_" & well_name
            jing_xie.Enabled = True   '井斜
            jing_shen.Enabled = True  '井身结构
            zy_selt.Enabled = True    '作业
            Menu_mdc.Enabled = True   '目的层
            If Len(zuoye_name) = 0 Then
                guan_zhu1.Enabled = False
                gongkuan_ipt.Enabled = False
                ans_press.Enabled = False
            Else
                guan_zhu1.Enabled = True
                gongkuan_ipt.Enabled = True
                ans_press.Enabled = True
            End If
            new_se_well.Enabled = True  '油气井基本参数
            saveas.Enabled = True        '另存为
        End If
        Me.Refresh()
    End Sub
    '**********************************************************************************
    '文件管理->新建
    '**********************************************************************************
    Public Sub newfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles newfile.Click
        new_well.ShowDialog()
        new_well.Dispose()
        If Not well_name = "" Then
            'zct_TSM_Activated(Me, New System.EventArgs())
            Me.Text = sofe_name & "_" & well_name
            Call creat_tables()
            Call my_refresh()
        End If
    End Sub
    '**********************************************************************************
    '文件管理->打开
    '**********************************************************************************
    Public Sub open_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles open.Click
        Dim i As Integer
        Dim j As Integer
        Dim curpath As String
        curpath = My.Application.Info.DirectoryPath
        '下面为分离油井名称
        'On Error GoTo errhandler
        CommonDialog1Open.InitialDirectory = My.Application.Info.DirectoryPath & "data\"
        CommonDialog1Open.Filter = "Access(*.mdb)|*.mdb"
        CommonDialog1Open.ShowDialog()
        use_dbname = CommonDialog1Open.FileName '获得井名路径
        i = InStrRev(use_dbname, ".")
        j = InStrRev(use_dbname, "\")
        If i - j - 1 > 0 Then
            well_name = Mid(use_dbname, j + 1, i - j - 1)
            zuoye_name = ""
            ChDir(curpath)
            If check_file() = True Then
                Me.Text = sofe_name & "_" & well_name
                '给全局变量userDbPath赋值，李20180203增加
                userDbPath = Mid(use_dbname, 1, j - 1)
            Else
                Me.Text = sofe_name
                well_name = ""
                zuoye_name = ""
                msg_prompt = "你所打开的数据文件不是本系统设定的数据！请重新选择打开文件。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                Exit Sub
            End If
            CommonDialog1Open.Dispose()
            Call creat_tables() '由于软件的升级，当打开旧的数据时，需要往里边添加新的数据表
            Call user_database_update() '由于软件升级，数据结构改动，用程序自动升级相应数据表的数据结构。
            Call my_refresh()   '使能主菜单中只关于油井的菜单项
        Else
            CommonDialog1Open.Dispose()
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        ChDir(curpath)
    End Sub
    '**********************************************************************************
    '文件管理->另存为
    '**********************************************************************************
    Public Sub saveas_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles saveas.Click
        If Not well_name = "" Then
            MainMenu1.Enabled = False
            save_as.TopLevel = False
            save_as.Parent = Panel1
            'save_as.WindowState = FormWindowState.Maximized
            save_as.Left = Panel1.Left + 0.5 * (Panel1.Width - save_as.Width)
            save_as.Top = Panel1.Top + 0.5 * (Panel1.Height - save_as.Height)
            save_as.Show()
            'Call my_refresh()
        Else
            msg_prompt = "没有打开的文件！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '***************************************************************************
    '文件管理->退出
    '****************************************************************************
    Public Sub tuichu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tuichu.Click
        msg_prompt = "确定要退出" & sofe_name & "?"
        msg_buttons = 4 + 32
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        If msg_return = 6 Then
            End
        End If
    End Sub
    '**********************************************************************************
    '数据输入与管理->油气井基本参数
    '**********************************************************************************
    Public Sub new_se_well_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles new_se_well.Click
        MainMenu1.Enabled = False
        well_manage.TopLevel = False
        well_manage.Parent = Panel1
        'well_manage.WindowState = FormWindowState.Maximized
        well_manage.Left = Panel1.Left + 0.5 * (Panel1.Width - well_manage.Width)
        well_manage.Top = Panel1.Top + 0.5 * (Panel1.Height - well_manage.Height)
        well_manage.Show()
        'well_manage.ShowDialog()
        'well_manage.Dispose()
    End Sub
    '**********************************************************************************
    '数据输入与管理->井眼轨迹
    '**********************************************************************************
    Public Sub jing_xie_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles jing_xie.Click
        If yqj_ok() = True Then
            jing_shen.Enabled = False
            MainMenu1.Enabled = False
            frmipt_testwell.TopLevel = False
            frmipt_testwell.Parent = Panel1
            frmipt_testwell.WindowState = FormWindowState.Maximized
            frmipt_testwell.Left = Panel1.Left + 0.5 * (Panel1.Width - frmipt_testwell.Width)
            frmipt_testwell.Top = Panel1.Top + 0.5 * (Panel1.Height - frmipt_testwell.Height)
            frmipt_testwell.Show()
            jing_shen.Enabled = True
        Else
            msg_prompt = "请输入油气井基本数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '**********************************************************************************
    '数据输入与管理->井身结构
    '**********************************************************************************
    Public Sub jing_shen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles jing_shen.Click
        If yqj_ok() = True Then '判断油气井表中是否有数据
            MainMenu1.Enabled = False
            frmipt_wellstruct.TopLevel = False
            frmipt_wellstruct.Parent = Panel1
            frmipt_wellstruct.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
            frmipt_wellstruct.WindowState = FormWindowState.Maximized
            frmipt_wellstruct.Show()
        Else
            msg_prompt = "请输入油气井基本数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '**********************************************************************************
    '数据输入与管理->作业名称
    '**********************************************************************************
    Public Sub zy_selt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zy_selt.Click
        Dim sd As Short
        If well_name <> "" Then
            MainMenu1.Enabled = False
            selt_zuoye.TopLevel = False
            selt_zuoye.Parent = Panel1
            'selt_zuoye.WindowState = FormWindowState.Maximized
            selt_zuoye.Left = Panel1.Left + 0.5 * (Panel1.Width - selt_zuoye.Width)
            selt_zuoye.Top = Panel1.Top + 0.5 * (Panel1.Height - selt_zuoye.Height)
            selt_zuoye.Show()
            'selt_zuoye.ShowDialog()
            'selt_zuoye.Dispose()
        Else
            sd = MsgBox("请打开油气井文件。", 4 + 32, "注意")
        End If
    End Sub
    '**********************************************************************************
    '数据输入与管理->作业管柱组合
    '**********************************************************************************
    Public Sub guan_zhu1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles guan_zhu1.Click
        If yqj_ok() = True Then
            MainMenu1.Enabled = False
            ipt_gzhu.TopLevel = False
            ipt_gzhu.Parent = Panel1
            ipt_gzhu.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
            ipt_gzhu.WindowState = FormWindowState.Maximized
            ipt_gzhu.Show()
        Else
            msg_prompt = "请输入油气井基本数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '**********************************************************************************
    '数据输入与管理->作业工况参数
    '**********************************************************************************
    Public Sub gongkuan_ipt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles gongkuan_ipt.Click
        Dim SQL_command As String
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        If yqj_ok() = True Then
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select top 1 * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            If Not RECreader.Read Then
                RECreader.Close()
                cn_userdb.Close()
                msg_prompt = "请先输入管柱结构数据！"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                RECreader.Close()
                cn_userdb.Close()
            Else
                RECreader.Close()
                cn_userdb.Close()
                MainMenu1.Enabled = False
                ipt_gkuang.TopLevel = False
                ipt_gkuang.Parent = Panel1
                ipt_gkuang.WindowState = FormWindowState.Normal
                ipt_gkuang.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
                ipt_gkuang.WindowState = FormWindowState.Maximized
                ipt_gkuang.Show()
            End If
        Else
            msg_prompt = "请输入油气井基本数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '***************************************************************************
    '数据输入与管理->目的层参数    李润洲2025年6月15日增加
    '****************************************************************************
    Private Sub Menu_mdc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menu_mdc.Click
        Dim SQL_command As String
        If yqj_ok() = True Then
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                SQL_command = "select top 3 *  from 测井数据表 where  井号='" & well_name & "'"
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        If RECreader.Read Then
                            MainMenu1.Enabled = False
                            ipt_mdceng.TopLevel = False
                            ipt_mdceng.Parent = Panel1
                            ipt_mdceng.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
                            ipt_mdceng.WindowState = FormWindowState.Maximized
                            ipt_mdceng.Show()
                        Else
                            msg_prompt = "请先导入井眼轨迹数据。"
                            msg_buttons = 0 + 48
                            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                        End If
                    End Using
                End Using
            End Using
        Else
            msg_prompt = "请输入油气井基本数据。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End If
    End Sub
    '**********************************************************************************
    '分析计算->压力场分析
    '**********************************************************************************
    Public Sub ans_press_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ans_press.Click
        TSM_ver_switch = 1   '管柱力学分析软件版本控制开关。等于1时，垂直井节点法，井斜、方位均等于0，不判断、不读取井斜数据
        If (gzlxfx_data_ok() = True) Then
            MainMenu1.Enabled = False
            Frm_PRS.TopLevel = False
            Frm_PRS.Parent = Me.Panel1
            Frm_PRS.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Frm_PRS.WindowState = FormWindowState.Maximized
            Frm_PRS.Show()
        End If
    End Sub
    '***************************************************************************
    ' 基础数据管理->钻杆
    '****************************************************************************
    Public Sub zuang_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles zuang.Click
        MainMenu1.Enabled = False
        frmdb_drillpipe.TopLevel = False
        frmdb_drillpipe.Parent = Me.Panel1
        frmdb_drillpipe.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_drillpipe.WindowState = FormWindowState.Maximized
        frmdb_drillpipe.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->套管
    '**********************************************************************************
    Public Sub taog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles taog.Click
        MainMenu1.Enabled = False
        frmdb_casing.TopLevel = False
        frmdb_casing.Parent = Panel1
        frmdb_casing.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_casing.WindowState = FormWindowState.Maximized
        frmdb_casing.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->油管
    '**********************************************************************************
    Public Sub youg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles youg.Click
        MainMenu1.Enabled = False
        frmdb_youguan.TopLevel = False
        frmdb_youguan.Parent = Panel1
        frmdb_youguan.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_youguan.WindowState = FormWindowState.Maximized
        frmdb_youguan.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->封隔器
    '**********************************************************************************
    Public Sub fenggq_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fenggq.Click
        MainMenu1.Enabled = False
        frmdb_fgq.TopLevel = False
        frmdb_fgq.Parent = Panel1
        frmdb_fgq.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_fgq.WindowState = FormWindowState.Maximized
        frmdb_fgq.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->锚定工具
    '**********************************************************************************
    Public Sub msgj_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgj.Click
        MainMenu1.Enabled = False
        frmdb_mdgj.TopLevel = False
        frmdb_mdgj.Parent = Panel1
        frmdb_mdgj.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_mdgj.WindowState = FormWindowState.Maximized
        frmdb_mdgj.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->开关工具
    '**********************************************************************************
    Public Sub openelem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles openelem.Click
        MainMenu1.Enabled = False
        frmdb_kgyj.TopLevel = False
        frmdb_kgyj.Parent = Panel1
        frmdb_kgyj.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_kgyj.WindowState = FormWindowState.Maximized
        frmdb_kgyj.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->节流工具
    '**********************************************************************************
    Public Sub jieliuyuanian_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles jieliuyuanian.Click
        MainMenu1.Enabled = False
        frmdb_jlyj.TopLevel = False
        frmdb_jlyj.Parent = Panel1
        frmdb_jlyj.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_jlyj.WindowState = FormWindowState.Maximized
        frmdb_jlyj.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理->伸缩管
    '**********************************************************************************
    Public Sub flexele_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexele.Click
        MainMenu1.Enabled = False
        frmdb_ssyj.TopLevel = False
        frmdb_ssyj.Parent = Panel1
        frmdb_ssyj.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmdb_ssyj.WindowState = FormWindowState.Maximized
        frmdb_ssyj.Show()
    End Sub
    '**********************************************************************************
    '基础数据管理-管柱工具图库
    '**********************************************************************************
    Public Sub L_20002_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'frmdb_LGuanZhuTu.ShowDialog()
    End Sub
    Public Sub lxfx_smpest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'Dialog.ShowDialog()
    End Sub
    '***************************************************************************
    '帮助->如何使用
    '****************************************************************************
    Public Sub help_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        System.Windows.Forms.SendKeys.Send("{F1}")
    End Sub
    '**********************************************************************************
    '帮助->关于...
    '**********************************************************************************
    Public Sub about_soft_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        MainMenu1.Enabled = False
        frmAbout1.TopLevel = False
        frmAbout1.Parent = Panel1
        'frmAbout1.WindowState = FormWindowState.Maximized
        frmAbout1.Left = Panel1.Left + 0.5 * (Panel1.Width - frmAbout1.Width)
        frmAbout1.Top = Panel1.Top + 0.5 * (Panel1.Height - frmAbout1.Height)
        frmAbout1.Show()
        'frmAbout1.ShowDialog()
        'frmAbout1.Dispose()
    End Sub
    '***************************************************************************
    '主菜单上退出
    '****************************************************************************
    Public Sub tuichu2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tuichu2.Click
        Call tuichu_Click(tuichu, New System.EventArgs())
    End Sub

    Private Sub AddLabelToPanel()
        ' 1. 创建临时容器（用于承载阴影Label，避免被主Label穿透）
        Dim shadowContainer As New Panel()
        With shadowContainer
            .Parent = Panel1 ' 父容器是主Panel
            .Location = New Point(50, 50) ' 整体位置
            .Size = New Size(200, 30) ' 与文本大小匹配
            .BackColor = Color.Transparent ' 容器透明（显示主Panel背景）
            .BorderStyle = BorderStyle.None ' 无边框
        End With
        Panel1.Controls.Add(shadowContainer)

        ' 2. 阴影Label（放在临时容器中）
        Dim shadowLbl As New Label()
        With shadowLbl
            .Parent = shadowContainer ' 父容器是临时Panel
            .Text = "带阴影的文本"
            .ForeColor = Color.DimGray ' 阴影颜色
            .Location = New Point(5, 5) ' 相对于容器偏移（形成阴影）
            .Dock = DockStyle.Fill ' 填满容器
            .Font = New Font("微软雅黑", 14, FontStyle.Bold)
            .BackColor = Color.Transparent ' 透明（显示临时容器背景）
            .BorderStyle = BorderStyle.None
        End With
        shadowContainer.Controls.Add(shadowLbl)
        ' 3. 主文本Label（直接放在主Panel，覆盖临时容器）
        Dim mainLbl As New Label()
        With mainLbl
            .Parent = Panel1 ' 父容器是主Panel
            .Text = "带阴影的文本"
            .ForeColor = Color.OrangeRed  ' 主文本颜色
            .Location = New Point(50, 50) ' 与临时容器位置重合
            .Size = New Size(200, 30) ' 覆盖临时容器
            .Font = New Font("微软雅黑", 14, FontStyle.Bold)
            .BackColor = Color.Transparent ' 透明（显示下方的临时容器及阴影）
            .BorderStyle = BorderStyle.None
        End With
        Panel1.Controls.Add(mainLbl)
    End Sub
    '***************************************************************************
    '释放主窗体Panel使用的图片资源   李润洲2025年7月27增加
    '****************************************************************************
    Private Sub zct_main_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Panel1.BackgroundImage IsNot Nothing Then
            Panel1.BackgroundImage.Dispose()
        End If
    End Sub
    '******************************************************************************************************************
    '（1）显示带阴影的软件名称，（2）根据Panel大小动态调整显示的字号 （3）可设置字符间距  李润洲2025年7月27增加
    '*******************************************************************************************************************
    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        Dim g As Graphics = e.Graphics
        Dim shadowColor As Color = Color.FromArgb(150, Color.Yellow) ' 半透明深灰阴影
        Dim textColor As Color = Color.Red        ' 字体颜色
        Dim textToDraw As String = sofe_name
        Dim letterSpacing As Single = -10

        If sofe_name Is Nothing Then Exit Sub
        ' 优化绘图质量
        g.SmoothingMode = SmoothingMode.HighQuality
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit  '文字抗锯齿
        ' 定义颜色
        '-----------------------------------------------------
        '主软件名
        ' 字间距（可根据需要调整，单位：像素）
        ' 1. 计算目标宽度（Panel宽度的三分之二）
        Dim targetWidth As Single = Panel1.ClientSize.Width * 3 / 5
        ' 最小宽度限制（避免Panel过小导致计算异常）
        If targetWidth < 50 Then targetWidth = 50
        ' 2. 动态计算合适的字体大小
        Dim fontName As String = "微软雅黑"
        For Each family As FontFamily In FontFamily.Families
            If family.Name.Length >= 4 Then
                If family.Name.Substring(0, 4) = "华光行楷" Or family.Name.Substring(0, 4) = "华文行楷" Then
                    fontName = family.Name
                    Exit For
                End If
            End If
        Next
        Dim fontSize As Single = CalculateFontSizeWithSpace(g, textToDraw, fontName, FontStyle.Regular, targetWidth, letterSpacing)
        ' 3. 创建字体（限制最大/最小字号，避免过大或过小）
        Dim maxFontSize As Single = 100 ' 最大字号
        Dim minFontSize As Single = 12  ' 最小字号
        fontSize = Math.Min(Math.Max(fontSize, minFontSize), maxFontSize)
        Dim font As New Font(fontName, fontSize, FontStyle.Bold)
        ' 4.计算文本总宽度（含字间距）
        Dim totalWidth As Single = MeasureTextWithSpacing(g, textToDraw, font, letterSpacing)
        ' 5. 计算文本位置（水平居中，垂直居上方5分之一处）
        Dim startX As Single = (Panel1.ClientSize.Width - totalWidth) / 2
        Dim startY As Single = Panel1.ClientSize.Height / 5
        ' 6. 绘制带阴影的文本（逐个字符绘制，支持字间距），
        DrawTextWithSpacingAndShadow(g, textToDraw, font, New PointF(startX, startY), letterSpacing, shadowColor, textColor) ' 传入颜色参数
        '-------------------------------------------------------------
        '软件版本号，不设字间距，字号是主软件名字号的80%
        textToDraw = sofe_ver
        ' 字号，主软件名字号的1/2
        Dim ver_fontSize = fontSize * 1 / 2
        Dim ver_font As New Font(fontName, ver_fontSize, FontStyle.Bold)
        '计算文本位置
        Dim zhu_dz_size As SizeF = g.MeasureString("字", font)
        Dim ver_Text_Size As SizeF = g.MeasureString(textToDraw, font)
        ' 主软件名下半行，居中
        Dim verLocation As New PointF(startX + totalWidth / 2 - ver_Text_Size.Width / 2, startY + zhu_dz_size.Height * 1.2)
        '绘制阴影
        Using shadowBrush As New SolidBrush(Color.FromArgb(150, Color.Black))
            g.DrawString(textToDraw, ver_font, shadowBrush, verLocation.X + 4, verLocation.Y + 4)
        End Using
        '  绘制主文本
        Using mainBrush As New SolidBrush(textColor)
            g.DrawString(textToDraw, ver_font, mainBrush, verLocation)
        End Using
        '-------------------------------------------------------------
        '使用单位，字间距，字号是主软件名字号的0.9%
        textToDraw = sofe_danwei
        ' 字间距（可根据需要调整，单位：像素）
        letterSpacing = -5
        Dim dw_fontSize As Single = fontSize * 0.9
        Dim dw_font As New Font(fontName, dw_fontSize, FontStyle.Bold) '.Regular
        ' 计算文本总宽度（含字间距）
        Dim dw_totalWidth As Single = MeasureTextWithSpacing(g, textToDraw, dw_font, letterSpacing)
        '单位文字过多
        If dw_totalWidth > totalWidth Then
            Dim dw_targetWidth As Single = Panel1.ClientSize.Width * 0.9
            ' 最小宽度限制（避免Panel过小导致计算异常）
            If targetWidth < 50 Then targetWidth = 50
            ' 2. 动态计算合适的字体大小
            ' 备选：隶书 "SimLi"（宋体-隶书）、"STLiti"（华文隶书）,"微软雅黑","黑体"
            dw_fontSize = CalculateFontSizeWithSpace(g, textToDraw, fontName, FontStyle.Regular, targetWidth, letterSpacing)
            ' 创建字体（限制最大/最小字号，避免过大或过小）
            maxFontSize = 100 ' 最大字号
            minFontSize = 12  ' 最小字号
            dw_fontSize = Math.Min(Math.Max(dw_fontSize, minFontSize), maxFontSize)
            dw_font.Dispose()
            dw_font = New Font(fontName, dw_fontSize, FontStyle.Bold)
            dw_totalWidth = MeasureTextWithSpacing(g, textToDraw, dw_font, letterSpacing)
        End If
        ' 计算文本位置,行距1.5倍
        Dim dw_startX As Single = startX + totalWidth / 2 - dw_totalWidth / 2
        Dim dw_startY As Single = verLocation.Y + ver_Text_Size.Height * 1.5
        If Panel1.ClientSize.Height * 0.65 > dw_startY Then
            dw_startY = Panel1.ClientSize.Height * 0.65
        End If
        '绘制带阴影的文本（逐个字符绘制，支持字间距），函数内设置显示颜色
        DrawTextWithSpacingAndShadow(g, textToDraw, dw_font, New PointF(dw_startX, dw_startY), letterSpacing, shadowColor, textColor) ' 传入颜色参数
        '-------------------------------------------------------------
        '时间，字间距=主软件名字间距，字号是主软件名字号的70%，行距1.2
        textToDraw = sofe_time
        ' 字间距（可根据需要调整，单位：像素）
        letterSpacing = 0
        Dim sj_fontSize As Single = fontSize * 0.5
        Dim sj_font As New Font(fontName, sj_fontSize, FontStyle.Bold)
        ' 计算文本总宽度（含字间距）
        Dim sj_totalWidth As Single = MeasureTextWithSpacing(g, textToDraw, sj_font, letterSpacing)
        ' 计算文本位置,单位下变空半行
        Dim sj_startX As Single = startX + totalWidth / 2 - sj_totalWidth / 2
        Dim sj_startY As Single = dw_startY + zhu_dz_size.Height * 1.1
        '绘制带阴影的文本（逐个字符绘制，支持字间距），函数内设置显示颜色
        DrawTextWithSpacingAndShadow(g, textToDraw, sj_font, New PointF(sj_startX, sj_startY), letterSpacing, shadowColor, textColor) ' 传入颜色参数
        font.Dispose()
        ver_font.Dispose()
        dw_font.Dispose()
        sj_font.Dispose()
    End Sub
    '******************************************************************************************************************
    '' 计算包含字间距的文本总宽度  李润洲2025年7月27日增加
    '*******************************************************************************************************************
    Private Function MeasureTextWithSpacing(ByVal g As Graphics, ByVal text As String, ByVal font As Font, ByVal spacing As Single) As Single
        If text.Length = 0 Then Return 0
        Dim totalWidth As Single = 0
        ' 测量每个字符的宽度
        For Each c As Char In text
            totalWidth += g.MeasureString(c.ToString(), font).Width
        Next
        ' 加上字间距（字符数-1个间隔）
        totalWidth += spacing * (text.Length - 1)
        Return totalWidth
    End Function
    '******************************************************************************************************************
    ' 计算适合目标宽度的字号（包含字间距）   李润洲2025年7月27日增加
    '*******************************************************************************************************************
    Private Function CalculateFontSizeWithSpace(ByVal g As Graphics, ByVal text As String, ByVal fontName As String, _
                                      ByVal style As FontStyle, ByVal targetWidth As Single, ByVal spacing As Single) As Single
        Dim testSize As Single = 20 ' 测试字号
        Dim testFont As New Font(fontName, testSize, style)
        ' 测量测试字号下的总宽度（含字间距）
        Dim measuredWidth As Single = MeasureTextWithSpacing(g, text, testFont, spacing)
        ' 按比例计算目标字号
        Dim targetSize As Single = If(measuredWidth = 0, testSize, (targetWidth / measuredWidth) * testSize)
        testFont.Dispose()
        Return targetSize
    End Function
    '******************************************************************************************************************
    ' 逐个字符绘制文本（支持字间距和阴影）,带颜色参数的绘制函数（核心修改）   李润洲2025年7月27日增加
    '*******************************************************************************************************************
    Private Sub DrawTextWithSpacingAndShadow(ByVal g As Graphics, ByVal text As String, ByVal font As Font, _
                                            ByVal startPos As PointF, ByVal spacing As Single, _
                                            ByVal shadowColor As Color, ByVal textColor As Color) ' 新增颜色参数
        Dim currentX As Single = startPos.X
        Dim currentY As Single = startPos.Y
        ' 绘制阴影（使用传入的阴影颜色）
        For Each c As Char In text
            g.DrawString(c.ToString(), font, New SolidBrush(shadowColor), _
                        currentX + 4, currentY + 4)
            ' 计算字符宽度
            Dim charSize As SizeF = g.MeasureString(c.ToString(), font)
            currentX += charSize.Width + spacing
        Next
        ' 绘制主文本（使用传入的文本颜色）
        currentX = startPos.X
        For Each c As Char In text
            g.DrawString(c.ToString(), font, New SolidBrush(textColor), currentX, currentY)
            Dim charSize As SizeF = g.MeasureString(c.ToString(), font)
            currentX += charSize.Width + spacing
        Next
    End Sub
    '******************************************************************************************************************
    ' 计算使文本宽度等于目标宽度的字体大小，不带字间距   李润洲2025年7月27日增加
    '*******************************************************************************************************************
    Private Function CalculateFontSize(ByVal g As Graphics, ByVal text As String, ByVal fontFamily As String, ByVal style As FontStyle, ByVal targetWidth As Single) As Single
        ' 初始字号（用于计算比例）
        Dim testSize As Single = 20
        Dim testFont As New Font(fontFamily, testSize, style)
        ' 测量初始字号下的文本宽度
        Dim measuredWidth As Single = g.MeasureString(text, testFont).Width

        ' 根据比例计算目标字号（目标宽度/实际宽度 = 目标字号/测试字号）
        Dim targetSize As Single = (targetWidth / measuredWidth) * testSize
        testFont.Dispose()
        Return targetSize
    End Function
    Private Sub zct_main_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Panel1.Invalidate() ' 触发Panel重绘
    End Sub
End Class