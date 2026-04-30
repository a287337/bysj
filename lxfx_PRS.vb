Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_lxfx_prs
    '***************************************************************************************************************************
    '                                                   用节点法进行（定长度）管柱力学分析计算子程序
    '                                                                                            秦彦斌 2022年2月27日最后整理
    ' 主要程序升级记事：
    '    2017年5月-2017年12月
    '           实现管柱中锚定工具参与计算。实现了多封隔器、多锚定工具的多点约束管柱力学分析。重新编写了坐封后工况受力及变形计算
    '       部分程序的算法。提出了约束表的概念，遍历约束表，在各分段间加或减轴向力，让各约束点回到坐封工况时的位置，进而计算出各
    '       节点的受力及变形。
    '
    '   2018年1月-2018年6月,实现管柱中钻杆参与计算。
    '   2018年8月13日,(1)统一、完善坐封工况等效轴力计算部分;(2)解决了投球打压坐封时，若管柱发生屈曲，程序异常退出的问题。
    '   2018年8月31日-2018年10月31日,
    '           实现有伸缩管管柱，伸缩管动作判断，管柱轴力、变形分析。在用户数据库中，为记录各工况下，各个伸缩管的拉伸、压缩状态，
    '       建立“工况_伸缩管状态”数据表。
    '   2019年1月31日,对于迭代计算附加轴向力时收敛性差的提示,修改了最大迭代次数及退出提示。
    '   2019年11月11日,解决了对于非坐封工况当有伸缩管且伸缩管剪销状态未剪断时轴力计算错误BUG
    '   2020年3月1日,计算中若出现工况序号不连续，给出提示.
    '   2020年3月2日,管柱力学分析软件开发工具迁移到VS2008基本完成。
    '   2020年12月20日-21日,弹性模量、泊松比、热胀系数均不再使用固定常量，采用节点数据jd_base结构中的相应变量。
    '   2021年5月12日,在首次往“节点计算参数表”中插入数据时，给其它所有字段赋初值。
    '   2021年8月23日-2021年8月24日,(1)修正“考虑流体摩阻修正悬持等效轴力”部分代码;(2)修正若有定位且不是坐封工况的等效轴力计算算法。
    '   2021年10月20日-2021年10月24日，理清了不同工况中封隔器、锚定工具可能的设置情况，实现在力学分析计算时识别并分别处理。
    '   2022年2月10日-27日，实现了非坐封工况“封隔器正确归位”，“伸缩管正确动作”。
    '   2024年3月15日-17日，实现了屈曲临界载荷分工况计算。
    '
    '子程序功能
    ' 输入参数：
    ' dlt_L:            划分节点时节点的最大间距
    '                   
    '
    '编程思路：
    '       由于用线性插值法计算当前井任意深度in_depth处的井斜角、方位角、垂深子程序cal_jx_fw_cs需要频繁调用，故改该子程序原来
    ' 频繁读取井斜数据表的方法，利用ADO.NET提供的DataTable，设置全局变量Testwell_Table，一次填充，多次使用，避免频繁读取数据库。
    '       为此，在本程序开始运行时，对于非水平井，需把测斜数据读到“Testwell_Table”表中，以便子程序cal_jx_fw_cs正常运行。
    '
    '   在判断数据完整、合法的基础上，对每一工况
    '   1、对管柱进行分段，分段结果保存在“节点情况表”中；
    '   2、计算各工况情况下各分段的液压压强，计算结果填写到“节点计算参数表”中；
    '   3、分工况进行管柱力学分析计算，计算结果填写到“节点计算参数表”中。
    '***************************************************************************************************************************
    Public Function lxfx_prs(ByVal dlt_L As Double) As Boolean
        Dim SQL_command As String
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim ad As New System.Data.OleDb.OleDbDataAdapter
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim gz_shsg_row As DataRow
        Dim gk_row As DataRow
        Dim gk_mdgjzt_row As DataRow
        Dim gk_fgqzt_row As DataRow
        Dim jdqk_row As DataRow
        Dim jdsj_base() As jd_base
        Dim jdsj_cacu(,) As jd_cacu
        Dim total_gk As Integer     '总工况数
        Dim PB1_cout As Integer     '进度条计数器
        Dim total_jd As Integer     '总节点数
        Dim gk_gkxh As Integer      '当前工况序号
        Dim gk_gkxh_zd As Integer   '当前工况序号字段值
        Dim zf_gkxh As Integer      '座封工况序号
        Dim i As Integer
        Dim j As Integer
        Dim tmp_i As Integer
        Dim tmp_j As Integer
        Dim tmp_k As Integer
        Dim tmp_l As Integer     '临时整形变量
        Dim tmp_m As Integer     '临时整形变量
        Dim tmp_n As Integer     '临时整形变量
        Dim gm As Double         '完井液密度
        Dim TMP_so As Double     '下钻时地面温度
        Dim Ffi As Double        '单元摩擦力
        Dim z_Ffi As Double      '总摩擦力
        Dim pi As Double         '圆周率
        Dim bx_wd() As Double
        Dim bx_zl() As Double
        Dim bx_gz() As Double
        Dim bx_lx() As Double
        Dim zf_wd_bx() As Double   '坐封好时节点的温度变形
        Dim zf_zl_bx() As Double   '坐封好时节点的轴力变形
        Dim zf_gz_bx() As Double   '坐封好时节点的臌胀变形 
        Dim zf_lx_bx() As Double   '坐封好时节点的螺旋变形
        Dim zf_zh_bx() As Double   '坐封好时节点的综合变形
        Dim qe_arr() As Double   '管柱在井液中的线重
        Dim total_length As Double              '管柱总长度
        Dim shsg_Count As Integer               '伸缩管数量
        Dim shsg_shsdzzh() As Double            '伸缩管伸缩动作载荷
        Dim shsg_shsdzyl() As Double            '伸缩管伸缩动作压力
        Dim shsg_shsxch() As Double             '伸缩管伸缩行程m
        Dim shsg_qsdchd() As Double             '伸缩管全缩短长度
        Dim shsg_shszht() As String             '伸缩管剪销状态
        Dim shsg_yjxh() As Integer              '伸缩管元件序号
        Dim shsg_yjmch() As String              '伸缩管元件名称
        Dim shsg_xrchd() As Double              '伸缩管下入长度
        Dim shsg_shschd() As Double             '伸缩管的伸缩长度
        Dim shsg_gkchshchd() As Double          '伸缩管本工况初始长度
        Dim shsg_shslzg() As String             '伸缩管伸缩状态 情况
        Dim shsg_shyushsl As Double             '伸缩管剩余伸缩量m
        Dim qe_weifu As Boolean                 '若算出的qe为负值，qe_weifu=true， 各工况都用，不得借用
        Dim qe_weifumsg As String               'msgbox函数显示文本内容， 各工况都用，不得借用
        Dim jixu_ditui As Boolean               '考虑摩擦力递推轴向力标志，若自锁或上提下放载荷被摩擦力克服完，程序不再递推，jixu_ditui=false，否则，继续递推jixu_ditui=true。
        Dim sf_zisuo As Boolean                 '考虑摩擦力递推轴向力时，标记是否发生自锁，若自锁，sf_zisuo=true，若没自锁，sf_zisuo=false
        Dim xd_zf_jd As Boolean                 '标记伸缩管销钉是否在坐封工况被剪断
        '常数赋值     
        pi = 3.14159265358979          '圆周率
        lxfx_prs = True            '函数返回赋初值
        'tam = 206842.72#              'MPa  钢材弹性模量,考虑到非碳钢管材，直接用管材数据，此变量不再使用
        'psb = 0.3                     '钢材泊松比,考虑到非碳钢管材，直接用管材数据，此变量不再使用
        qe_weifu = False
        qe_weifumsg = ""
        jixu_ditui = True              '考虑摩擦力递推轴向力标志，若自锁或上提下放载荷被摩擦力克服完，程序不再递推，jixu_ditui=false，否则，继续递推jixu_ditui=true。
        sf_zisuo = False
        xd_zf_jd = False
        shsg_Count = 1
        Ffi = 0.0
        z_Ffi = 0.0
        ReDim shsg_shsdzzh(shsg_Count)
        ReDim shsg_shsdzyl(shsg_Count)
        ReDim shsg_shsxch(shsg_Count)
        ReDim shsg_qsdchd(shsg_Count)
        ReDim shsg_shszht(shsg_Count)
        ReDim shsg_yjxh(shsg_Count)
        ReDim shsg_yjmch(shsg_Count)
        ReDim shsg_xrchd(shsg_Count)
        ReDim shsg_shschd(shsg_Count)
        ReDim shsg_gkchshchd(shsg_Count)
        ReDim shsg_shslzg(shsg_Count)
        For i = 0 To shsg_Count - 1
            shsg_shsdzzh(i) = 0
            shsg_shsdzyl(i) = 0
            shsg_shsxch(i) = 0
            shsg_qsdchd(i) = 0
            shsg_shszht(i) = ""
            shsg_yjxh(i) = 0
            shsg_yjmch(i) = ""
            shsg_xrchd(i) = 0
            shsg_shschd(i) = 0
            shsg_gkchshchd(i) = 0
            shsg_shslzg(i) = ""
        Next i
        shsg_shyushsl = 0.0         '伸缩管剩余伸缩量m
        '简单判断数据完整性、合法性--各数据表里有数据
        If (gzlxfx_data_ok() = False) Then
            msg_prompt = "管柱力学分析计算所需数据不完整，请输入有关数据后再算！"
            lxfx_prs = False
            Exit Function
        End If

        '计算前先重新生成节点计算参数表
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = dbSchema.Select("TABLE_NAME='节点计算参数表'")
        If foundRows.Length <> 0 Then
            SQL_command = "DROP TABLE 节点计算参数表"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        foundRows = dbSchema.Select("TABLE_NAME='节点情况表'")
        If foundRows.Length <> 0 Then
            SQL_command = "DROP TABLE 节点情况表"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        foundRows = dbSchema.Select("TABLE_NAME='工况_伸缩管状态'")
        If foundRows.Length <> 0 Then
            SQL_command = "DROP TABLE 工况_伸缩管状态"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        dbSchema.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()

        Call database_creat(12, 2) '建立数据表  节点情况表
        Call database_creat(13, 2) '建立数据表  节点计算参数表
        Call database_creat(34, 2) '建立数据表  工况_伸缩管状态

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '***************************************************************************************************************************
        '0、把测斜数据读到“Testwell_Table”表中
        '***************************************************************************************************************************
        If Not TSM_ver_switch = 1 Then
            SQL_command = "select  * from 测井数据表 where 井号='" & well_name & "' order by [井  深(m)]"
            ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
            Testwell_Table.Clear()
            ad.Fill(Testwell_Table)
            ad.Dispose()
        End If
        '***************************************************************************************************************************
        '读取工况参数
        '***************************************************************************************************************************
        SQL_command = "select * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_Table.Clear()
        ad.Fill(gk_Table)
        ad.Dispose()
        total_gk = dst.Tables("gk_Table").Rows.Count
        '***************************************************************************************************************************
        '读取伸缩元件数据，顺便进行数据合理性判断
        '    Dim shsg_Count As Integer                          '伸缩管数量
        '    Dim shsg_shsdzzh() As Double                       '伸缩管伸缩动作载荷
        '    Dim shsg_shsdzyl() As Double                       '伸缩管伸缩动作压力
        '    Dim shsg_shsxch() As Double                        '伸缩管伸缩行程m
        '    Dim shsg_qsdchd() As Double                        '伸缩管全缩短长度
        '    Dim shsg_shszht() As String                        '伸缩管剪销状态
        '    Dim shsg_yjxh() As integer                         '伸缩管元件序号
        '    Dim shsg_yjmch() As String                         '伸缩管元件名称
        '    Dim shsg_xrchd() As Double                         '伸缩管下入长度
        '    Dim shsg_shschd() As Double                        '伸缩管的伸缩长度
        '    Dim shsg_gkchshchd() As Double                     '伸缩管本工况初始长度
        '    Dim shsg_shslzg() As String                        '伸缩管伸缩状态 情况
        '***************************************************************************************************************************
        shsg_Count = 0
        SQL_command = "select * from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'order by 元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gz_shsg_Table.Clear()
        ad.Fill(gz_shsg_Table)
        ad.Dispose()
        If gz_shsg_Table.Rows.Count > 0 Then
            shsg_Count = gz_shsg_Table.Rows.Count
            ReDim shsg_shsdzzh(shsg_Count)
            ReDim shsg_shsdzyl(shsg_Count)
            ReDim shsg_shsxch(shsg_Count)
            ReDim shsg_qsdchd(shsg_Count)
            ReDim shsg_shszht(shsg_Count)
            ReDim shsg_yjxh(shsg_Count)
            ReDim shsg_yjmch(shsg_Count)
            ReDim shsg_xrchd(shsg_Count)
            ReDim shsg_shschd(shsg_Count)
            ReDim shsg_gkchshchd(shsg_Count)
            ReDim shsg_shslzg(shsg_Count)
            i = 0
            For Each gz_shsg_row In gz_shsg_Table.Select
                shsg_shsdzzh(i) = Val(gz_shsg_row.Item("伸缩动作载荷kN").ToString)
                shsg_shsdzyl(i) = Val(gz_shsg_row.Item("伸缩动作压力MPa").ToString)
                shsg_shsxch(i) = Val(gz_shsg_row.Item("伸缩行程m").ToString)
                shsg_qsdchd(i) = Val(gz_shsg_row.Item("全缩短长度m").ToString)
                shsg_shszht(i) = "未剪断"
                shsg_yjxh(i) = Val(gz_shsg_row.Item("元件序号").ToString)
                shsg_yjmch(i) = gz_shsg_row.Item("元件名称").ToString
                If Val(gz_shsg_row.Item("伸缩动作载荷kN").ToString) < 0 Then
                    msg_prompt = "编号为" & gz_shsg_row.Item("元件序号").ToString & "的伸缩元件数据不合理，其伸缩动作载荷应该大于零，请核对数据后再算！"
                    lxfx_prs = False
                    cn_userdb.Close()
                    cn_userdb.Dispose()
                    Exit Function
                End If
                If Val(gz_shsg_row.Item("伸缩动作压力MPa").ToString) < 0 Then
                    msg_prompt = "编号为" & gz_shsg_row.Item("元件序号").ToString & "的伸缩元件数据不合理，其伸缩动作压力应该大于零，请核对数据后再算！"
                    lxfx_prs = False
                    cn_userdb.Close()
                    cn_userdb.Dispose()
                    Exit Function
                End If
                SQL_command = "select * from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(shsg_yjxh(i))
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                EXECOleDbCommand.Dispose()
                RECreader.Read()
                shsg_xrchd(i) = IIf(RECreader.Item("元件长度m").ToString() = "", 0.0, Val(RECreader.Item("元件长度m").ToString()))
                shsg_shschd(i) = 0.0
                shsg_gkchshchd(i) = shsg_xrchd(i)
                shsg_shslzg(i) = "未伸缩"
                i = i + 1
            Next
        End If
        '***************************************************************************************************************************
        '填写“工况_伸缩管状态”数据表初始值
        'SQL_command = "CREATE TABLE 工况_伸缩管状态(" _
        '        & "作业名称 TEXT(50)    ,工况序号 INTEGER   ,工况名称 TEXT(50)    ,元件序号 INTEGER       ,元件名称 TEXT(50)," _
        '        & " 伸缩动作载荷kN FLOAT,伸缩行程m FLOAT    ,全缩短长度m FLOAT    ,剪切力kN FLOAT ,剪销状态 text(50)," _
        '        & " 伸缩长度m FLOAT     ,伸缩状态 text(50),本工况初始长度m FLOAT,伸缩动作压力MPa FLOAT  ,不伸缩等效轴力kN FLOAT ," _
        '        & " 内压MPa FLOAT       ,外压MPa FLOAT      ,压差剪切力kN         ,等效轴力kN FLOAT ,井号 TEXT(50))"
        'Dim shsg_shsdzyl() As Double    '伸缩管伸缩动作压力
        '***************************************************************************************************************************
        For Each gk_row In gk_Table.Select
            For i = 0 To shsg_Count - 1 Step 1
                SQL_command = "insert into  工况_伸缩管状态 (" _
                    & "作业名称,工况序号,工况名称,元件序号,元件名称," _
                    & "伸缩动作载荷kN,伸缩行程m,全缩短长度m,剪切力kN,剪销状态," _
                    & "伸缩长度m,伸缩状态,本工况初始长度m,伸缩动作压力MPa,不伸缩等效轴力kN," _
                    & "内压MPa,外压MPa,压差剪切力kN,等效轴力kN,井号) values (" _
                    & "'" & zuoye_name & "'," & gk_row.Item("工况序号").ToString & ",'" & gk_row.Item("工况名称").ToString & "'," & CStr(shsg_yjxh(i)) & ",'" & shsg_yjmch(i) & "'," _
                    & CStr(shsg_shsdzzh(i)) & "," & CStr(shsg_shsxch(i)) & "," & CStr(shsg_qsdchd(i)) & ",0,'未剪断'," _
                    & "0,'" & shsg_shslzg(i) & "'," & CStr(shsg_gkchshchd(i)) & "," & CStr(shsg_shsdzyl(i)) & ",0," _
                    & "0,0,0,0,'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            Next i
        Next
        '**********************************************************************************************************************
        '                                                      先找坐封工况
        '思路与原则：
        '   1.所有工况中都没有坐封的封隔器，没有锚定的锚定工具，不需要找坐封工况。否则，需找坐封工况，以便在后续工况中计算效应。
        '   2.坐封工况可能是下面情况之一
        '   坐封(座卡)情况1：(1)有锚定工具，在某一工况下锚定工具均锚定；(2)有封隔器，在某一工况下封隔器均座封(座卡)
        '   坐封(座卡)情况2：(1)无锚定工具；(2)有封隔器，在某一工况下封隔器均座封(座卡)
        '   坐封(座卡)情况3：(1)有锚定工具，在某一工况下锚定工具均锚定；(2)无封隔器
        '
        '工况数据循环一遍，找出座封(坐卡)工况-第一个全部封隔器、锚定工具均坐卡的工况
        '    检查坐封工况各个封隔器是否均坐封。若有多个封隔器，多个封隔器均应在座封(座卡)工况座封(座卡)
        '**********************************************************************************************************************
        '**********************************************************************************************************************
        '读该井选定作业的“工况_锚定元件”数据
        '**********************************************************************************************************************
        SQL_command = "select * from 工况_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号,元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_mdgjzt_Table.Clear()
        ad.Fill(gk_mdgjzt_Table)
        ad.Dispose()
        '**********************************************************************************************************************
        '读该井选定作业的“工况_封隔定位元件”数据
        '**********************************************************************************************************************
        SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 工况序号,元件序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        gk_fgqzt_Table.Clear()
        ad.Fill(gk_fgqzt_Table)
        ad.Dispose()
        zf_gkxh = 0
        'all_fgq_zk = True
        For Each gk_row In gk_Table.Select
            '******************************************************************************************************************************************************
            '查所有的锚定工具是否坐卡
            '
            'tmp_i---当前工况锚定工具的个数，                        tmp_i=0，在“工况_锚定元件”表中找当前工况锚定工具的个数，若找到大于0，若没找到等于0。--反映有锚定工具  
            'tmp_j---当前工况"未(无)锚定"的锚定工具个数              反映是否全部锚定。
            '     没有锚定工具时：tmp_i = 0 and tmp_j = 0
            '         全部锚定时：tmp_i > 0 and tmp_j = 0  
            '         部分锚定时：tmp_j > 0 And tmp_j < tmp_i
            '       全部未锚定时：tmp_j > 0 And tmp_j = tmp_i
            '******************************************************************************************************************************************************
            foundRows = gk_mdgjzt_Table.Select("工况序号=" & gk_row.Item("工况序号").ToString)
            tmp_i = foundRows.Length
            tmp_j = 0
            If tmp_i > 0 Then
                For Each gk_mdgjzt_row In foundRows
                    If gk_mdgjzt_row.Item("定位方式").ToString = "未(无)锚定" Then
                        tmp_j = tmp_j + 1
                    End If
                Next
            End If
            '******************************************************************************************************************************************************
            '            查所有的封隔器是否坐卡（锚定）。对于多封隔器，包括水力扩张式的，应选“双向固定”。将来利用轴力差判断是否滑（蠕）动。
            '
            'tmp_k---当前工况封隔器的个数，                         在“工况_封隔定位元件”表中找当前工况，若找到大于0，若没找到等于0。--反映有封隔器
            'tmp_l---当前工况"未坐封"的封隔器个数                   反映是否全部坐封。
            '     没有封隔器时：tmp_k = 0  and tmp_l = 0
            '       全部坐封时：tmp_k > 0 and tmp_l = 0 
            '       部分坐封时：tmp_l > 0 And tmp_l < tmp_k
            '     全部未坐封时：tmp_l > 0 And tmp_l = tmp_k
            '******************************************************************************************************************************************************
            foundRows = gk_fgqzt_Table.Select("工况序号=" & gk_row.Item("工况序号").ToString)
            tmp_k = foundRows.Length
            tmp_l = 0
            If tmp_k > 0 Then
                For Each gk_fgqzt_row In foundRows
                    If gk_fgqzt_row.Item("定位方式").ToString = "未(无)锚定" Or gk_fgqzt_row.Item("封隔器状态").ToString = "未坐封" Then
                        tmp_l = tmp_l + 1
                    End If
                Next
            End If
            '******************************************************************************************************************************************************
            '                             坐卡（锚定）工况判断,顺便发现坐封(座卡)工况的前面工况锚定工具及封隔器设置不合理的情况
            '                                                                                                                 2021年10月24日秦彦斌最后修订整理
            '1.可计算的“正确”坐封工况。需给变量zf_gkxh赋值，记录坐封工况序号。
            '   可算的坐封(座卡)情况1：(1)有锚定工具，全部锚定；(2)有封隔器，全部坐封锚定。                tmp_i > 0 and tmp_j = 0 and tmp_k > 0 and tmp_l = 0
            '   可算的坐封(座卡)情况2：(1)无锚定工具；(2)有封隔器，全部坐封锚定。                          tmp_i = 0 and tmp_j = 0 and tmp_k > 0 and tmp_l = 0
            '   可算的坐封(座卡)情况3：(1)有锚定工具，全部锚定；(2)无封隔器。                              tmp_i > 0 and tmp_j = 0 and tmp_k = 0  and tmp_l = 0
            '                     
            '
            '2.矛盾的、让计算混乱的“不正确”坐封工况，需要终止程序，提醒用户。
            '   矛盾的、让计算混乱的情况1：（1）有锚定工具，部分锚定；（2）有无封隔器无所谓。              tmp_j > 0 And tmp_j < tmp_i
            '   矛盾的、让计算混乱的情况2：（1）有无锚定工具无所谓；（2）有封隔器，部分坐封锚定。          tmp_l > 0 And tmp_l < tmp_k
            '   矛盾的、让计算混乱的情况3：（1）有锚定工具，全部未锚定；(2)有封隔器，全部坐封锚定。        tmp_j > 0 And tmp_j = tmp_i and tmp_k > 0 and tmp_l = 0
            '   矛盾的、让计算混乱的情况4：（1）有锚定工具，全部锚定；(2)有封隔器，全部未坐封锚定。        tmp_i > 0 and tmp_j = 0 and tmp_l > 0 And tmp_l = tmp_k
            '
            '3.可计算的“未”坐封(座卡)工况。变量zf_gkxh=0，正常计算，不提示。
            '   可计算的“未”坐封(座卡)工况情况1：(1)无锚定工具；(2)无封隔器。
            '   可计算的“未”坐封(座卡)工况情况2：(1)无锚定工具；(2)有封隔器，全部未坐封锚定。
            '   可计算的“未”坐封(座卡)工况情况3：(1)有锚定工具，全部未锚定；(2)无封隔器。
            '******************************************************************************************************************************************************
            If (tmp_i > 0 And tmp_j = 0 And tmp_k > 0 And tmp_l = 0) Or (tmp_i = 0 And tmp_j = 0 And tmp_k > 0 And tmp_l = 0) Or (tmp_i > 0 And tmp_j = 0 And tmp_k = 0 And tmp_l = 0) Then
                zf_gkxh = Val(gk_row.Item("工况序号").ToString)
                Exit For
            End If

            If (tmp_j > 0 And tmp_j < tmp_i) Then
                lxfx_prs = False
                msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，锚定工具定位方式设置不合理（锚定工具应全锚定）。" & "请退出本界面，在工况参数设置界面正确设置锚定工具的定位方式。"
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
            If (tmp_l > 0 And tmp_l < tmp_k) Then
                lxfx_prs = False
                msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，封隔器坐封及定位方式设置不合理（封隔器应全坐封全锚定）。" & "请退出本界面，在工况参数设置界面正确设置封隔器的坐封及定位方式。"
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
            If (tmp_j > 0 And tmp_j = tmp_i And tmp_k > 0 And tmp_l = 0) Then
                lxfx_prs = False
                msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，封隔器已全坐封锚定，但锚定工具未全锚定，如此设置不合理（锚定工具应全锚定，封隔器应全坐封全锚定）。" & "请退出本界面，在工况参数设置界面正确设置坐封工况封隔器或锚定工具的坐封及定位方式。"
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
            If (tmp_i > 0 And tmp_j = 0 And tmp_l > 0 And tmp_l = tmp_k) Then
                lxfx_prs = False
                msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，锚定工具已全锚定，但封隔器未全坐封锚定，如此设置不合理（锚定工具应全锚定，封隔器应全坐封全锚定）。" & "请退出本界面，在工况参数设置界面正确设置坐封工况封隔器或锚定工具的坐封及定位方式。"
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
        Next
        If zf_gkxh <> 0 Then
            msg_prompt = "        坐封工况是管柱力学计算中的重要工况，是后续工况载荷、变形、效应计算的基准。" & Chr(13) & Chr(10) & "        " & CStr(zf_gkxh) & "号工况是坐封工况吗？"
            msg_buttons = 4 + 32
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            If msg_return <> 6 Then
                lxfx_prs = False
                msg_prompt = "        坐封工况封隔器或锚定工具定位方式设置不合理（锚定工具应全锚定，封隔器应全坐封全锚定），请退出本界面，" & "在工况参数设置界面正确设置坐封工况封隔器或锚定工具的坐封及定位方式。"
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
        End If
        '******************************************************************************************************************************************************
        '                                          对多封隔器（锚定工具）管柱的封隔器（锚定工具）定位方式进一步检验
        '                                                                                                                 2022年2月25日秦彦斌最后修订整理
        '     由于每一封隔器（锚定工具）理论上有四种定位方式，对于多封隔器（锚定工具）管柱，多个封隔器（锚定工具）的定位就会有多种组合。每种组合变形协调计算
        ' 的区间会因定位方式、管柱伸长与缩短而不同，穷举各种迭代情况并算法实现非常困难。
        '     为简化算法并程序实现，对于单封隔器（锚定工具）管柱，允许那一个封隔器（锚定工具）可有四种定位方式；对于多封隔器（锚定工具）管柱，所有的定位方式
        ' 只能是“双向固定”。所以对于多封隔器，包括水力扩张式的，应选“双向固定”。将来利用轴力差判断是否滑（蠕）动，利用变形效应判断滑（蠕）动距离。
        '     在工况输入保存及管柱力学分析计算前进行此检验
        '
        '     封隔器（锚定工具）定位方式的进一步检验算法：
        '       统计前工况锚定工具的个数tmp_i，当前工况定位方式为“双向固定”的锚定工具个数tmp_m
        '       统计前工况封隔器的个数tmp_k，当前工况定位方式为“双向固定”的封隔器个数tmp_n
        '       if （当前工况锚定工具的个数tmp_i+当前工况封隔器的个数tmp_k）>1，表示该管柱是多封隔器（锚定工具）管柱
        '           if 定位方式为“双向固定”的锚定工具个数tmp_m<>当前工况锚定工具的个数tmp_i
        '              给提示，退出
        '           endif
        '           if  定位方式为“双向固定”的封隔器个数tmp_n<>当前工况锚定工具的个数tmp_k
        '              给提示，退出
        '           endif
        '      endif
        '******************************************************************************************************************************************************
        For Each gk_row In gk_Table.Select
            '******************************************************************************************************************************************************
            '查所有的锚定工具是否坐卡
            '
            'tmp_i---当前工况锚定工具的个数，                        tmp_i=0，在“工况_锚定元件”表中找当前工况锚定工具的个数，若找到大于0，若没找到等于0。--反映有锚定工具  
            'tmp_m---当前工况定位方式为“双向固定”的锚定工具个数
            '******************************************************************************************************************************************************
            foundRows = gk_mdgjzt_Table.Select("工况序号=" & gk_row.Item("工况序号").ToString)
            tmp_i = foundRows.Length
            tmp_j = 0
            tmp_m = 0
            If tmp_i > 0 Then
                For Each gk_mdgjzt_row In foundRows
                    If gk_mdgjzt_row.Item("定位方式").ToString = "双向固定" Then
                        tmp_m = tmp_m + 1
                    End If
                    If gk_mdgjzt_row.Item("定位方式").ToString = "未(无)锚定" Then
                        tmp_j = tmp_j + 1
                    End If
                Next
            End If
            '******************************************************************************************************************************************************
            '            查所有的封隔器是否坐卡（锚定）。对于多封隔器，包括水力扩张式的，应选“双向固定”。将来利用轴力差判断是否滑（蠕）动。
            '
            'tmp_k---当前工况封隔器的个数，                         在“工况_封隔定位元件”表中找当前工况，若找到大于0，若没找到等于0。--反映有封隔器
            'tmp_n---当前工况定位方式为“双向固定”的封隔器个数封隔器个数 
            '******************************************************************************************************************************************************
            foundRows = gk_fgqzt_Table.Select("工况序号=" & gk_row.Item("工况序号").ToString)
            tmp_k = foundRows.Length
            tmp_l = 0
            tmp_n = 0
            If tmp_k > 0 Then
                For Each gk_fgqzt_row In foundRows
                    If gk_fgqzt_row.Item("定位方式").ToString = "未(无)锚定" Or gk_fgqzt_row.Item("封隔器状态").ToString = "未坐封" Then
                        tmp_l = tmp_l + 1
                    End If
                    If gk_fgqzt_row.Item("定位方式").ToString = "双向固定" Then
                        tmp_n = tmp_n + 1
                    End If
                Next
            End If
            If (tmp_i + tmp_k) > 1 Then
                If tmp_j = 0 And (tmp_i <> tmp_m) Then
                    lxfx_prs = False
                    msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，多封隔器（锚定工具）管柱的锚定工具定位方式应全设置为【双向固定】。" & "请退出本界面，在工况参数设置界面正确设置锚定工具的定位方式。"
                    lxfx_prs = False
                    cn_userdb.Close()
                    cn_userdb.Dispose()
                    Exit Function
                End If
                If tmp_l = 0 And (tmp_k <> tmp_n) Then
                    lxfx_prs = False
                    msg_prompt = "序号为" & Trim(gk_row.Item("工况序号").ToString) & "的工况中，多封隔器（锚定工具）管柱的封隔器定位方式应全设置为【双向固定】。" & "请退出本界面，在工况参数设置界面正确设置封隔器的定位方式。"
                    lxfx_prs = False
                    cn_userdb.Close()
                    cn_userdb.Dispose()
                    Exit Function
                End If
            End If
        Next
        '计算管柱总长度
        total_length = 0.0#
        SQL_command = "select sum(元件长度m) as 总长度  from 管柱数据表 where  井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        RECreader.Read()
        total_length = IIf(RECreader.Item("总长度").ToString() = "", 0.0, Val(RECreader.Item("总长度").ToString()))
        cn_userdb.Close()
        cn_userdb.Dispose()
        ' 1、对管柱进行划分节点，划分结果保存在“节点情况表”中；
        Call jslb(dlt_L, total_length, "节点情况表")
        ' 计算各节点内外径狗腿度等，填写节点情况表
        Call cal_zhijing("节点情况表")
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '读取节点参数
        SQL_command = "select * from 节点情况表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 节点编号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        jdqk_Table.Clear()
        ad.Fill(jdqk_Table)
        ad.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
        '节点数
        total_jd = jdqk_Table.Rows.Count
        ReDim jdsj_base(total_jd)
        ReDim jdsj_cacu(total_gk, total_jd)
        ReDim bx_wd(total_jd)
        ReDim bx_zl(total_jd)
        ReDim bx_gz(total_jd)
        ReDim bx_lx(total_jd)
        ReDim zf_wd_bx(total_jd)
        ReDim zf_zl_bx(total_jd)
        ReDim zf_gz_bx(total_jd)
        ReDim zf_lx_bx(total_jd)
        ReDim zf_zh_bx(total_jd)
        ReDim qe_arr(total_jd)
        '给几个数组赋初值
        For i = 0 To total_jd - 1 Step 1
            bx_wd(i) = 0.0#
            bx_zl(i) = 0.0#
            bx_gz(i) = 0.0#
            bx_lx(i) = 0.0#
            zf_wd_bx(i) = 0.0#
            zf_zl_bx(i) = 0.0#
            zf_gz_bx(i) = 0.0#
            zf_lx_bx(i) = 0.0#
            zf_zh_bx(i) = 0.0#
            qe_arr(i) = 0.0#
            For j = 0 To total_gk - 1 Step 1
                jdsj_cacu(j, i).aqxs_S = 0
                jdsj_cacu(j, i).gnyl = 0 '节点管内压MPa  FLOAT,
                jdsj_cacu(j, i).gnyl_ok = 0 '管内液压OK  TINYINT ,
                jdsj_cacu(j, i).gwyl = 0 '节点管外压MPa FLOAT,
                jdsj_cacu(j, i).gwyl_ok = 0 '管外液压OK TINYINT,
                jdsj_cacu(j, i).sj_zaihe = 0 '节点管柱实载N FLOAT,
                jdsj_cacu(j, i).sj_zh_ok = 0 '管柱实载OK  TINYINT ,
                jdsj_cacu(j, i).dx_zaihe = 0 '节点管柱等效载荷N  FLOAT,
                jdsj_cacu(j, i).dx_xczl = 0 '节点等效悬持轴力
                jdsj_cacu(j, i).dx_zh_ok = 0 '管柱效载OK  TINYINT,
                jdsj_cacu(j, i).jchl_N = 0 '接触力N ,
                jdsj_cacu(j, i).hwj_M = 0 '合弯矩Nm  FLOAT,
                jdsj_cacu(j, i).hwj_ok = 0 '合弯矩OK TINYINT,
                jdsj_cacu(j, i).wendu = 0 '节点温度   FLOAT,
                jdsj_cacu(j, i).wd_bx = 0 '温度变形m  FLOAT
                jdsj_cacu(j, i).zl_bx = 0 '轴力变形m FLOAT,
                jdsj_cacu(j, i).gz_bx = 0 '鼓胀变形m FLOAT,
                jdsj_cacu(j, i).lx_bx = 0 '螺旋变形m FLOAT,
                jdsj_cacu(j, i).zh_bx = 0 '综合变形m FLOAT,
                jdsj_cacu(j, i).wd_xy = 0 '温度变形效应m  FLOAT
                jdsj_cacu(j, i).zl_xy = 0 '轴力变形效应m FLOAT,
                jdsj_cacu(j, i).gz_xy = 0 '鼓胀变形效应m FLOAT,
                jdsj_cacu(j, i).lx_xy = 0 '螺旋变形效应m FLOAT,
                jdsj_cacu(j, i).zh_xy = 0 '综合变形效应m FLOAT,
                jdsj_cacu(j, i).xgm_xd4 = 0 '合成应力MPa
                jdsj_cacu(j, i).aqxs_S = 0 '安全系数
                jdsj_cacu(j, i).niuju = 0 '扭矩Nm
                jdsj_cacu(j, i).TSM_OK = 0 'TSM_OK    TINYINT
                jdsj_cacu(j, i).zzhpl = 0 '纵振频率Hz    FLOAT
                jdsj_cacu(j, i).hzhpl = 0 '横振频率Hz    FLOAT
                jdsj_cacu(j, i).zzhss_dep = 0 '纵振损伤深mm  FLOAT
                jdsj_cacu(j, i).hzhss_dep = 0 '横振损伤深mm  FLOAT
                jdsj_cacu(j, i).Fecrs = 999999999.99 '正弦屈曲临界载荷    FLOAT
                jdsj_cacu(j, i).Fecrh = 999999999.99 '螺旋屈曲临界载荷    FLOAT
                jdsj_cacu(j, i).mcxs = 0.3 '油套摩擦系数  FLOAT
                jdsj_cacu(j, i).hk_Yeti_midu = 0 '环液密度g╱cm3 FLOAT
                jdsj_cacu(j, i).gn_Yeti_midu = 0 '管液密度g╱cm3 FLOAT
                jdsj_cacu(j, i).kjqd_p = 0 '抗挤强度MPa FLOAT
            Next j
        Next i
        i = 0
        For Each jdqk_row In jdqk_Table.Select
            jdsj_base(i).leixing = jdqk_row.Item("节点类型").ToString
            jdsj_base(i).ID = jdqk_row.Item("节点ID").ToString
            jdsj_base(i).xingzhi = jdqk_row.Item("节点性质").ToString
            jdsj_base(i).xiashen = Val(jdqk_row.Item("节点下深m").ToString)
            jdsj_base(i).bianhao = Val(jdqk_row.Item("节点编号").ToString)
            jdsj_base(i).chuishen = Val(jdqk_row.Item("节点垂深m").ToString)
            jdsj_base(i).jxj = Val(jdqk_row.Item("井斜角rad").ToString)
            jdsj_base(i).fwj = Val(jdqk_row.Item("方位角rad").ToString)
            jdsj_base(i).tgnj = Val(jdqk_row.Item("套管内径mm").ToString)
            jdsj_base(i).ygwj = Val(jdqk_row.Item("油管外径mm").ToString)
            jdsj_base(i).ygnj = Val(jdqk_row.Item("油管内径mm").ToString)
            jdsj_base(i).xzh = Val(jdqk_row.Item("线重kg╱m").ToString)
            jdsj_base(i).xgm_b = Val(jdqk_row.Item("屈服强度MPa").ToString)
            jdsj_base(i).yg_gj = jdqk_row.Item("油管钢级").ToString
            jdsj_base(i).gtd = Val(jdqk_row.Item("曲率rad╱m").ToString)
            jdsj_base(i).e = Val(jdqk_row.Item("弹性模量MPa").ToString)
            jdsj_base(i).psb = Val(jdqk_row.Item("泊松比").ToString)
            jdsj_base(i).rzhxsh = Val(jdqk_row.Item("热膨胀系数").ToString)
            jdsj_base(i).kjqdu = Val(jdqk_row.Item("抗挤强度MPa").ToString) '抗挤强度MPa
            jdsj_base(i).gtkny = Val(jdqk_row.Item("管体抗内压MPa").ToString) '管体抗内压MPa
            jdsj_base(i).jtkny = Val(jdqk_row.Item("接头抗内压MPa").ToString) '接头抗内压MPa
            jdsj_base(i).gtkla = Val(jdqk_row.Item("管体抗拉kN").ToString) '管体抗拉kN
            jdsj_base(i).jtkla = Val(jdqk_row.Item("接头抗拉kN").ToString) '接头抗拉kN 
            i = i + 1
        Next
        frmprogress.Show()
        frmprogress.Label1.Text = "正在进行计算，请耐心等候......"
        frmprogress.PB1.Maximum = total_gk
        '**************************************************************************************************************************************
        '再次循环工况，进行计算
        '   循环前取得下钻时的地面温度
        '   循环前取得完井液密度
        '**************************************************************************************************************************************
        TMP_so = Val(gk_Table.Rows(0).Item("井口温度℃").ToString)
        gm = Val(gk_Table.Rows(0).Item("环液密度g╱cm3").ToString) * 1000.0
        PB1_cout = 0
        gk_gkxh = 0
        qe_weifu = False
        For Each gk_row In gk_Table.Select
            '**************************************************************************************************************************************
            '1. 工况序号连续性检查
            '**************************************************************************************************************************************
            gk_gkxh_zd = Val(gk_row.Item("工况序号").ToString)
            gk_gkxh = gk_gkxh + 1
            If gk_gkxh <> gk_gkxh_zd Then
                msg_prompt = "   工况数据中，" & CStr(gk_gkxh_zd) & "号工况的工况序号与前一工况序号不连续，计算无法进行。请退出本界面，在工况参数输入界面中改正，然后再次进入本界面进行计算。"
                frmprogress.Hide()
                lxfx_prs = False
                Exit Function
            End If
            PB1_cout = PB1_cout + 1
            frmprogress.PB1.Value = PB1_cout
            '**************************************************************************************************************************************
            '2. 对每一工况，给节点计算参数表填初始值。当计算过程中异常退出造成数据未填入而为NULL，打开数据文件时，因NULL会出错，故在此全部赋初值。
            '**************************************************************************************************************************************
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            For i = 0 To total_jd - 1
                SQL_command = "insert into  节点计算参数表 (作业名称,工况序号,节点编号,节点下深m,管内压力MPa,管内液压OK,管外压力MPa,管外液压OK," _
                    & "真实轴力N,真实轴力OK,等效轴力N,等效轴力OK,接触力N,合弯矩Nm,合弯矩OK,节点温度℃,温度变形m,轴力变形m,鼓胀变形m,螺旋变形m," _
                    & " 综合变形m,合成应力MPa,安全系数,等效悬持力N,温度效应m,轴力效应m,鼓胀效应m,螺旋效应m,综合效应m,扭矩Nm,TSM_OK,纵振频率Hz," _
                    & " 横振频率Hz,纵振损伤深mm,横振损伤深mm,Fecrs_N,Fecrh_N,油套摩擦系数,环液密度g╱cm3,管液密度g╱cm3,抗挤强度MPa,井号) values (" _
                    & "'" & zuoye_name & "'," & CStr(gk_gkxh) & "," & CStr(jdsj_base(i).bianhao) & "," & CStr(jdsj_base(i).xiashen) & "," _
                    & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,999999999.99,999999999.99,0.3,0,0,0,'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            Next i
            cn_userdb.Close()
            cn_userdb.Dispose()
            '***************************************************************************************************************************
            '3. 计算该工况下各节点的液压压强，递推管内、管外流体密度，计算结果填写到“节点计算参数表”中；
            '***************************************************************************************************************************
            If yeya(gk_Table, gk_gkxh, total_jd, jdsj_base, jdsj_cacu) = False Then
                frmprogress.Hide()
                lxfx_prs = False
                Exit Function
            End If
        Next
        frmprogress.Hide()
        If qe_weifu = True Then
            MsgBox(qe_weifumsg)
        End If
        MsgBox("管柱压力场分析计算完成。") '调试语句
    End Function
End Module
