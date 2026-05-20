Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_cal_yeya
    '************************************************************************************************************************************************
	'                                                   每一工况下各分段液压计算子程序
    '                                                                                                        2024年7月14日秦彦斌最后修订整理
    '输入参数：
    '   gk_Table               按工况序号排序的工况参数表
    '   gk_gkxh                工况序号
    '   total_jd               节点总数
    '   jdsj_base()            一维数组，存放当前作业管柱的“节点情况表”数据
    '   jdsj_cacu()            二维数组，
    '全局变量使用：
    '   Public well_name As String          '当前井号
    '   Public zuoye_name As String         '当前管柱作业名称
    '   Public TSM_ver_switch As Integer    '管柱力学分析软件版本控制开关，等于1，垂直井节点法，井斜、方位均等于0，不判断、不读取井斜数据
    '   Public Testwell_Table As DataTable = dst.Tables.Add("Testwell_Table")          测斜数据
    '编程思路：
    '   (1) 对于每一工况，分管内、管外分别计算各节点压力。
    '   (2)
    '   技巧之1：设置压力连续标志变量jd_gnyl_cont()和jd_gwyl_cont()，用于标志该节点处压力是否连续及计算取值方法
    '   技巧之2：设置压力连续标志变量nwyl_cont_hk和nwyl_cont_gz，用于标志井底是否连通。
    '另：
    'gnyl_ok、gwyl_ok值若为1 表示根据已知条件递推赋值
    'gnyl_ok、gwyl_ok值若为2 表示没有能正常递推赋值，赋值地层压力
    'gnyl_ok、gwyl_ok值若为3 ，原来的设置，现放弃不用。原表示非初始工况，没有能正常递推赋值，赋值为上一工况的对应点的压力值
    'gnyl_ok、gwyl_ok值若为4 表示根据开关工具或节流工具内外连通条件递推赋值
    'gnyl_ok、gwyl_ok值若为5 表示根据井底内外连通条件递推赋值
    '
    '使用注意事项：
    '   1、使用前需用函数gzlxfx_data_ok()判断数据是否完整，否则运行程序会出错。
    '   2、程序没有考虑流动时的沿程压力损失或者说压力变化
    '
    '   开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
    '
    '运行结果：在“节点计算参数表”中填入合适的数据
    '
    ' 重要修订记录:
    '    2013年02月，增加了“井底套压MPa、井底管压MPa、压力计算开关”3个字段，实现相应压力计算。
    '    2013年04月01日，删“压力计算开关”字段，增“井底管压开关”、“井底套压开关”两字段，用以描述井底压力计算方式。实现相应压力计算。
    '    2015年02月22日，增“井口管压开关”、“井口套压开关”两字段，用以描述井口压力计算方式。实现相应压力计算。
    '    2015年02月27日，改字段“井底深度m”为“管压井底深度m”，增“套压井底深度m”，用以描述井底压力对应的下深。实现相应压力计算。
    '    2017年01月-7月，修订程序，实现考虑开关元件嘴损引起的压力变化。
    '    2017年07月23日，实现了考虑开关工具嘴损压力计算，考虑可洗井封隔器在洗井时环空压力可递推计算。
    '    2020年01月27日，程序迁移到VS2008完成，工况参数用gk_Table As DataTable来传递,涉及一行参数时，用dadarow类型变量进行传递，测斜数据全局变量化。
    '    2022年01月20日，增加对于同一深度的两节点，将已知压力的节点压力直接赋值给未知压力的节点压力变量代码。解决因管柱元件两头都设节点，液压递推时
    '                    管柱元件两头压力计算错误问题。在测试枣1288-5井时发现该问题。
    '    2022年07月19-20日，修改用流体摩阻计算模型计算管内（环空）单位长度流体摩阻函数，改工况记录行传递为结构体calmz_ltcs传递，以便通用。
    '                       修改每一工况下各分段液压计算yeya函数代码，以适应如此改动。
    '                       参照封隔器管柱力学的压力计算部分重写了压力计算代码。
    '    2024年07月11-14日，修改压力递推计算代码，实现正确计算坐封且井下关井工况下封隔器以下环空及关井阀以下管内的压力。
    '************************************************************************************************************************************************
    Public Function yeya(ByVal gk_Table As DataTable, ByVal gk_gkxh As Short, ByVal total_jd As Short, ByRef jdsj_base() As jd_base, ByRef jdsj_cacu(,) As jd_cacu) As Boolean '各分段液压计算
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim gk_fdRow() As DataRow
        Dim SQL_command As String
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim RECreader2 As OleDbDataReader
        Dim gk_dytd As Double '地压梯度
        Dim jd_count As Short
        Dim i As Short
        Dim nwyl_cont_hk As Boolean
        Dim nwyl_cont_gn As Boolean
        Dim jd_gnyl_cont() As Byte
        Dim jd_gwyl_cont() As Byte
        Dim dlt_ppm_gn() As Double '每米长度管内压力变化量即压力梯度。单位MP/m
        Dim dlt_ppm_gw() As Double '每米长度管外压力变化量即压力梯度。单位MP/m
        Dim bl_temp0 As Double '计算中的临时变量
        Dim bl_temp1 As Double '计算中的临时变量
        Dim bl_temp2 As Double '计算中的临时变量
        Dim bl_temp3 As Double '计算中的临时变量
        Dim bl_temp4 As Double '计算中的临时变量
        Dim hym_chuishen As Double '环空液面垂深
        Dim gym_chuishen As Double '管内液面垂深
        Dim gn_calmz_ltcs As calmz_ltcs
        Dim gw_calmz_ltcs As calmz_ltcs

        ReDim jd_gnyl_cont(total_jd - 1)
        ReDim jd_gwyl_cont(total_jd - 1)
        ReDim dlt_ppm_gn(total_jd - 1)
        ReDim dlt_ppm_gw(total_jd - 1)
        yeya = True '若运行无误，则返回true
        nwyl_cont_hk = True                     '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
        nwyl_cont_gn = True                     '管内上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
        bl_temp0 = 0.0#
        bl_temp1 = 0.0#
        bl_temp2 = 0.0#
        bl_temp3 = 0.0#
        bl_temp4 = 0.0#
        '**************************************************************************************************************************************
        '查算压力梯度（压力系数）
        '    先在“作业地层参数表”表中找对应作业的压力系数，若找不到，自己算。算法认为第一工况时，井底地层压力等于压井液压力，故压力系数等于压
        ' 井液密度, 计算时用管外压力
        '*************************************************************************************************************************************
        gk_dytd = Val(gk_Table.Rows(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000.0#

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select top 1 * from 作业地层参数表 where 井号='" & well_name & " ' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            If Val(RECreader.Item("压力系数").ToString) <> 0 Then
                gk_dytd = Val(RECreader.Item("压力系数").ToString) * 9.8 / 1000.0#
            End If
        End If
        RECreader.Close()
        cn_userdb.Close()
        '**************************************************************************************************************************************
        '定位到指定工况
        '*************************************************************************************************************************************
        gk_fdRow = gk_Table.Select("工况序号=" & CStr(gk_gkxh))
        gw_calmz_ltcs.mozu_model = gk_fdRow(0).Item("环空流阻模型").ToString              '流体摩阻计算模型
        gw_calmz_ltcs.liuliang = Val(gk_fdRow(0).Item("环流流量m3╱m").ToString)          '流体流量(m^3/min)
        gw_calmz_ltcs.Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)        '流体密度(g/cm^3)
        gw_calmz_ltcs.Yeti_niandu = Val(gk_fdRow(0).Item("环液粘度mPaS").ToString)        '流体动力粘度(mPa·s)
        gw_calmz_ltcs.mozu_xishu = Val(gk_fdRow(0).Item("环牛模折减系数").ToString)       '牛模摩阻折减系数
        gw_calmz_ltcs.liuxiang = gk_fdRow(0).Item("环空流体流向").ToString                '流体流向(不流动、向下、向上)
        gw_calmz_ltcs.chjnd = Val(gk_fdRow(0).Item("环空稠剂浓度").ToString)              '稠化剂浓度(kg/m^3)
        gw_calmz_ltcs.zcjnd = Val(gk_fdRow(0).Item("环空撑剂浓度").ToString)              '支撑剂浓度(kg/m^3)
        gw_calmz_ltcs.lbzsh = Val(gk_fdRow(0).Item("环空流变指数n").ToString)             '流体流变指数n
        gw_calmz_ltcs.chdxsh = Val(gk_fdRow(0).Item("环空稠度系数K").ToString)            '稠度系数K(Pa·s^n)

        gn_calmz_ltcs.mozu_model = gk_fdRow(0).Item("管内流阻模型").ToString              '流体摩阻计算模型
        gn_calmz_ltcs.liuliang = Val(gk_fdRow(0).Item("管流流量m3╱m").ToString)          '流体流量(m^3/min)
        gn_calmz_ltcs.Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)        '流体密度(g/cm^3)
        gn_calmz_ltcs.Yeti_niandu = Val(gk_fdRow(0).Item("管流粘度mPaS").ToString)        '流体动力粘度(mPa·s)
        gn_calmz_ltcs.mozu_xishu = Val(gk_fdRow(0).Item("管牛模折减系数").ToString)       '牛模摩阻折减系数
        gn_calmz_ltcs.liuxiang = gk_fdRow(0).Item("管内流体流向").ToString                '流体流向(不流动、向下、向上)
        gn_calmz_ltcs.chjnd = Val(gk_fdRow(0).Item("管内稠剂浓度").ToString)              '稠化剂浓度(kg/m^3)
        gn_calmz_ltcs.zcjnd = Val(gk_fdRow(0).Item("管内撑剂浓度").ToString)              '支撑剂浓度(kg/m^3)
        gn_calmz_ltcs.lbzsh = Val(gk_fdRow(0).Item("管内流变指数n").ToString)             '流体流变指数n
        gn_calmz_ltcs.chdxsh = Val(gk_fdRow(0).Item("管内稠度系数K").ToString)            '稠度系数K(Pa·s^n)
        '**************************************************************************************************************************************
        ' 环空、管内液面垂深计算
        ' 若井筒内上部为有压力或密度较大的气，下部为液体，压力计算还是不对。
        ' 空气密度=1.293*(实际压力/标准物理大气压)*(273/实际绝对温度)，绝对温度=摄氏温度+273.15
        ' 通常情况下，即20摄氏度时，取1.205kg/m3
        '**************************************************************************************************************************************
        hym_chuishen = 0.0#
        If Val(gk_fdRow(0).Item("环液深度m").ToString) > 0 Then
            If TSM_ver_switch = 1 Then
                hym_chuishen = Val(gk_fdRow(0).Item("环液深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("环液深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, hym_chuishen)
            End If
        End If
        gym_chuishen = 0.0#
        If Val(gk_fdRow(0).Item("管液深度m").ToString) > 0 Then
            If TSM_ver_switch = 1 Then
                gym_chuishen = Val(gk_fdRow(0).Item("管液深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("管液深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, gym_chuishen)
            End If
        End If
        '**************************************************************************************************************************************
        ' 在井口、井底管内压力输入的情况下，管内流体因粘滞摩阻压力梯度赋值
        '**************************************************************************************************************************************
        For i = 1 To total_jd - 1
            dlt_ppm_gn(i) = 0.0#
        Next i
        If gk_fdRow(0).Item("井底管压开关").ToString = "输入" And gk_fdRow(0).Item("井口管压开关").ToString = "输入" Then
            '**********************************************************************************************************************************
            '求管压井底深度，管内液面处的垂深，bl_temp3，bl_temp4
            '
            'Sub cal_jx_fw_cs(ByVal in_depth As Double, ByRef ret_jxj As Double, ByRef ret_fwj As Double, ByRef ret_csh As Double)
            '子程序以传地址的方式计算出给定井任意深度in_depth处的：
            '      井斜角 单位：弧度 实型数
            '      方位角 单位：弧度 实型数
            '      垂深   单位：米 实型数
            '**********************************************************************************************************************************
            If TSM_ver_switch = 1 Then
                bl_temp3 = Val(gk_fdRow(0).Item("管压井底深度m").ToString)
                bl_temp4 = Val(gk_fdRow(0).Item("管液深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("管压井底深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp3)
                bl_temp0 = Val(gk_fdRow(0).Item("管液深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp4)
            End If
            For i = 1 To total_jd - 1
                '**********************************************************************************************************************************
                ' 问题：若中间有开关元件怎么办？
                '**********************************************************************************************************************************
                dlt_ppm_gn(i) = (Val(gk_fdRow(0).Item("井口管压MPa").ToString) + (bl_temp3 - bl_temp4) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 - Val(gk_fdRow(0).Item("井底管压MPa").ToString)) / (Val(gk_fdRow(0).Item("管压井底深度m").ToString) - Val(gk_fdRow(0).Item("管液深度m").ToString))
                'If SQL_rst_gk.Fields("管内流体流向") = "向上" Then
                '    dlt_ppm_gn(i) = (-1) * dlt_ppm_gn(i)
                'End If
            Next i
        End If
        '**************************************************************************************************************************************
        ' 在井口、井底环空压力输入的情况下，环空流体因粘滞摩阻压力梯度赋值
        '**************************************************************************************************************************************
        For i = 1 To total_jd - 1
            dlt_ppm_gw(i) = 0.0#
        Next i
        If gk_fdRow(0).Item("井底套压开关").ToString = "输入" And gk_fdRow(0).Item("井口套压开关").ToString = "输入" Then
            '**********************************************************************************************************************************
            '求套压井底深度，环空液面处的垂深，bl_temp3，bl_temp4
            '
            'Sub cal_jx_fw_cs(ByVal in_depth As Double, ByRef ret_jxj As Double, ByRef ret_fwj As Double, ByRef ret_csh As Double)
            '子程序以传地址的方式计算出给定井任意深度in_depth处的：
            '      井斜角 单位：弧度 实型数
            '      方位角 单位：弧度 实型数
            '      垂深   单位：米 实型数
            '**********************************************************************************************************************************
            If TSM_ver_switch = 1 Then
                bl_temp3 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
                bl_temp4 = Val(gk_fdRow(0).Item("环液深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp3)
                bl_temp0 = Val(gk_fdRow(0).Item("环液深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp4)
            End If
            For i = 1 To total_jd - 1
                '**********************************************************************************************************************************
                ' 问题：若中间有坐封的封隔器怎么办？
                '**********************************************************************************************************************************
                dlt_ppm_gw(i) = (Val(gk_fdRow(0).Item("井口环压MPa").ToString) + (bl_temp3 - bl_temp4) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - Val(gk_fdRow(0).Item("井底套压MPa").ToString)) / (Val(gk_fdRow(0).Item("套压井底深度m").ToString) - Val(gk_fdRow(0).Item("环液深度m").ToString))
                'If SQL_rst_gk.Fields("环空流体流向") = "向上" Then
                '    dlt_ppm_gw(i) = (-1) * dlt_ppm_gw(i)
                'End If
            Next i
        End If
        '**************************************************************************************************************************************
        '                                                         压力分析计算思路及算法设计
        '                                                                           秦彦斌    2015年2月23日初稿，2017年7月最后修改
        '压力分析计算思路：
        '   1. 管内井口压力为输入，环空井口压力为输入时
        '           （1）从井口由上向下递推管内压力；
        '           （2）从井口由上向下递推环空压力；
        '   2. 管内井底压力为输入，环空井口压力为输入时
        '           （1）从井底（管内井底压力值处深度）由下向上递推管内压力，从井底（管内井底压力值处深度）由上向下递推管内压力；
        '           （2）从井口由上向下递推环空压力；
        '   3. 管内井口压力为输入，环空井底压力为输入时
        '           （1）从井口由上向下递推管内压力；
        '           （2）从井底（环空井底压力值处深度）由下向上递推环空压力，从井底（环空井底压力值处深度）由上向下递推环空压力；
        '   4. 管内井底压力为输入，环空井底压力为输入时
        '           （1）从井底（管内井底压力值处深度）由下向上递推管内压力，从井底（管内井底压力值处深度）由上向下递推管内压力；
        '           （2）从井底（环空井底压力值处深度）由下向上递推环空压力，从井底（环空井底压力值处深度）由上向下递推环空压力；
        '   5.补漏计算
        '           （3）从井底由下向上递推管内压力，环空压力。递推起始点上（管柱的最下端），若管内压力已算出、环空压力未算出，则令环空压力等于
        '       管内压力；若环空压力已算出、管内压力未算出，则令管内压力等于环空压力。
        '           （4）从井口由上向下循环，给没有管内、环空压力赋值的节点赋值。赋值方法：若是初始工况，则将地层压力赋给还没有算出的节点，若是
        '       其他工况，将上一工况的对应点的数据赋给还没有算出的节点
        '
        '压力分析算法设计：
        '   1.管外压力递推：
        '       （1）环空井口压力为输入时，从井口由上向下递推环空压力；
        '       （2）环空井底压力为输入时，从井底（环空井底压力值处深度）由下向上递推环空压力，从井底（环空井底压力值处深度）由上向下递推环空压力；
        '   2.管内压力递推：
        '       （1）管内井口压力为输入时，从井口由上向下递推管内压力；
        '       （2）管内井底压力为输入时，从井底（管内井底压力值处深度）由下向上递推管内压力，从井底（管内井底压力值处深度）由上向下递推管内压力；
        '   3.利用管柱最下端管内、管外压力连通的边界条件，对没有赋值的且压力连续的点用递推公式计算管内、环空压力：
        '       （1）从井底由下向上递推管内压力，环空压力。递推起始点上（管柱的最下端），若管内压力已算出、环空压力未算出，则令环空压力等于管内
        '       压力；若环空压力已算出、管内压力未算出，则令管内压力等于环空压力。
        '   4.给没有管内、环空压力赋值的节点赋值。
        '       （1）从井口由上向下循环，对于未赋值节点，若是初始工况，则将地层压力赋给还没有算出的节点，若是其他工况，将上一工况的对应点的数据
        '       赋给还没有算出的节点
        '**************************************************************************************************************************************
        '**************************************************************************************************************************************
        '环空、管内压力参数赋初值。
        '**************************************************************************************************************************************
        For i = 0 To total_jd - 1
            '赋初值
            jdsj_cacu(gk_gkxh - 1, i).gwyl = 0.0#
            jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
            jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0
            jd_gwyl_cont(i) = 0

            jdsj_cacu(gk_gkxh - 1, i).gnyl = 0.0#
            jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
            jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0
            jd_gnyl_cont(i) = 0
        Next i
        '**************************************************************************************************************************************
        '   1.管外压力递推：
        '       （1）环空井口压力为输入时，从井口由上向下递推环空压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井口套压开关").ToString = "输入" Then
            '将井口套压直接赋值
            jdsj_cacu(gk_gkxh - 1, 0).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString)
            jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
            '如果环空液面深度不为零，井口环空流体密度为空气密度。
            '   空气密度随温度、压力的不同而不同，25摄氏度时，0.1MPa：1.1691kg/m^3，2.5MPa：29.228kg/m^3（数据来自：https://baike.baidu.com/item/%E7%A9%BA%E6%B0%94%E5%AF%86%E5%BA%A6/2995215）
            '   空气密度=1.293*(实际压力/标准物理大气压)x(273.15 / 实际绝对温度)，绝对温度=摄氏温度+273.15
            If hym_chuishen <> 0 Then
                jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu = 0.001 * 1.293 * (1) * (273.15 / (273.15 + Val(gk_fdRow(0).Item("井口温度℃").ToString)))
            End If
            jdsj_cacu(gk_gkxh - 1, 0).gwyl_ok = 1
            For i = 1 To total_jd - 1
                '************************************************************************************************************************
                '正向管外压力递推
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                    '********************************************************************************************************************
                    '上面结点管外压力已知，下面结点管外压力未知，符合递推条件
                    '********************************************************************************************************************
                    jd_gwyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).hk_Yeti_midu
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                        jd_gwyl_cont(i) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号封隔器坐封状态未知，计算无法进行。请在工况数据输入界面确定！"
                                    yeya = False
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    Exit Function
                                End If
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If
                            '环空压力计算
                            '若高于环空液面，则将井口环压直接赋值
                            If jdsj_base(i).chuishen < hym_chuishen Then
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString) + jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu * 1000 * 9.8 * jdsj_base(i).chuishen / 1000
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                            Else
                                '环空液面位于两节点之间
                                If jdsj_base(i).chuishen > hym_chuishen And jdsj_base(i - 1).chuishen < hym_chuishen Then
                                    '当前点液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu * 9.8 * (hym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                Else
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        '**************************************************************************************************************************************
        '   1.管外压力递推：
        '       （2）环空井底压力为输入时，从井底（环空井底压力值处深度）由下向上递推环空压力，从井底（环空井底压力值处深度）由上向下递推环空压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井底套压开关").ToString = "输入" Then
            '**********************************************************************************************************************************
            ' 井底环空压力值对应深度处的垂深
            '**********************************************************************************************************************************
            If TSM_ver_switch = 1 Then
                bl_temp3 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp3)
            End If
            '**********************************************************************************************************************************
            ' 找到环空压力值对应深度处的节点并给节点环空压力赋值
            '**********************************************************************************************************************************
            If Val(gk_fdRow(0).Item("套压井底深度m").ToString) > jdsj_base(total_jd - 1).xiashen Then
                jd_count = total_jd - 1
                jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString) - (bl_temp3 - jdsj_base(jd_count).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                jdsj_cacu(gk_gkxh - 1, jd_count).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
            Else
                For i = total_jd - 1 To 1 Step -1
                    If Val(gk_fdRow(0).Item("套压井底深度m").ToString) = jdsj_base(i).xiashen Then
                        jd_count = i
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString)
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
                        Exit For
                    End If
                    If Val(gk_fdRow(0).Item("套压井底深度m").ToString) < jdsj_base(i).xiashen And Val(gk_fdRow(0).Item("套压井底深度m").ToString) > jdsj_base(i - 1).xiashen Then
                        jd_count = i - 1
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString) - (bl_temp3 - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                        jdsj_cacu(gk_gkxh - 1, jd_count).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
                    End If
                Next i
            End If
            '**********************************************************************************************************************************
            ' 从井底（环空井底压力值处深度）由下向上递推环空压力。
            '**********************************************************************************************************************************
            For i = jd_count - 1 To 0 Step -1
                jd_gwyl_cont(i + 1) = 1
                If jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then '下面结点管外压力已知，上面结点管外压力未知
                    If jdsj_base(i + 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                        jd_gwyl_cont(i + 1) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i + 1).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i + 1) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i + 1) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i + 1) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i + 1) = 1 Then
                            '********************************************************************************************************************
                            '如果：井口套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If

                            '管外液面位于两节点之间
                            If jdsj_base(i).chuishen >= hym_chuishen Then
                                '前一点的液压等于当前点的压力减去液柱压力
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - (jdsj_base(i + 1).chuishen - jdsj_base(i).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gw(i) * (jdsj_base(i + 1).xiashen - jdsj_base(i).xiashen) '单位MPa
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                            Else
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - (jdsj_base(i + 1).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gw(i) * (jdsj_base(i + 1).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                            End If
                        End If
                    End If
                End If
            Next i
            '**********************************************************************************************************************************
            ' 从井底（环空井底压力值处深度）由上向下递推环空压力；
            '**********************************************************************************************************************************
            For i = jd_count + 1 To total_jd - 1
                '************************************************************************************************************************
                '正向环空压力递推
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                    '********************************************************************************************************************
                    '上面结点管外压力已知，下面结点管外压力未知，符合递推条件
                    '********************************************************************************************************************
                    jd_gwyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).hk_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                        jd_gwyl_cont(i) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If
                            '环空压力计算
                            '若高于环空液面，则将井口环压直接赋值
                            If jdsj_base(i).chuishen < hym_chuishen Then
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString)
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                            Else
                                '环空液面位于两节点之间
                                If jdsj_base(i).chuishen > hym_chuishen And jdsj_base(i - 1).chuishen < hym_chuishen Then
                                    '当前点液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu * 9.8 * (hym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                Else
                                    '环空液面不位于两节点之间
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        '**************************************************************************************************************************************
        '   2.管内压力递推：
        '       （1）管内井口压力为输入时，从井口由上向下递推管内压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井口管压开关").ToString = "输入" Then
            '则将井口管压直接赋值
            jdsj_cacu(gk_gkxh - 1, 0).gnyl = Val(gk_fdRow(0).Item("井口管压MPa").ToString)
            jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
            '如果管内液面深度不为零，井口管内流体密度为空气密度。
            '   空气密度随温度、压力的不同而不同，25摄氏度时，0.1MPa：1.1691kg/m^3，2.5MPa：29.228kg/m^3（数据来自：https://baike.baidu.com/item/%E7%A9%BA%E6%B0%94%E5%AF%86%E5%BA%A6/2995215）
            '   空气密度=1.293*(实际压力/标准物理大气压)x(273.15 / 实际绝对温度)，绝对温度=摄氏温度+273.15
            If gym_chuishen <> 0 Then
                jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu = 0.001 * 1.293 * (1) * (273.15 / (273.15 + Val(gk_fdRow(0).Item("井口温度℃").ToString)))
            End If
            jdsj_cacu(gk_gkxh - 1, 0).gnyl_ok = 1
            For i = 1 To total_jd - 1
                '************************************************************************************************************************
                '正向管内压力递推   用i-1推i，若i点不通，推不下去
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then '上面结点管内压力已知，下面结点管内压力未知
                    jd_gnyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl
                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).gn_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                        jd_gnyl_cont(i) = 1
                        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

                        '********************************************************************************************************************
                        '如果是管柱，处理打开的开关元件压力传递问题
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关元件,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                End If
                                RECreader.Read()
                                '************************************************************************************************************************
                                '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                                ' 对于开关元件
                                '若内外连通，
                                '   若管外压力未赋值，
                                '       （1）若流体流动，则考虑嘴损计算管外压力；
                                '       （2）若流体不流动，将管内压力赋值给管外压力；
                                '   若管外压力已赋值
                                '       判断压力差别是否符合哈数。
                                '若上下连通、流体流动，考虑嘴损修正计算点管内压力。
                                '************************************************************************************************************************
                                '************************************************************************************************************************
                                ' 处理内外连通
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开油套" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    '********************************************************************************************************************
                                    ' 管内知，管外未知，由管内算管外
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    '********************************************************************************************************************
                                    ' 管外知，管内未知，由管外算管内
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 And gk_fdRow(0).Item("管内流体流向").ToString = "不流动" Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            If jdsj_base(i).xingzhi = "节流工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If RECreader.Item("节流流向").ToString = "内外流通" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Else
                        '不同深度节点，递推
                        '********************************************************************************************************************
                        '如果是管柱，对于关闭的开关元件，设置压力不连续
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关元件,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                Else
                                    RECreader.Read()
                                    If RECreader.Item("开关状态").ToString = "单关管内" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "关油套关管内" Then
                                        jd_gnyl_cont(i) = 0
                                        nwyl_cont_gn = False '管内上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                        If jd_gnyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底管压开关 等于 "计算"，且管内流体流向等于“向上”或“向下”，需要用摩阻计算方法计算管内流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gn(i) = dlt_ppm_gn_ndlt(gn_calmz_ltcs, jdsj_base(i).ygnj)
                            End If
                            '管内压力计算
                            '若高于管内液面，则将井口管压直接赋值
                            If jdsj_base(i).chuishen < gym_chuishen Then
                                jdsj_cacu(gk_gkxh - 1, i).gnyl = Val(gk_fdRow(0).Item("井口管压MPa").ToString) + jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu * 9.8 * jdsj_base(i).chuishen / 1000
                                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                            Else
                                '管内液面位于两节点之间
                                If jdsj_base(i).chuishen > gym_chuishen And jdsj_base(i - 1).chuishen < gym_chuishen Then
                                    '液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl + jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu * 9.8 * (gym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - gym_chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gn(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("管液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                                Else
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gn(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                                End If
                            End If
                        End If
                        '********************************************************************************************************************
                        '如果是管柱，处理打开的开关元件压力传递问题
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关元件,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                End If
                                RECreader.Read()
                                '************************************************************************************************************************
                                '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                                ' 对于开关元件
                                '若内外连通，
                                '   若管外压力未赋值，
                                '       （1）若流体流动，则考虑嘴损计算管外压力；
                                '       （2）若流体不流动，将管内压力赋值给管外压力；
                                '   若管外压力已赋值
                                '       判断压力差别是否符合哈数。
                                '若上下连通、流体流动，考虑嘴损修正计算点管内压力。
                                '************************************************************************************************************************
                                '************************************************************************************************************************
                                ' 处理内外连通
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开油套" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    '********************************************************************************************************************
                                    ' 管内知，管外未知，由管内算管外
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    '********************************************************************************************************************
                                    ' 管外知，管内未知，由管外算管内
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 And gk_fdRow(0).Item("管内流体流向").ToString = "不流动" Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                '************************************************************************************************************************
                                ' 处理上下连通时嘴损问题
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开管内" Or RECreader.Item("开关状态").ToString = "关油套开管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                            If jdsj_base(i).xingzhi = "节流工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If RECreader.Item("节流流向").ToString = "内外流通" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        '**************************************************************************************************************************************
        '   2.管内压力递推：
        '       （2）管内井底压力为输入时，从井底（管内井底压力值处深度）由下向上递推管内压力，从井底（管内井底压力值处深度）由下向上递推管内压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井底管压开关").ToString = "输入" Then
            '**********************************************************************************************************************************
            ' 井底管内压力值对应深度处的垂深
            '**********************************************************************************************************************************
            If TSM_ver_switch = 1 Then
                bl_temp3 = Val(gk_fdRow(0).Item("管压井底深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("管压井底深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp3)
            End If
            '**********************************************************************************************************************************
            ' 找到井底管内压力值对应深度处的节点并给节点管内压力赋值
            '**********************************************************************************************************************************
            If Val(gk_fdRow(0).Item("管压井底深度m").ToString) > jdsj_base(total_jd - 1).xiashen Then
                jd_count = total_jd - 1
                jdsj_cacu(gk_gkxh - 1, jd_count).gnyl = Val(gk_fdRow(0).Item("井底管压MPa").ToString) - (bl_temp3 - jdsj_base(jd_count).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                jdsj_cacu(gk_gkxh - 1, jd_count).gnyl_ok = 1
            Else
                For i = total_jd - 1 To 1 Step -1
                    If Val(gk_fdRow(0).Item("管压井底深度m").ToString) = jdsj_base(i).xiashen Then
                        jd_count = i
                        jdsj_cacu(gk_gkxh - 1, jd_count).gnyl = Val(gk_fdRow(0).Item("井底管压MPa").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gnyl_ok = 1
                        Exit For
                    End If
                    If Val(gk_fdRow(0).Item("管压井底深度m").ToString) < jdsj_base(i).xiashen And Val(gk_fdRow(0).Item("管压井底深度m").ToString) > jdsj_base(i - 1).xiashen Then
                        jd_count = i - 1
                        jdsj_cacu(gk_gkxh - 1, jd_count).gnyl = Val(gk_fdRow(0).Item("井底管压MPa").ToString) - (bl_temp3 - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                        jdsj_cacu(gk_gkxh - 1, jd_count).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gnyl_ok = 1
                    End If
                Next i
            End If
            '**********************************************************************************************************************************
            ' 从井底（管内井底压力值处深度）由下向上递推环空压力。  用i+1推i，若i+1不通，推不上去
            '**********************************************************************************************************************************
            For i = jd_count - 1 To 0 Step -1
                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then '下面结点管内压力已知，上面结点管内压力未知
                    jd_gnyl_cont(i + 1) = 1
                    If jdsj_base(i + 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                        jd_gnyl_cont(i + 1) = 1
                    Else
                        '不同深度节点，递推
                        '如果是管柱，修正开关工具及封隔定位元件引起的压力不连续
                        If InStr(jdsj_base(i + 1).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关工具,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关工具的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i + 1).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i + 1).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                Else
                                    RECreader.Read()
                                    If RECreader.Item("开关状态").ToString = "单关管内" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "关油套关管内" Then
                                        jd_gnyl_cont(i + 1) = 0
                                        nwyl_cont_gn = False '管内上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                        If jd_gnyl_cont(i + 1) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底管压开关 等于 "计算"，且管内流体流向等于“向上”或“向下”，需要用摩阻计算方法计算管内流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gn(i) = dlt_ppm_gn_ndlt(gn_calmz_ltcs, jdsj_base(i).ygnj)
                            End If
                            If jdsj_base(i).chuishen >= gym_chuishen Then
                                '前一点的液压=当前点的压力-液柱压力+流体粘滞摩阻损耗的压力
                                jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - (jdsj_base(i + 1).chuishen - jdsj_base(i).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gn(i) * (jdsj_base(i + 1).xiashen - jdsj_base(i).xiashen) '单位MPa
                                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                            Else
                                jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                            End If
                        End If
                        '********************************************************************************************************************
                        '处理开关工具嘴损问题
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i + 1).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关工具,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关工具的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i + 1).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i + 1).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                End If
                                RECreader.Read()
                                '************************************************************************************************************************
                                ' 处理内外连通
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开油套" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    '********************************************************************************************************************
                                    ' 管内知，管外未知，由管内算管外
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 4
                                        If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    '********************************************************************************************************************
                                    ' 管外知，管内未知，由管外算管内
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 4
                                        If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 And gk_fdRow(0).Item("管内流体流向").ToString = "不流动" Then
                                        If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i + 1).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                '************************************************************************************************************************
                                ' 处理上下连通时嘴损问题
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开管内" Or RECreader.Item("开关状态").ToString = "关油套开管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                        If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                            If jdsj_base(i + 1).xingzhi = "节流工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If RECreader.Item("节流流向").ToString = "内外流通" Then
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                        If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i + 1).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                    End If
                End If
            Next i
            '**********************************************************************************************************************************
            ' 从井底（管内井底压力值处深度）由上向下递推管内压力；   用i-1推i，若i不通，则推不下去
            '**********************************************************************************************************************************
            For i = jd_count + 1 To total_jd - 1
                '************************************************************************************************************************
                '正向管内压力递推   用i-1推i，若i点不通，推不下去
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then '上面结点管内压力已知，下面结点管内压力未知
                    jd_gnyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl
                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).gn_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                        jd_gnyl_cont(i) = 1
                    Else
                        '不同深度节点，递推
                        '********************************************************************************************************************
                        '如果是管柱，对于关闭的开关元件，设置压力不连续
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关元件,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                Else
                                    RECreader.Read()
                                    If RECreader.Item("开关状态").ToString = "单关管内" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "关油套关管内" Then
                                        jd_gnyl_cont(i) = 0
                                        nwyl_cont_gn = False '管内上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                        If jd_gnyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底管压开关 等于 "计算"，且管内流体流向等于“向上”或“向下”，需要用摩阻计算方法计算管内流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gn(i) = dlt_ppm_gn_ndlt(gn_calmz_ltcs, jdsj_base(i).ygnj)
                            End If
                            '管内压力计算
                            '若高于管内液面，则将井口管压直接赋值
                            If jdsj_base(i).chuishen < gym_chuishen Then
                                jdsj_cacu(gk_gkxh - 1, i).gnyl = Val(gk_fdRow(0).Item("井口管压MPa").ToString) + jdsj_cacu(gk_gkxh - 1, i - 1).gn_Yeti_midu * 9.8 * jdsj_base(i).chuishen / 1000
                                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).gn_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                            Else
                                '管内液面位于两节点之间
                                If jdsj_base(i).chuishen > gym_chuishen And jdsj_base(i - 1).chuishen < gym_chuishen Then
                                    '液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl + jdsj_cacu(gk_gkxh - 1, i - 1).gn_Yeti_midu * 9.8 * (gym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - gym_chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gn(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("管液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                                Else
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gn(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl_ok
                                End If
                            End If
                        End If
                        '********************************************************************************************************************
                        '如果是管柱，处理打开的开关元件压力传递问题
                        '********************************************************************************************************************
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是开关元件,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                            '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                            If jdsj_base(i).xingzhi = "开关工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                End If
                                RECreader.Read()
                                '************************************************************************************************************************
                                '开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                                ' 对于开关元件
                                '若内外连通，
                                '   若管外压力未赋值，
                                '       （1）若流体流动，则考虑嘴损计算管外压力；
                                '       （2）若流体不流动，将管内压力赋值给管外压力；
                                '   若管外压力已赋值
                                '       判断压力差别是否符合哈数。
                                '若上下连通、流体流动，考虑嘴损修正计算点管内压力。
                                '************************************************************************************************************************
                                '************************************************************************************************************************
                                ' 处理内外连通
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开油套" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    '********************************************************************************************************************
                                    ' 管内知，管外未知，由管内算管外
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gwyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    '********************************************************************************************************************
                                    ' 管外知，管内未知，由管外算管内
                                    '********************************************************************************************************************
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                        End If
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 And gk_fdRow(0).Item("管内流体流向").ToString = "不流动" Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                '************************************************************************************************************************
                                ' 处理上下连通时嘴损问题
                                '************************************************************************************************************************
                                If RECreader.Item("开关状态").ToString = "单开管内" Or RECreader.Item("开关状态").ToString = "关油套开管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                        If gk_fdRow(0).Item("井底管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                            If jdsj_base(i).xingzhi = "节流工具" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If RECreader.Item("节流流向").ToString = "内外流通" Then
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i).gnyl
                                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gwyl
                                        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 4
                                    End If
                                    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok <> 0 Then
                                        If jdsj_cacu(gk_gkxh - 1, i).gnyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        ElseIf jdsj_cacu(gk_gkxh - 1, i).gwyl <> 0 Then
                                            If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gnyl - jdsj_cacu(gk_gkxh - 1, i).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i).gwyl) > 0.2 Then
                                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                                RECreader.Close()
                                                cn_userdb.Close()
                                                cn_userdb.Dispose()
                                                yeya = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If
                    End If
                End If
            Next i
        End If


        '**************************************************************************************************************************************
        ' 考虑到有开关工具将压力传到环空，故再次计算一遍环空压力
        '                                                                                         考虑循环解决--2024年07初闪念
        '**************************************************************************************************************************************
        '**************************************************************************************************************************************
        '   1.管外压力递推：
        '       （1）环空井口压力为输入时，从井口由上向下递推环空压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井口套压开关").ToString = "输入" Then
            '将井口套压直接赋值
            jdsj_cacu(gk_gkxh - 1, 0).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString)
            jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
            jdsj_cacu(gk_gkxh - 1, 0).gwyl_ok = 1
            For i = 1 To total_jd - 1
                '************************************************************************************************************************
                '正向管外压力递推
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                    '********************************************************************************************************************
                    '上面结点管外压力已知，下面结点管外压力未知，符合递推条件
                    '********************************************************************************************************************
                    jd_gwyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).hk_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = (jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok)
                        jd_gwyl_cont(i) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                If Not RECreader.HasRows Then
                                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i).ID & "号封隔器坐封状态未知，计算无法进行。请在工况数据输入界面确定！"
                                    RECreader.Close()
                                    cn_userdb.Close()
                                    cn_userdb.Dispose()
                                    yeya = False
                                    Exit Function
                                End If
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井底套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If
                            '环空压力计算
                            '若高于环空液面，则将井口环压直接赋值
                            If jdsj_base(i).chuishen < hym_chuishen Then
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString) + jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu * 9.8 * jdsj_base(i).chuishen / 1000
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = (jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok)
                            Else
                                '环空液面位于两节点之间
                                If jdsj_base(i).chuishen > hym_chuishen And jdsj_base(i - 1).chuishen < hym_chuishen Then
                                    '当前点液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + jdsj_cacu(gk_gkxh - 1, i - 1).hk_Yeti_midu * 9.8 * (hym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = (jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok)
                                Else
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = (jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok)
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        '**************************************************************************************************************************************
        '   1.管外压力递推：
        '       （2）环空井底压力为输入时，从井底（环空井底压力值处深度）由下向上递推环空压力，从井底（环空井底压力值处深度）由上向下递推环空压力；
        '**************************************************************************************************************************************
        If gk_fdRow(0).Item("井底套压开关").ToString = "输入" Then
            '**********************************************************************************************************************************
            ' 井底环空压力值对应深度处的垂深
            '**********************************************************************************************************************************
            If TSM_ver_switch = 1 Then
                bl_temp3 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
            Else
                bl_temp0 = Val(gk_fdRow(0).Item("套压井底深度m").ToString)
                Call cal_jx_fw_cs(bl_temp0, bl_temp1, bl_temp2, bl_temp3)
            End If
            '**********************************************************************************************************************************
            ' 找到环空压力值对应深度处的节点并给节点环空压力赋值
            '**********************************************************************************************************************************
            If Val(gk_fdRow(0).Item("套压井底深度m").ToString) > jdsj_base(total_jd - 1).xiashen Then
                jd_count = total_jd - 1
                jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString) - (bl_temp3 - jdsj_base(jd_count).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                jdsj_cacu(gk_gkxh - 1, jd_count).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
            Else
                For i = total_jd - 1 To 1 Step -1
                    If Val(gk_fdRow(0).Item("套压井底深度m").ToString) = jdsj_base(i).xiashen Then
                        jd_count = i
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
                        Exit For
                    End If
                    If Val(gk_fdRow(0).Item("套压井底深度m").ToString) < jdsj_base(i).xiashen And Val(gk_fdRow(0).Item("套压井底深度m").ToString) > jdsj_base(i - 1).xiashen Then
                        jd_count = i - 1
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl = Val(gk_fdRow(0).Item("井底套压MPa").ToString) - (bl_temp3 - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 '单位MPa
                        jdsj_cacu(gk_gkxh - 1, jd_count).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                        jdsj_cacu(gk_gkxh - 1, jd_count).gwyl_ok = 1
                    End If
                Next i
            End If
            '**********************************************************************************************************************************
            ' 从井底（环空井底压力值处深度）由下向上递推环空压力。
            '**********************************************************************************************************************************
            For i = jd_count - 1 To 0 Step -1
                jd_gwyl_cont(i + 1) = 1
                If jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then '下面结点管外压力已知，上面结点管外压力未知
                    If jdsj_base(i + 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                        jd_gwyl_cont(i + 1) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i + 1).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i + 1) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i + 1) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i + 1) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i + 1) = 1 Then
                            '********************************************************************************************************************
                            '如果：井口套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If
                            '管外液面位于两节点之间
                            If jdsj_base(i).chuishen >= hym_chuishen Then
                                '前一点的液压等于当前点的压力减去液柱压力
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - (jdsj_base(i + 1).chuishen - jdsj_base(i).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gw(i) * (jdsj_base(i + 1).xiashen - jdsj_base(i).xiashen) '单位MPa
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                            Else
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - (jdsj_base(i + 1).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gw(i) * (jdsj_base(i + 1).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                            End If
                        End If
                    End If
                End If
            Next i
            '**********************************************************************************************************************************
            ' 从井底（环空井底压力值处深度）由上向下递推环空压力；
            '**********************************************************************************************************************************
            For i = jd_count + 1 To total_jd - 1
                '************************************************************************************************************************
                '正向环空压力递推
                '************************************************************************************************************************
                If jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                    '********************************************************************************************************************
                    '上面结点管外压力已知，下面结点管外压力未知，符合递推条件
                    '********************************************************************************************************************
                    jd_gwyl_cont(i) = 1
                    If jdsj_base(i - 1).xiashen = jdsj_base(i).xiashen Then
                        '同深度节点，压力相等,直接赋值
                        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl
                        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).hk_Yeti_midu
                        'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                        jd_gwyl_cont(i) = 1
                    Else
                        '不同深度节点，递推
                        If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                            '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                            If jdsj_base(i).xingzhi = "封隔器" Then
                                cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                                cn_userdb.Open()
                                SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i).ID)
                                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                RECreader = EXECOleDbCommand.ExecuteReader()
                                EXECOleDbCommand.Dispose()
                                RECreader.Read()
                                If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                    jd_gwyl_cont(i) = 0
                                    '************************************************************************************************************
                                    '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                    '隔器处于洗井状态，压力连续、可递推。
                                    '************************************************************************************************************
                                    If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i).ID)
                                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                        RECreader2 = EXECOleDbCommand.ExecuteReader()
                                        EXECOleDbCommand.Dispose()
                                        If RECreader2.Read Then
                                            If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                                jd_gwyl_cont(i) = 1
                                            End If
                                        End If
                                        RECreader2.Close()
                                    End If
                                    If jd_gwyl_cont(i) = 0 Then
                                        nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                    End If
                                End If
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                            End If
                        End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                        If jd_gwyl_cont(i) = 1 Then
                            '********************************************************************************************************************
                            '如果：井底套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                            '********************************************************************************************************************
                            If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And (gk_fdRow(0).Item("环空流体流向").ToString = "向上" Or gk_fdRow(0).Item("环空流体流向").ToString = "向下") Then
                                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                                dlt_ppm_gw(i) = dlt_ppm_gw_ndlt(gw_calmz_ltcs, jdsj_base(i).tgnj, jdsj_base(i).ygwj)
                            End If
                            '环空压力计算
                            '若高于环空液面，则将井口环压直接赋值
                            If jdsj_base(i).chuishen < Val(gk_fdRow(0).Item("环液深度m").ToString) Then
                                jdsj_cacu(gk_gkxh - 1, i).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString)
                                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                                'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                            Else
                                '环空液面位于两节点之间
                                If jdsj_base(i).chuishen > hym_chuishen And jdsj_base(i - 1).chuishen < hym_chuishen Then
                                    '当前点液压=前一节点压力+液柱压力-流体粘滞摩阻损耗的压力
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu * 9.8 * (hym_chuishen - jdsj_base(i - 1).chuishen) / 1000 + (jdsj_base(i).chuishen - hym_chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - Val(gk_fdRow(0).Item("环液深度m").ToString)) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                Else
                                    '环空液面不位于两节点之间
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl + (jdsj_base(i).chuishen - jdsj_base(i - 1).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 - dlt_ppm_gw(i) * (jdsj_base(i).xiashen - jdsj_base(i - 1).xiashen) '单位MPa
                                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i - 1).gwyl_ok
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        '**************************************************************************************************************************************
        '   3.利用管柱最下端管内、管外压力连通的边界条件，对没有赋值的且压力连续的点用递推公式计算管内、环空压力：
        '       （1）从井底由下向上递推管内压力，环空压力。递推起始点上（管柱的最下端），若管内压力已算出、环空压力未算出，则令环空压力等于管内
        '       压力；若环空压力已算出、管内压力未算出，则令管内压力等于环空压力。
        '   反向搜索列表,对没有赋值的且压力连续的点用递推公式计算压力
        '**************************************************************************************************************************************
        '最下面管内、管外是联通的
        If jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl_ok = 0 Then
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl = jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, total_jd - 1).gn_Yeti_midu
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl_ok = 5
        End If
        If jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl_ok <> 0 Then
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl = jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, total_jd - 1).hk_Yeti_midu
            jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl_ok = 5
        End If
        For i = total_jd - 1 - 1 To 0 Step -1
            '************************************************************************************************************************
            '管内压力计算
            '************************************************************************************************************************
            If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then '下面结点管内压力已知，上面结点管内压力未知
                jd_gnyl_cont(i + 1) = 1
                If jdsj_base(i + 1).xiashen = jdsj_base(i).xiashen Then
                    '同深度节点，压力相等,直接赋值
                    jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                    jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                    'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                    jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                    jd_gnyl_cont(i + 1) = 1
                Else
                    '不同深度节点，递推
                    '如果是管柱，修正开关工具及封隔定位元件引起的压力不连续
                    If InStr(jdsj_base(i + 1).leixing, "管柱") <> 0 Then
                        '判断节点是否是开关工具,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                        '开关工具的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                        If jdsj_base(i + 1).xingzhi = "开关工具" Then
                            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                            cn_userdb.Open()
                            SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                            RECreader = EXECOleDbCommand.ExecuteReader()
                            EXECOleDbCommand.Dispose()
                            If Not RECreader.HasRows Then
                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i + 1).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                                yeya = False
                                Exit Function
                            Else
                                RECreader.Read()
                                If RECreader.Item("开关状态").ToString = "单关管内" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "关油套关管内" Then
                                    jd_gnyl_cont(i + 1) = 0
                                    nwyl_cont_gn = False '管内上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                End If
                            End If
                            RECreader.Close()
                            cn_userdb.Close()
                            cn_userdb.Dispose()
                        End If
                    End If
                    If jd_gnyl_cont(i + 1) = 1 Then
                        '********************************************************************************************************************
                        '如果：井底管压开关 等于 "计算"，且管内流体流向等于“向上”或“向下”，需要用摩阻计算方法计算管内流体粘滞摩阻压力梯度
                        '********************************************************************************************************************
                        If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                            '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                            dlt_ppm_gn(i) = dlt_ppm_gn_ndlt(gn_calmz_ltcs, jdsj_base(i).ygnj)
                        End If
                        If jdsj_base(i).chuishen >= gym_chuishen Then
                            '前一点的液压=当前点的压力-液柱压力+流体粘滞摩阻损耗的压力
                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - (jdsj_base(i + 1).chuishen - jdsj_base(i).chuishen) * Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gn(i) * (jdsj_base(i + 1).xiashen - jdsj_base(i).xiashen) '单位MPa
                            jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                            'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                            jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                        Else
                            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i - 1).gnyl
                            jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i - 1).gn_Yeti_midu
                            'jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 1
                            jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok
                        End If
                    End If
                    '********************************************************************************************************************
                    '处理开关工具嘴损问题
                    '********************************************************************************************************************
                    If InStr(jdsj_base(i + 1).leixing, "管柱") <> 0 Then
                        '判断节点是否是开关工具,若是，需判断开关情况，若是关闭，则压力不连续，上面管内压力的赋值错误，要更正
                        '开关工具的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
                        If jdsj_base(i + 1).xingzhi = "开关工具" Then
                            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                            cn_userdb.Open()
                            SQL_command = "select * from 工况_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                            RECreader = EXECOleDbCommand.ExecuteReader()
                            EXECOleDbCommand.Dispose()
                            If Not RECreader.HasRows Then
                                msg_prompt = "在" & CStr(gk_gkxh) & "号工况中找不到" & jdsj_base(i + 1).ID & "号开关工具的开关状态，请在工况输入界面中检查该工况开关工具状态。"
                                RECreader.Close()
                                cn_userdb.Close()
                                cn_userdb.Dispose()
                                yeya = False
                                Exit Function
                            End If
                            RECreader.Read()
                            '************************************************************************************************************************
                            ' 处理内外连通
                            '************************************************************************************************************************
                            If RECreader.Item("开关状态").ToString = "单开油套" Or RECreader.Item("开关状态").ToString = "开油套关管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                '********************************************************************************************************************
                                ' 管内知，管外未知，由管内算管外
                                '********************************************************************************************************************
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 0 Then
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                                    jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 4
                                    If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                    End If
                                    If gk_fdRow(0).Item("井口套压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                    End If
                                End If
                                '********************************************************************************************************************
                                ' 管外知，管内未知，由管外算管内
                                '********************************************************************************************************************
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 4
                                    If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                    End If
                                    If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                        jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl + Val(RECreader.Item("油套嘴损压差MPa").ToString)
                                    End If
                                End If
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 And gk_fdRow(0).Item("管内流体流向").ToString = "不流动" Then
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl <> 0 Then
                                        If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl) > 0.2 Then
                                            msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                            RECreader.Close()
                                            cn_userdb.Close()
                                            cn_userdb.Dispose()
                                            yeya = False
                                            Exit Function
                                        End If
                                    ElseIf jdsj_cacu(gk_gkxh - 1, i + 1).gwyl <> 0 Then
                                        If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) > 0.2 Then
                                            msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号开关工具内外连通，但管内外压力相差太大！"
                                            RECreader.Close()
                                            cn_userdb.Close()
                                            cn_userdb.Dispose()
                                            yeya = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            '************************************************************************************************************************
                            ' 处理上下连通时嘴损问题
                            '************************************************************************************************************************
                            If RECreader.Item("开关状态").ToString = "单开管内" Or RECreader.Item("开关状态").ToString = "关油套开管内" Or RECreader.Item("开关状态").ToString = "开油套开管内" Then
                                If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok <> 0 And (gk_fdRow(0).Item("管内流体流向").ToString = "向上" Or gk_fdRow(0).Item("管内流体流向").ToString = "向下") Then
                                    If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl - Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                    End If
                                    If gk_fdRow(0).Item("井口管压开关").ToString = "计算" And gk_fdRow(0).Item("管内流体流向").ToString = "向下" Then
                                        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 1, i).gnyl + Val(RECreader.Item("管内嘴损压差MPa").ToString)
                                    End If
                                End If
                            End If
                            RECreader.Close()
                            cn_userdb.Close()
                            cn_userdb.Dispose()
                        End If
                        If jdsj_base(i + 1).xingzhi = "节流工具" Then
                            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                            cn_userdb.Open()
                            SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                            RECreader = EXECOleDbCommand.ExecuteReader()
                            EXECOleDbCommand.Dispose()
                            RECreader.Read()
                            If RECreader.Item("节流流向").ToString = "内外流通" Then
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 0 Then
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gnyl
                                    jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok = 4
                                End If
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gnyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                                    jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok = 4
                                End If
                                If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 Then
                                    If jdsj_cacu(gk_gkxh - 1, i + 1).gnyl <> 0 Then
                                        If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl) > 0.2 Then
                                            msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                            RECreader.Close()
                                            cn_userdb.Close()
                                            cn_userdb.Dispose()
                                            yeya = False
                                            Exit Function
                                        End If
                                    ElseIf jdsj_cacu(gk_gkxh - 1, i + 1).gwyl <> 0 Then
                                        If System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gnyl - jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) / System.Math.Abs(jdsj_cacu(gk_gkxh - 1, i + 1).gwyl) > 0.2 Then
                                            msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，" & jdsj_base(i + 1).ID & "号节流工具内外连通，但管内外压力相差太大！"
                                            RECreader.Close()
                                            cn_userdb.Close()
                                            cn_userdb.Dispose()
                                            yeya = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            RECreader.Close()
                            cn_userdb.Close()
                            cn_userdb.Dispose()
                        End If
                    End If
                End If
            End If
            '************************************************************************************************************************
            '管外压力计算
            '************************************************************************************************************************
            If jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok <> 0 And jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then '下面结点管外压力已知，上面结点管外压力未知
                jd_gwyl_cont(i + 1) = 1
                If jdsj_base(i + 1).xiashen = jdsj_base(i).xiashen Then
                    '同深度节点，压力相等,直接赋值
                    jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl
                    jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, i + 1).hk_Yeti_midu
                    'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                    jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                    jd_gwyl_cont(i + 1) = 1
                Else
                    '不同深度节点，递推
                    If InStr(jdsj_base(i).leixing, "管柱") <> 0 Then
                        '判断节点是否是封隔定位元件,若是，需判断坐封情况，若是坐封，则环空压力不连续，上面的环空压力赋值错误，要更正
                        If jdsj_base(i + 1).xingzhi = "封隔器" Then
                            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
                            cn_userdb.Open()
                            SQL_command = "select * from 工况_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & " and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                            RECreader = EXECOleDbCommand.ExecuteReader()
                            EXECOleDbCommand.Dispose()
                            RECreader.Read()
                            If Trim(RECreader.Item("封隔器状态").ToString) = "坐封" Then
                                jd_gwyl_cont(i + 1) = 0
                                '************************************************************************************************************
                                '    如果封隔器坐封、环空流体“向下”流动、管内流体“向上”流动、封隔器：能否反洗井=“能”，判断为可反洗井封
                                '隔器处于洗井状态，压力连续、可递推。
                                '************************************************************************************************************
                                If gk_fdRow(0).Item("环空流体流向").ToString = "向下" And gk_fdRow(0).Item("管内流体流向").ToString = "向上" Then
                                    SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  元件序号=" & CStr(jdsj_base(i + 1).ID)
                                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                                    RECreader2 = EXECOleDbCommand.ExecuteReader()
                                    EXECOleDbCommand.Dispose()
                                    If RECreader2.Read Then
                                        If Trim(RECreader2.Item("能否反洗井").ToString) = "能" Then
                                            jd_gwyl_cont(i + 1) = 1
                                        End If
                                    End If
                                    RECreader2.Close()
                                End If
                                If jd_gwyl_cont(i + 1) = 0 Then
                                    nwyl_cont_hk = False '环空上下畅通，具备管内与管外在底部形成连通器条件，此变量为真
                                End If
                            End If
                        End If
                    End If 'If InStr(jdsj_base(i).leixing, "管柱") <> 0
                    If jd_gwyl_cont(i + 1) = 1 Then
                        '********************************************************************************************************************
                        '如果：井口套压开关 等于 "计算"，且环空流体流向等于“向上”或“向下”，需要用摩阻计算方法计算环空流体粘滞摩阻压力梯度
                        '环空流体粘滞摩阻压力梯度已在之前算出。
                        '********************************************************************************************************************
                        '管外液面位于两节点之间
                        If jdsj_base(i).chuishen > hym_chuishen Then
                            '前一点的液压等于当前点的压力减去液柱压力
                            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl - (jdsj_base(i + 1).chuishen - jdsj_base(i).chuishen) * Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString) * 9.8 / 1000 + dlt_ppm_gw(i) * (jdsj_base(i + 1).xiashen - jdsj_base(i).xiashen) '单位MPa
                            jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                            'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                            jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                        Else
                            jdsj_cacu(gk_gkxh - 1, i).gwyl = Val(gk_fdRow(0).Item("井口环压MPa").ToString)
                            jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 1, 0).hk_Yeti_midu
                            'jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 1
                            jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = jdsj_cacu(gk_gkxh - 1, i + 1).gwyl_ok
                        End If
                    End If
                End If
            End If
        Next i

        '**************************************************************************************************************************************
        '   4.给没有管内、环空压力赋值的节点赋值。
        '       从井口由上向下循环，对于未赋值节点，将地层压力赋给还没有算出的节点。
        '***********************************************************这里有问题，可以吗？*******************************************************
        '注：
        'gnyl_ok、gwyl_ok值若为1 表示根据已知条件递推赋值
        'gnyl_ok、gwyl_ok值若为2 表示没有能正常递推赋值，赋值地层压力
        'gnyl_ok、gwyl_ok值若为3 ，原来的设置，现放弃不用。原表示非初始工况，没有能正常递推赋值，赋值为上一工况的对应点的压力值
        'gnyl_ok、gwyl_ok值若为4 表示根据开关工具或节流工具内外连通条件递推赋值
        'gnyl_ok、gwyl_ok值若为5 表示根据井底内外连通条件递推赋值
        '**************************************************************************************************************************************
        For i = 0 To total_jd - 1
            If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then
                jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_base(i).chuishen * gk_dytd
                jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
                jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 2
            End If
            If jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
                jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_base(i).chuishen * gk_dytd
                jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
                jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 2
            End If

            '若是初始工况，则将地层压力赋给还没有算出的节点，若是其他工况，将上一工况的对应点的数据赋给还没有算出的节点
            'If gk_gkxh = 1 Then
            '    If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then
            '        jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_base(i).chuishen * gk_dytd
            '        jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = Val(gk_fdRow(0).Item("管液密度g╱cm3").ToString)
            '        jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 2
            '    End If
            '    If jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
            '        jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_base(i).chuishen * gk_dytd
            '        jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = Val(gk_fdRow(0).Item("环液密度g╱cm3").ToString)
            '        jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 2
            '    End If
            'Else
            '    If gk_gkxh > 1 Then
            '        If jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 0 Then
            '            jdsj_cacu(gk_gkxh - 1, i).gnyl = jdsj_cacu(gk_gkxh - 2, i).gnyl
            '            jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu = jdsj_cacu(gk_gkxh - 2, i).gn_Yeti_midu
            '            jdsj_cacu(gk_gkxh - 1, i).gnyl_ok = 3
            '        End If
            '        If jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 0 Then
            '            jdsj_cacu(gk_gkxh - 1, i).gwyl = jdsj_cacu(gk_gkxh - 2, i).gwyl
            '            jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu = jdsj_cacu(gk_gkxh - 2, i).hk_Yeti_midu
            '            jdsj_cacu(gk_gkxh - 1, i).gwyl_ok = 3
            '        End If
            '    End If
            'End If
        Next i
        '**************************************************************************************************************************************
        '判断在内外构成连通器的情况下，井低管内外压力是否接近，符合哈数
        '**************************************************************************************************************************************
        If nwyl_cont_hk = True And nwyl_cont_gn = True Then
            If jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl = 0 Then
                msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，井底管内压力居然为零！计算结果不符合流体力学基本原理，请检查该工况有关参数是否正确。"
                msg_buttons = 0 + 48
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                'yeya = False
                'Exit Function
            Else
                If System.Math.Abs((jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl - jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl) / jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl) > 0.05 Then
                    msg_prompt = "在" & CStr(gk_gkxh) & "号工况下，管内管外在井底连通，但算出的管内压力为" & Str(jdsj_cacu(gk_gkxh - 1, total_jd - 1).gnyl) _
                        & "MPa，环空压力为" & Str(jdsj_cacu(gk_gkxh - 1, total_jd - 1).gwyl) & "MPa，两者压力相差太远，计算结果不符合流体力学基本原理，请检查该工况有关参数是否正确。"
                    msg_buttons = 0 + 48
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    'yeya = False
                    'Exit Function
                End If
            End If
        End If
        '**************************************************************************************************************
        '填写节点计算参数表
        '**************************************************************************************************************
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        For i = 0 To total_jd - 1
            SQL_command = "update 节点计算参数表 set " _
                & "管内压力MPa=" & CStr(jdsj_cacu(gk_gkxh - 1, i).gnyl) & "," & "管内液压OK=" & CStr(jdsj_cacu(gk_gkxh - 1, i).gnyl_ok) & "," _
                & "管外压力MPa=" & CStr(jdsj_cacu(gk_gkxh - 1, i).gwyl) & "," & "管外液压OK=" & CStr(jdsj_cacu(gk_gkxh - 1, i).gwyl_ok) & "," _
                & "环液密度g╱cm3=" & CStr(jdsj_cacu(gk_gkxh - 1, i).hk_Yeti_midu) & "," & "管液密度g╱cm3=" & CStr(jdsj_cacu(gk_gkxh - 1, i).gn_Yeti_midu) _
                & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and  工况序号=" & CStr(gk_gkxh) & "and 节点编号=" & CStr(jdsj_base(i).bianhao)
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        Next i
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function

    '*************************************************************************************************************************************
    '                                                   用流体摩阻计算模型计算环空单位长度流体摩阻子程序
    '                                                                                              2022年7月20日秦彦斌最后修订整理
    '输入参数：
    '   gw_calmz_ltcs         流体摩阻计算模型计算所需环空流体参数
    '   tgnj                  套管内径  单位：mm
    '   ygwj                  油管外径  单位：mm
    '
    '全局变量使用：
    '   Global well_name As String          '当前井号
    '   Global zuoye_name As String         '当前管柱作业名称
    '   开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
    '
    ' 流体摩阻计算模型：无、牛顿流体模型、降阻比模型、幂律流体模型
    '
    '运行结果：返回环空单位长度流体摩阻值，单位：MPa
    '
    ' 重要修订记录:
    '    20220719 改工况记录行传递为结构体calmz_ltcs传递，以便通用。
    '*************************************************************************************************************************************
    Public Function dlt_ppm_gw_ndlt(ByVal gw_calmz_ltcs As calmz_ltcs, ByVal tgnj As Double, ByVal ygwj As Double) As Double
        Dim ltmj As Double '流通面积，单位：平方米
        Dim ltls As Double '流体流速，单位：米/秒
        Dim Re As Double '流动雷诺数
        Dim zlxs_lmd As Double '阻力系数
        Dim dlt_pn As Double '牛顿流体水的摩阻
        Dim jzb_xgm As Double '降阻比
        Dim lnjzb_xgm As Double '降阻比
        Dim GHPG As Double '稠化剂浓度 英制单位
        Dim Rec As Double '临界雷诺数
        Dim N As Double '流变指数
        Dim k As Double '稠度系数
        Dim xd_d As Double '当量直径 单位：米
        Dim a As Double '幂律流体计算系数 a
        Dim b As Double '幂律流体计算系数 b
        'On Error GoTo ErrHandler
        dlt_ppm_gw_ndlt = 0.0#
        '当量直径
        xd_d = tgnj / 1000.0# - ygwj / 1000.0#

        Select Case gw_calmz_ltcs.mozu_model
            Case "无"
                dlt_ppm_gw_ndlt = 0
            Case "牛顿流体模型"
                '(1) 计算雷诺数
                '（1）流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * ((tgnj / 1000.0#) ^ 2 - (ygwj / 1000.0#) ^ 2)
                '（2）流体流速，单位：米/秒
                ltls = gw_calmz_ltcs.liuliang / 60.0# / ltmj
                '（3）流动雷诺数，单位：
                Re = gw_calmz_ltcs.Yeti_midu * 1000 * ltls * xd_d / (gw_calmz_ltcs.Yeti_niandu / 1000.0#)
                '（4）根据雷诺数计算阻力系数
                If Re < 2300 Then
                    zlxs_lmd = 64.0# / Re
                ElseIf Re < 100000 Then
                    zlxs_lmd = 0.11 * (68.0# / Re + (12.5 * 0.000001) / xd_d) ^ 0.25
                Else
                    zlxs_lmd = 0.0032 + 0.221 / (Re ^ 0.237)
                End If
                zlxs_lmd = gw_calmz_ltcs.mozu_xishu * zlxs_lmd '算出的摩阻太大，减小处理
                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                dlt_ppm_gw_ndlt = 0.5 * 0.000001 * zlxs_lmd * gw_calmz_ltcs.Yeti_midu * 1000 * ltls ^ 2 / xd_d
                If gw_calmz_ltcs.liuxiang = "向上" Then
                    dlt_ppm_gw_ndlt = (-1) * dlt_ppm_gw_ndlt
                    'dlt_ppm_gw_ndlt = 0
                End If
            Case "降阻比模型"
                '(1) 计算雷诺数
                '（1）流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * ((tgnj / 1000.0#) ^ 2 - (ygwj / 1000.0#) ^ 2)
                '（2）流体流速，单位：米/秒
                ltls = gw_calmz_ltcs.liuliang / 60.0# / ltmj
                '（3）流动雷诺数，单位：
                Re = 1.0# * 1000 * ltls * xd_d / (1.0# / 1000.0#)
                '（4）根据雷诺数计算阻力系数
                If Re < 2300 Then
                    zlxs_lmd = 64.0# / Re
                ElseIf Re < 100000 Then
                    zlxs_lmd = 0.11 * (68.0# / Re + (12.5 * 0.000001) / xd_d) ^ 0.25
                Else
                    zlxs_lmd = 0.0032 + 0.221 / (Re ^ 0.237)
                End If
                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                dlt_pn = 0.5 * 0.000001 * zlxs_lmd * 1.0# * 1000 * ltls ^ 2 / xd_d
                If gw_calmz_ltcs.liuxiang = "向上" Then
                    dlt_pn = (-1) * dlt_pn
                    'dlt_pn = 0
                End If
                '****************************************************************************************************************
                'Lord 方法
                '****************************************************************************************************************
                '        GHPG = 11982.64 * val(gk_row.Item("环空稠剂浓度")
                '        lnjzb_xgm = 2.38 - 8.024 / (3.2808 * ltls) - 0.2365 * GHPG / (3.2808 * ltls) - 0.1639 * Log(GHPG) _
                ''                    - 0.028 * (11982.64 * val(gk_row.Item("环空撑剂浓度")) * Exp(1 / GHPG)
                '****************************************************************************************************************
                '郭建春, 罗波, 卢聪, 等 方法，见“试油完井系统流体压力分布计算方法与步骤王治国20150826.doc”
                '****************************************************************************************************************
                GHPG = gw_calmz_ltcs.chjnd
                lnjzb_xgm = 2.38 - 1.152 / 10000.0# * xd_d / gw_calmz_ltcs.liuliang - 0.2819 / 10000.0# * GHPG * xd_d / gw_calmz_ltcs.liuliang - 0.1639 * System.Math.Log(GHPG / 0.11983) - 2.3372 / 10000.0# * gw_calmz_ltcs.zcjnd * System.Math.Exp(0.11983 / GHPG)
                jzb_xgm = 1 / System.Math.Exp(lnjzb_xgm)
                dlt_ppm_gw_ndlt = dlt_pn * jzb_xgm
            Case "幂律流体模型"
                '****************************************************************************************************************
                '计算依据见“试油完井系统流体压力分布计算方法与步骤王治国20150826.doc”
                '****************************************************************************************************************
                '(1)计算幂律流体临界雷诺数
                N = gw_calmz_ltcs.lbzsh
                k = gw_calmz_ltcs.chdxsh
                Rec = (6464 * N * (2 + N) ^ ((2 + N) / 1 + N)) / ((1 + 3 * N) * (1 + 3 * N))
                '流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * ((tgnj / 1000.0#) ^ 2 - (ygwj / 1000.0#) ^ 2)
                '流体流速，单位：米/秒
                ltls = gw_calmz_ltcs.liuliang / 60.0# / ltmj
                '(2)计算幂律流体雷诺数
                Re = (gw_calmz_ltcs.Yeti_midu * 1000.0# * xd_d ^ N * ltls ^ (2 - N)) / (0.125 * k * ((6 * N + 2) / N) ^ N)
                If Re < Rec Then
                    zlxs_lmd = 4.0# * 16.0# / Re
                Else
                    a = 0.07763 * N ^ 0.10457
                    b = -4.0# * 0.16588 * System.Math.Exp(0.72485 / (N + 0.77577))
                    zlxs_lmd = a * Re ^ b
                End If
                dlt_ppm_gw_ndlt = 0.5 * 0.000001 * zlxs_lmd * gw_calmz_ltcs.Yeti_midu * 1000.0# * ltls ^ 2 / xd_d
        End Select
        '    Exit Function
        'ErrHandler:
        '    dlt_ppm_gw_ndlt = 0#
    End Function

    '*************************************************************************************************************************************
    '                                                   用流体摩阻计算模型计算管内单位长度流体摩阻子程序
    '                                                                                               2022年7月20日秦彦斌最后修订整理
    '输入参数：
    '   gn_calmz_ltcs         流体摩阻计算模型计算所需管内流体参数
    '   ygnj                  油管内径  单位：mm
    '
    '全局变量使用：
    '   Global well_name As String          '当前井号
    '   Global zuoye_name As String         '当前管柱作业名称
    '   开关元件的开关状态有8种：单开管内、单关管内、单开油套、单关油套、开油套关管内、关油套开管内、开油套开管内、关油套关管内
    '
    '   流体摩阻计算模型：无、牛顿流体模型、降阻比模型、幂律流体模型
    '
    '运行结果：返回环空单位长度流体摩阻值，单位：MPa
    '
    ' 重要修订记录:
    '    20220719 改工况记录行传递为结构体calmz_ltcs传递，以便通用。
    '*************************************************************************************************************************************
    Public Function dlt_ppm_gn_ndlt(ByVal gn_calmz_ltcs As calmz_ltcs, ByVal ygnj As Double) As Double
        Dim ltmj As Double '流通面积，单位：平方米
        Dim ltls As Double '流体流速，单位：米/秒
        Dim Re As Double '流动雷诺数
        Dim zlxs_lmd As Double '阻力系数
        Dim dlt_pn As Double '牛顿流体水的摩阻
        Dim jzb_xgm As Double '降阻比
        Dim lnjzb_xgm As Double '降阻比
        Dim GHPG As Double '稠化剂浓度 英制单位
        Dim Rec As Double '临界雷诺数
        Dim N As Double '流变指数
        Dim k As Double '稠度系数
        Dim d As Double '当量直径 单位：米
        Dim a As Double '幂律流体计算系数 a
        Dim b As Double '幂律流体计算系数 b
        dlt_ppm_gn_ndlt = 0.0#
        '当量直径
        d = ygnj / 1000.0#
        Select Case gn_calmz_ltcs.mozu_model
            Case "无"
                dlt_ppm_gn_ndlt = 0
            Case "牛顿流体模型"
                '(1) 计算雷诺数
                '（1）流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * d ^ 2
                '（2）流体流速，单位：米/秒
                ltls = gn_calmz_ltcs.liuliang / 60.0# / ltmj
                '（3）流动雷诺数，单位：
                Re = gn_calmz_ltcs.Yeti_midu * 1000.0# * ltls * d / (gn_calmz_ltcs.Yeti_niandu / 1000.0#)
                '（4）根据雷诺数计算阻力系数
                If Re < 2300 Then
                    zlxs_lmd = 64.0# / Re
                ElseIf Re < 100000 Then
                    zlxs_lmd = 0.11 * (68.0# / Re + (12.5 * 0.000001) / d) ^ 0.25
                Else
                    zlxs_lmd = 0.0032 + 0.221 / (Re ^ 0.237)
                End If
                zlxs_lmd = gn_calmz_ltcs.mozu_xishu * zlxs_lmd '算出的摩阻太大，减小处理
                '（5）流体粘滞摩阻压力梯度，单位：MPa/米
                dlt_ppm_gn_ndlt = 0.5 * 0.000001 * zlxs_lmd * gn_calmz_ltcs.Yeti_midu * 1000.0# * ltls ^ 2 / d
                If gn_calmz_ltcs.liuxiang = "向上" Then
                    '************************************************************************************************************
                    '开井工况用上述公式计算管内流体粘滞摩阻压力梯度非常大，不符合哈数。故暂时不予采纳
                    '************************************************************************************************************
                    dlt_ppm_gn_ndlt = (-1) * dlt_ppm_gn_ndlt
                    'dlt_ppm_gn_ndlt = 0
                End If
            Case "降阻比模型"
                '(1) 计算雷诺数
                '（1）流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * d ^ 2
                '（2）流体流速，单位：米/秒
                ltls = gn_calmz_ltcs.liuliang / 60.0# / ltmj
                '（3）清水流动雷诺数，单位：
                Re = 1.0# * 1000.0# * ltls * d / (1.0# / 1000.0#)
                '（4）根据雷诺数计算阻力系数
                If Re < 2300 Then
                    zlxs_lmd = 64.0# / Re
                ElseIf Re < 100000 Then
                    zlxs_lmd = 0.11 * (68.0# / Re + (12.5 * 0.000001) / (ygnj / 1000.0#)) ^ 0.25
                Else
                    zlxs_lmd = 0.0032 + 0.221 / (Re ^ 0.237)
                End If
                '（5）流体（清水）粘滞摩阻压力梯度，单位：MPa/米
                dlt_pn = 0.5 * 0.000001 * zlxs_lmd * 1.0# * 1000.0# * ltls ^ 2 / (ygnj / 1000.0#)
                If gn_calmz_ltcs.liuxiang = "向上" Then
                    '************************************************************************************************************
                    '开井工况用上述公式计算管内流体粘滞摩阻压力梯度非常大，不符合哈数。故暂时不予采纳
                    '************************************************************************************************************
                    dlt_pn = (-1) * dlt_pn
                    'dlt_pn = 0
                End If
                '****************************************************************************************************************
                'Lord 方法
                '****************************************************************************************************************
                'GHPG = 11982.64 * SQL_rst_dqgk.Fields("管内稠剂浓度")
                'lnjzb_xgm = 2.38 - 8.024 / (3.2808 * ltls) - 0.2365 * GHPG / (3.2808 * ltls) - 0.1639 * Log(GHPG) _
                ''            - 0.028 * (11982.64 * SQL_rst_dqgk.Fields("管内撑剂浓度")) * Exp(1 / GHPG)
                '****************************************************************************************************************
                '郭建春, 罗波, 卢聪, 等 方法，见“试油完井系统流体压力分布计算方法与步骤王治国20150826.doc”
                '****************************************************************************************************************
                GHPG = gn_calmz_ltcs.chjnd
                lnjzb_xgm = 2.38 - 1.152 / 10000.0# * ygnj / gn_calmz_ltcs.liuliang - 0.2819 / 10000.0# * GHPG * ygnj / gn_calmz_ltcs.liuliang - 0.1639 * System.Math.Log(GHPG / 0.11983) - 2.3372 / 10000.0# * gn_calmz_ltcs.zcjnd * System.Math.Exp(0.11983 / GHPG)
                jzb_xgm = 1 / System.Math.Exp(lnjzb_xgm)
                dlt_ppm_gn_ndlt = dlt_pn * jzb_xgm
            Case "幂律流体模型"
                '****************************************************************************************************************
                '计算依据见“试油完井系统流体压力分布计算方法与步骤王治国20150826.doc”
                '****************************************************************************************************************
                '(1)计算幂律流体临界雷诺数
                N = gn_calmz_ltcs.lbzsh
                k = gn_calmz_ltcs.chdxsh
                Rec = (6464 * N * (2 + N) ^ ((2 + N) / 1 + N)) / ((1 + 3 * N) * (1 + 3 * N))
                '流通面积，单位：平方米
                ltmj = 0.25 * 3.1415926535 * d ^ 2
                '流体流速，单位：米/秒
                ltls = gn_calmz_ltcs.liuliang / 60.0# / ltmj
                '(2)计算幂律流体雷诺数
                Re = (gn_calmz_ltcs.Yeti_midu * 1000.0# * d ^ N * ltls ^ (2 - N)) / (0.125 * k * ((6 * N + 2) / N) ^ N)
                If Re < Rec Then
                    zlxs_lmd = 4.0# * 16.0# / Re
                Else
                    a = 0.07763 * N ^ 0.10457
                    b = -4.0# * 0.16588 * System.Math.Exp(0.72485 / (N + 0.77577))
                    zlxs_lmd = a * Re ^ b
                End If
                dlt_ppm_gn_ndlt = 0.5 * 0.000001 * zlxs_lmd * gn_calmz_ltcs.Yeti_midu * 1000.0# * ltls ^ 2 / d
        End Select
    End Function
End Module