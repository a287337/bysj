Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Module main_all
    '*********************************************************************************************************************************************
    '                                                   关于公共主程序模块的说明
    ' main模块是软件运行的起始，主要进行以下工作
    '     （1）定义本软件用到的结构体数据
    '     （2）定义全局变量
    '     （3）说明外部动态链接库引用
    '     （4）部分全局变量赋值
    '     （5）检测是否有基础数据库
    '     （6）调用base_database_update()过程对基础数据库升级
    '     （7）软件信息全局变量赋值
    '     （8）显示登录界面，若登录失败，退出软件
    '
    ' 程序升级记事：
    '                                                                                           秦彦斌 2020年01月08日最后整理更新
    '     （1）20190717-从VB6升级到VS2008基本完成，尚有几个UPGRADE_???。
    '     （2）由于VS新功能：模块结束后，所调用的表单也同时结束，不能使主菜单处于等待事件的状态。故项目起始从起始菜单“zct_???”开始，在起始菜单中
    '  Call Main()调用，进行定义数据结构、定义全局变量等等共性操作。
    '     （3）
    '*********************************************************************************************************************************************

    '***********************************************************************************************************************
    '定义数据结构....
    '由于在处理数据输入和数据输出的时候，需要处理很多的变量，用自定义数据
    '结构的办法，可以使变量清晰，意义明确，处理方便。
    '***********************************************************************************************************************
    '1 管柱力学分析节点基本数据
    Structure jd_base
        Dim leixing As String '节点类型  TEXT(50),
        Dim ID As String '节点ID  TEXT(50) ,
        Dim xingzhi As String '节点性质  TEXT(30),
        Dim xiashen As Double '节点下深m Float,
        Dim bianhao As Integer '节点编号 long,
        Dim chuishen As Double '节点垂深m FLOAT,
        Dim jxj As Double '井斜角rad   FLOAT,
        Dim fwj As Double '方位角rad  FLOAT,
        Dim tgnj As Double '套管内径mm  FLOAT,
        Dim ygwj As Double '油管外径mm  FLOAT,
        Dim ygnj As Double '油管内径mm  FLOAT,
        Dim xzh As Double '线重kg╱m   float,
        Dim e As Double '弹性模量MPa FLOAT,
        Dim psb As Double '泊松比 FLOAT
        Dim xgm_b As Double '屈服强度
        Dim yg_gj As String '油管钢级
        Dim gtd As Double '曲率rad╱m FLOAT
        Dim csss_chng As Double '初始损伤长mm Float
        Dim csss_kuan As Double '初始损伤宽mm FLOAT
        Dim csss_shen As Double '初始损伤深mm FLOAT
        Dim csss_jiao As Double '初始损伤角° FLOAT
        Dim syqdxsh As Double '剩余强度系数 FLOAT
        Dim rzhxsh As Double   '热胀系数(/℃)
        Dim kjqdu As Double '抗挤强度MPa
        Dim gtkny As Double '管体抗内压MPa
        Dim jtkny As Double '接头抗内压MPa
        Dim gtkla As Double '管体抗拉kN
        Dim jtkla As Double '接头抗拉kN 
    End Structure
    '2 管柱力学分析节点力学分析数据
    Structure jd_cacu
        Dim gnyl As Double '节点管内压MPa  FLOAT,
        Dim gnyl_ok As Byte '管内液压OK  TINYINT ,
        Dim gwyl As Double '节点管外压MPa FLOAT,
        Dim gwyl_ok As Byte '管外液压OK TINYINT,
        Dim sj_zaihe As Double '节点管柱实载N FLOAT,
        Dim sj_zh_ok As Byte '管柱实载OK  TINYINT ,
        Dim dx_zaihe As Double '节点管柱等效载荷N  FLOAT,
        Dim dx_xczl As Double '节点等效悬持轴力
        Dim dx_zh_ok As Byte '管柱效载OK  TINYINT,
        Dim jchl_N As Double '接触力N ,
        Dim hwj_M As Double '合弯矩Nm  FLOAT,
        Dim hwj_ok As Byte '合弯矩OK TINYINT,
        Dim wendu As Double '节点温度   FLOAT,
        Dim wd_bx As Double '温度变形m  FLOAT
        Dim zl_bx As Double '轴力变形m FLOAT,
        Dim gz_bx As Double '鼓胀变形m FLOAT,
        Dim lx_bx As Double '螺旋变形m FLOAT,
        Dim zh_bx As Double '综合变形m FLOAT,
        Dim wd_xy As Double '温度变形效应m  FLOAT
        Dim zl_xy As Double '轴力变形效应m FLOAT,
        Dim gz_xy As Double '鼓胀变形效应m FLOAT,
        Dim lx_xy As Double '螺旋变形效应m FLOAT,
        Dim zh_xy As Double '综合变形效应m FLOAT,
        Dim xgm_xd4 As Double '合成应力MPa
        Dim aqxs_S As Double '安全系数
        Dim niuju As Double '扭矩Nm
        Dim TSM_OK As Byte 'TSM_OK    TINYINT
        Dim zzhpl As Double '纵振频率Hz    FLOAT
        Dim hzhpl As Double '横振频率Hz    FLOAT
        Dim zzhss_dep As Double '纵振损伤深mm  FLOAT
        Dim hzhss_dep As Double '横振损伤深mm  FLOAT
        Dim Fecrs As Double '正弦屈曲临界载荷    FLOAT
        Dim Fecrh As Double '螺旋屈曲临界载荷    FLOAT
        Dim mcxs As Double '油套摩擦系数  FLOAT
        Dim hk_Yeti_midu As Double   '环液密度g╱cm3 FLOAT
        Dim gn_Yeti_midu As Double   '管液密度g╱cm3 FLOAT
        Dim kjqd_p As Double '抗挤强度MPa FLOAT
    End Structure
    '3 修井管柱力学分析节点力学分析数据
    Structure jd_xiuj
        Dim gnyl As Double '管内压力MPa FLOAT,
        Dim gnyl_ok As Byte '管内液压OK  TINYINT ,
        Dim gwyl As Double '管外压力MPa FLOAT,
        Dim gwyl_ok As Byte '管外液压OK TINYINT,
        Dim sj_zaihe As Double '真实轴力N   FLOAT,
        Dim sj_zh_ok As Byte '管柱实载OK  TINYINT ,
        Dim dx_zaihe As Double '等效轴力N  FLOAT,
        Dim dx_zh_ok As Byte '等效轴力OK TINYINT,
        Dim dx_xczl As Double '等效悬持力N FLOAT,
        Dim sj_xczl As Double '真实悬持力N  FLOAT,
        Dim jchl_N As Double '接触力N     FLOAT,
        Dim klmcl As Double '库伦摩擦力N  FLOAT,
        Dim niuju As Double '扭矩Nm      FLOAT,
        Dim mcnj As Double '摩擦扭矩Nm  FLOAT,
        Dim hwj_M As Double '合弯矩Nm  FLOAT,
        Dim hwj_ok As Byte '合弯矩OK TINYINT,
        Dim wendu As Double '节点温度   FLOAT,
        Dim wd_bx As Double '温度变形m  FLOAT
        Dim zl_bx As Double '轴力变形m FLOAT,
        Dim gz_bx As Double '鼓胀变形m FLOAT,
        Dim lx_bx As Double '螺旋变形m FLOAT,
        Dim zh_bx As Double '综合变形m FLOAT,
        Dim xgm_xd4 As Double '合成应力MPa FLOAT,
        Dim aqxs_S As Double '安全系数
        Dim Fecrs As Double '正弦屈曲临界载荷    FLOAT
        Dim Fecrh As Double '螺旋屈曲临界载荷    FLOAT
        Dim mcxs As Double '油套摩擦系数  FLOAT
        Dim hk_Yeti_midu As Double   '环液密度g╱cm3 FLOAT
        Dim gn_Yeti_midu As Double   '管液密度g╱cm3 FLOAT
        Dim kjqd_p As Double '抗挤强度MPa FLOAT
    End Structure
    '摩阻摩矩及修井管柱力学分析用工况参数数据            20220630秦彦斌增加，如此调用修井（变长度）管柱力学分析计算子程序lxfx_XiuJingGZH.vb时，传入此类型参数，可减少函数调用参数个数
    Structure xj_gkcs
        Dim klmcxs_f As Double           '    （13）mcxs_f             Double     库伦摩擦系数                单位：               _Text3_8.Text = CStr(0.25)                    '库伦摩擦系数
        Dim string_length As Double      '    （02）string_length      Double     管柱下入长度                单位：m              _Text3_9.Text = CStr(qzh_length)              '管柱末端深度(m)
        Dim shang_wendu As Double        '    （09）shang_wendu        Double     管柱上端温度                单位：℃             _Text3_4.Text = CStr(20.0#)                   '管柱上端（井口）温度(℃)
        Dim xia_wendu As Double          '    （10）xia_wendu          Double     管柱下端温度                单位：℃             _Text3_5.Text = CStr(75.0#)                   '管柱下端（井底）温度(℃)
        Dim jiazai_F As Double           '    （11）jiazai_F           Double     井口悬重上加力              单位：kN             _Text3_6.Text = CStr(0.0#)                    '管柱上轴向力加载情况
        Dim jiazai_T As Double           '    （12）jiazai_T           Double     井口加扭                    单位：Nm             _Text3_7.Text = CStr(0.0#)                    '管柱上扭矩加载情况
        Dim hk_liuxiang As String        '                                                                                         ComboBox1.Text = "不流动"                     '环空流体流向(不流动、向下、向上)
        Dim gn_liuxiang As String        '    （15）xh_fx              string     循环方向    （正循环、反循环、不循环）           Combo2.Text = "不流动"                        '管内流体流向(不流动、向下、向上)
        Dim hk_jkyl As Double            '                                                                                         TextBox5.Text = "0"                           '井口环空压力(MPa)
        Dim hk_jkyl_fs As String         '                                                                                         Combo5.Text = "输入"                          '井口环空压力(MPa)取得方式(输入、计算）
        Dim gn_jkyl As Double            '    （06）pump_P             Double     进口泵压                    单位：MPa            _Text3_3.Text = CStr(20.0#)                   '井口管内压力(MPa)
        Dim gn_jkyl_fs As String         '                                                                                         Combo9.Text = "输入"                          '井口管内压力(MPa)取得方式(输入、计算）
        Dim hk_jdyl As Double            '                                                                                         TextBox6.Text = "0"                           '井底环空压力(MPa)
        Dim hk_jdyl_fs As String         '                                                                                         Combo8.Text = "计算"                          '井底环空压力(MPa)取得方式输入、计算）
        Dim gn_jdyl As Double            '                                                                                         TextBox7.Text = "0"                           '井底管内压力(MPa)
        Dim gn_jdyl_fs As String         '                                                                                         Combo6.Text = "计算"                          '井底管内压力(MPa)取得方式输入、计算）
        Dim hk_jd_shendu As Double       '                                                                                         TextBox8.Text = _Text3_9.Text                 '井底环压（温度）对应深度(m)
        Dim gn_jd_shendu As Double       '                                                                                         TextBox9.Text = _Text3_9.Text                 '井底管压（温度）对应深度(m)
        Dim hk_YM_shendu As Double       '                                                                                         TextBox10.Text = "0"                          '环空液面深度(m)
        Dim gn_YM_shendu As Double       '    （03）YM_depth           Double     液面深度                    单位：m              _Text3_11.Text = CStr(0.0#)                   '管内液面深度(m)
        Dim hk_Yeti_midu As Double       '                                                                                         TextBox11.Text = CStr(1.0#)                   '环空流体密度(g/cm^3)
        Dim gn_Yeti_midu As Double       '    （04）Yeti_midu          Double     流体密度                    单位：g/cm^3         _Text3_1.Text = CStr(1.0#)                    '管内流体密度(g/cm^3)
        Dim hk_mozu_model As String      '                                                                                         ComboBox2.Text = "牛顿流体模型"               '环空流体摩阻计算模型
        Dim gn_mozu_model As String      '    （16）mozu_model         string     流体摩阻计算模型   （无、牛顿流体模型）          Combo1.Text = "牛顿流体模型"                  '管内流体摩阻计算模型
        Dim hk_liuliang As Double        '                                                                                         TextBox13.Text = CStr(0.5)                    '环空流体流量(m^3/min)
        Dim gn_liuliang As Double        '    （07）liuliang           Double     流体流量                    单位：m^3/min        _Text3_12.Text = CStr(0.5)                    '管内流体流量(m^3/min)
        Dim hk_Yeti_niandu As Double     '                                                                                         TextBox14.Text = CStr(0.9)                    '环空流体动力粘度(mPa·s)
        Dim gn_Yeti_niandu As Double     '    （05）Yeti_niandu        Double     流体粘度                    单位：mPa·s         _Text3_2.Text = CStr(0.9)                     '管内流体动力粘度(mPa·s)
        Dim hk_mozu_xishu As Double      '                                                                                         TextBox15.Text = CStr(0.25)                   '环空牛模摩阻折减系数
        Dim gn_mozu_xishu As Double      '    （08）mozu_xishu         Double     摩阻折减系数                单位：               _Text3_0.Text = CStr(0.25)                    '管内牛模摩阻折减系数
        Dim hk_chjnd As Double           '                                                                                         TextBox16.Text = "0"                          '环空稠化剂浓度(kg/m^3)
        Dim gn_chjnd As Double           '                                                                                         TextBox12.Text = "0"                          '管内稠化剂浓度(kg/m^3)
        Dim hk_zcjnd As Double           '                                                                                         TextBox17.Text = "0"                          '环空支撑剂浓度(kg/m^3)
        Dim gn_zcjnd As Double           '                                                                                         TextBox20.Text = "0"                          '管内支撑剂浓度(kg/m^3)
        Dim hk_lbzsh As Double           '                                                                                         TextBox18.Text = "0"                          '环空流体流变指数n
        Dim gn_lbzsh As Double           '                                                                                         TextBox21.Text = "0"                          '管内流体流变指数n
        Dim hk_chdxsh As Double          '                                                                                         TextBox19.Text = "0"                          '环流稠度系数K(Pa·s^n)
        Dim gn_chdxsh As Double          '                                                                                         TextBox22.Text = "0"                          '管流稠度系数K(Pa·s^n)
        Dim ld_time As Double            '                                                                                         TextBox23.Text = "0"                          '流体流动时间(h)
        Dim jzfs_zhouli As String        '管柱上轴向力加载情况，值取 "井口加力" 或 "井底余力"，若为"井口加力"，已知井口悬重上加力(kN)，计算末端轴力(kN)； 若为"井底余力"，已知末端轴力(kN)，计算井口悬重上加力(kN)。
        '                                     （17）zhouli_model       Boolean    轴力计算方式开关，若为真，已知井口悬重上加力(kN)，计算末端轴力(kN)；若为假，已知末端轴力(kN)，计算井口悬重上加力(kN)。
        Dim jzfs_niuju As String         '管柱上扭矩加载情况，值取 "井口扭矩" 或"井底扭矩"，若为"井口扭矩"，已知井口扭矩(Nm)，计算末端扭矩(Nm)；若为"井底扭矩"，已知末端扭矩(Nm)，计算知井口扭矩(Nm)。
        '                                     （18）niuju_model        Boolean    扭矩计算方式开关，若为真，已知井口扭矩(Nm)，计算末端扭矩(Nm)；若为假，已知末端扭矩(Nm)，计算知井口扭矩(Nm)。
        '                                     （14）kzh_aqxs           Double     控制安全系数                单位：无
        'Dim jkxy_enable As String        '井口能否加下压力，值取"井口不能加下压力"或"井口可以加下压力"。用于管柱下入性分析
    End Structure
    '流体摩阻计算模型计算所需流体参数
    Structure calmz_ltcs
        Dim mozu_model As String      '流体摩阻计算模型
        Dim liuliang As Double        '流体流量(m^3/min)
        Dim Yeti_midu As Double       '流体密度(g/cm^3)
        Dim Yeti_niandu As Double     '流体动力粘度(mPa·s)
        Dim mozu_xishu As Double      '牛模摩阻折减系数
        Dim liuxiang As String        '流体流向(不流动、向下、向上)
        Dim chjnd As Double           '稠化剂浓度(kg/m^3)
        Dim zcjnd As Double           '支撑剂浓度(kg/m^3)
        Dim lbzsh As Double           '流体流变指数n
        Dim chdxsh As Double          '稠度系数K(Pa·s^n)
    End Structure

    '4 管柱力学分析节点摩阻摩矩分析数据
    'Structure jd_mzmj
    '    Dim dy_cta As Double '单元全角变化   FLOAT,
    '    Dim xczl_F As Double '悬持时的轴力 N  FLOAT,
    '    Dim stzl_F As Double '上提时的轴力 N  FLOAT,
    '    Dim xfzl_F As Double '下放时的轴力 N  FLOAT,
    '    Dim xczl_RF As Double '悬持时的真实轴力 N  FLOAT,
    '    Dim stzl_RF As Double '上提时的真实轴力 N  FLOAT,
    '    Dim xfzl_RF As Double '下放时的真实轴力 N  FLOAT,
    '    Dim stzy_N As Double '上提时的侧向支撑力 N  FLOAT,
    '    Dim xfzy_N As Double '下放时的侧向支撑力 N  FLOAT,
    '    Dim xzmj_T As Double '旋转摩擦扭矩 N.m  FLOAT,
    'End Structure
    ' 钻井日志数据
    Structure zjrzh
        Dim xuhao As Integer '序号 long
        Dim qishen As Double '起深 FLOAT
        Dim jinchi As Double '进尺 FLOAT
        Dim cztime As Double '纯钻时间 FLOAT
        Dim zhuansu As Double '转速 FLOAT
        Dim zuanya As Double '钻压 FLOAT
        Dim mud_type As String '泥浆类型 TEXT(30)
        Dim mud_midu As Double '泥浆密度 FLOAT
        Dim hkytime As Double '划扩眼时间 FLOAT
        Dim bengya As Double '泵压 FLOAT
        Dim ygjs As Double '已固井深 FLOAT
    End Structure
    ' 套管磨损计算节点数据
    Structure tgmsjd
        Dim xuhao As Integer '序号 long
        Dim jsh As Double '井深m FLOAT
        Dim csh As Integer '层数 long
        Dim dsh As Integer '段数 long
        Dim hj As String '是否回接   TEXT(50)
        Dim tgwj As Double '套管外径mm FLOAT
        Dim tgbh As Double '套管壁厚mm FLOAT
        Dim xgsd As Double '悬挂深度m   FLOAT,
        Dim tgxsh As Double '套管下深m     FLOAT,
        Dim wzshd As Double '完钻深度m     FLOAT,
        Dim qd_kny As Double '抗内压MPa  FLOAT,
        Dim qd_kj As Double '抗挤强度MPa FLOAT
        Dim qd_kla As Double '抗拉强度kN FLOAT,
        Dim qfjx As Double '屈服极限MPa FLOAT
        Dim gtd As Double '狗腿度        FLOAT,
        Dim msshd As Double '磨损深度mm    FLOAT
        Dim msmj As Double '磨损面积m2    FLOAT
        Dim syqd_kj As Double '剩余抗挤强度MPa FLOAT
        Dim syqd_kny As Double '剩余抗内压强度MPa FLOAT
        Dim syqd_kl As Double '剩余抗拉强度kN FLOAT)
        Dim BHA_wj As Double '钻杆接头外径，song公式要用到，用接触过的钻杆外径平均值
        Dim BHA_cishu As Double '钻杆接头外径累加次数
        Dim tggj As String '套管钢级
        Dim max_N As Double '最大侧向力
        Dim BHA_mstime As Double '钻杆与套管磨损时间
    End Structure
    ' 套管磨损节点临时表数据
    Structure tgms_ls
        Dim xuhao As Integer '序号 long
        Dim jsh As Double '井深m FLOAT
        Dim csh As Integer '层数 long
        Dim dsh As Integer '段数 long
        Dim tgwj As Double '套管外径mm FLOAT
        Dim tgbh As Double '套管壁厚mm FLOAT
        Dim gtd As Double '狗腿度        FLOAT,
        Dim msshd As Double '磨损深度mm    FLOAT
        Dim msmj As Double '磨损面积m2    FLOAT
        Dim BHA_wj As Double '钻杆接头外径，song公式要用到，用接触过的钻杆外径平均值
        Dim BHA_cishu As Double '钻杆接头外径累加次数
        Dim tggj As String '套管钢级
        Dim qfjx As Double '屈服极限MPa FLOAT
        Dim max_N As String '最大侧向力
        Dim BHA_mstime As Double '钻杆与套管磨损时间
    End Structure
    '钻具数据数据
    Structure BHA
        Dim zh_xh As Integer '组合序号 long
        Dim yj_xh As Integer '元件序号 long
        Dim jsh As Double '井深m FLOAT
        Dim num_yj As Integer '元件个数
        Dim waijing As Double '外径
        Dim neijing As Double '内径
        Dim a_changdu As Double '单根长度m
        Dim zl_kgpm As Double '单位长度质量kgpm
    End Structure
    '套管磨损分析中钻具力学分析节点数据
    Structure BHA_jd
        Dim jd_xsh As Double '节点下深m
        Dim jd_csh As Double '节点垂深m
        Dim jd_jxj As Double '节点处井斜角rad
        Dim jd_fwj As Double '节点处方位角rad
        Dim jd_gtd As Double '节点处狗腿度
        Dim jd_tgnj As Double '套管或井壁内径mm
        Dim jd_BHAwj As Double '钻柱外径mm
        Dim jd_BHAnj As Double '钻柱内径mm
        Dim zl_kgpm As Double '单位长度质量kgpm
        Dim Fcrh As Double '屈曲临界载荷
        Dim sj_zaihe As Double '钻柱节点实载N FLOAT,
        Dim dx_zaihe As Double '钻柱节点等效载荷N  FLOAT,
        Dim dx_xczl As Double '钻柱节点等效悬持轴力
        Dim jchl_N As Double '侧向力N ,
        Dim hwj_M As Double '合弯矩Nm  FLOAT,
        Dim znyl As Double '钻柱节点内压MPa  FLOAT,
        Dim zwyl As Double '钻柱节点外压MPa FLOAT,
    End Structure
    '套管磨损分析中磨损参数数据
    Structure tgfx_cs
        Dim csh As Integer '层数 long
        Dim dsh As Integer '段数 long
        Dim mcxs_f As Double '摩擦系数
        Dim msxlN_xz As Double '磨损效率修正系数
    End Structure
    ''套管磨损分析中 完井后起下钻统计表 数据
    'Type wjh_qxztj
    '    xh As Integer       '序号
    '    xrsd As Double      '下入深度m
    '    xrjymd As Double    '下入井液密度
    '    xrjylx As String    '下入井液类型
    '    qchchd As Double    '起出长度m  一般等于 下入深度m，未起出 等于0
    '    qchjymd As Double   '起出井液密度
    '    qchjylx As String   '起出井液类型
    'End Type

    '管柱冲蚀程度预测分析中冲蚀预测分析结果数据
    Structure ERO_cs
        Dim xuhao As Integer '序号 long
        Dim REO_begin As Integer '冲蚀段起深m  FLOAT
        Dim REO_end As Double '冲蚀段止深m  FLOAT
        Dim tg_csh As Integer '套管层数 long
        Dim tg_dsh As Integer '套管段数 long
        Dim tg_wj As Double '套管外径mm FLOAT
        Dim tg_bihou As Double '套管初始壁厚mm FLOAT
        Dim tg_gangji As String '套管钢级
        Dim tg_ERO_hd As Double '套管冲蚀深度mm FLOAT
        Dim gz_zyname As String '管柱作业名称 TEXT(50)
        Dim gz_yjxuhao As Integer '管柱元件序号 long
        Dim gz_yjname As String '管柱元件名称 TEXT(50)
        Dim gz_yj_xzh As String '管柱元件性质 TEXT(30)
        Dim gz_waijing As Double '管柱元件外径mm FLOAT
        Dim gz_neijing As Double '管柱元件内径mm FLOAT
        Dim gz_gangji As String '管柱元件钢级 TEXT(30)
        Dim gz_gwydu As Double '管柱元件外硬度 FLOAT
        Dim gz_gnydu As Double '管柱元件内硬度 FLOAT
        Dim gz_gw_ERO As Double '管外冲蚀深度mm FLOAT
        Dim gz_gn_ERO As Double '管内冲蚀深度mm FLOAT
    End Structure
    '***********************************************************************************************************************
    '用于射孔段管柱安全性评价以及带井斜的管柱图绘制所需的结构体和函数    李润洲   2025年10月22日
    '***********************************************************************************************************************
    '射孔段管柱计算参数结构体
    Public Structure skdGZPars
        Public gkxh As Integer
        Public gkmc As String
        Public has_zf_fgq As Boolean
        '2025年7月26日，将tgwj、tgbh换为jyzj（井眼直径mm）
        'Public tgwj As Single
        'Public tgbh As Single
        Public jyzj As Single  '井眼直径mm
        '-------------------------- 射孔枪弹
        Public skqxh As Integer
        Public skqName As String
        Public zym As String
        Public zymd As Single
        Public zyl As Single
        Public km As Single
        Public skdName As String
        '------------------------------
        Public zyhot As Single '炸药物质生成热
        '------------------------------
        Public C_shu As Single   '炸药C数等
        Public H_shu As Single
        Public N_shu As Single
        Public O_shu As Single
        Public bw As Single     '爆温
        '------------------------------
        Public iskcd As Single  '射孔段长度m
        Public ifsjl As Single '封射距离m
        Public issjl As Single '筛射距离m
        Public ikdcd As Single '口袋长度m
        Public ijdwd As Single '井底温度摄氏度
        Public iybyl As Single '引爆压力MPa，也即井口加压
        Public icsyl As Single '井底初始压力MPa
        '------------------------------
        Public oBR As Single   '爆热
        Public oBRo As Single    '爆容
        Public obs As Single
        Public obw As Single
        Public oby As Single
        Public ofgyl As Single
        Public ojdyl As Single
        Public ojsdfz As Single

    End Structure

    '封隔器计算参数结构体
    Public Structure skdFgqPars
        Public gkxh As Integer
        Public gkmc As String
        '------------------------------钻油管
        Public ygxh As Integer
        Public ygwj As Single
        Public ygbh As Single
        Public yggg As String
        Public yjxz As String
        Public ygqfqd As Single
        '------------------------------封隔器
        Public fgq_xh As Integer
        Public fgqname As String
        Public zxgwj As Single
        Public zxgnj As Single
        Public zxgQfqd As Single
        '-----------------------------射孔工况
        Public jyzj As Single       '套管内径mm--->井眼直径mm,2025年7月
        Public icsyl As Single
        Public iybyl As Single
        Public ijdwd As Single
        Public ifsjl As Single
        Public skdcs As Single  '射孔段顶部垂深m
        Public skdcd As Single   '射孔段长度m
        Public kdcd As Single   '口袋长度m
        Public skymd As Single  '射孔液密度g/cm3
        Public jdfjfzyl As Single   '井底附加峰值应力MPa
        '------------------------------射孔枪弹
        Public skqxh As Integer
        Public skqName As String   '射孔枪名称
        Public skqnj As Single     '射孔枪内径mm
        Public km As Single     '射孔枪孔密 孔/m
        Public kj As Single     '射孔孔径mm
        Public zyl As Single   '装药量g
        Public zymd As Single  '装药密度g/cm3
        '------------------------------ 地层物性
        Public yskyqd As Single   '岩石抗压强度MPa
        Public ystxml As Single   '岩石弹性模量MPa
        Public dckxd As Single   '岩石孔隙度%
        Public dcstl As Single   '地层渗透率um2
        Public dcyl As Single    '地层压力MPa
        '------------------------------计算输出
        '------------------------------
        '射孔段的内压=井底初始压力+引爆压力+井底附加峰值压力
        Public skdny As Single
        '射孔后负压
        Public skhfy As Single
        '下拽力
        Public xzl As Single
        '------------------------------'封隔器中心管 
        Public zxgyl As Single
        Public zxgaqxs As Single
        '------------------------------管柱计算输出
        Public gz_xz_yl As Single   '下拽应力与安全系数
        Public gz_xz_aqxs As Single

        Public gz_sj_yl As Single
        Public gz_sj_aqxs As Single  '上挤应力与安全系数    
        '------------------------------
        '温度—地层参数校正后的输出
        Public wd_jz_qfqd As Single  '温度校正后的油管屈服强度MPa
        Public wd_jz_zxgqfqd As Single '温度校正后的油管屈服强度MPa

        Public wd_jz_sj_aqxs As Single  '温度校正后的管柱上挤安全系数
        Public wd_jz_xz_aqxs As Single  '温度校正后管柱下拽安全系数
        Public wd_jz_zxg_aqxs As Single '温度校正后中性安全系数

        Public dc_jz_xz_ygyl As Single  '地层参数校正后的油管下拽应力
        Public dc_jz_sj_ygyl As Single  '地层参数校正后的油管上挤应力
        Public dc_jz_zxgyl As Single   '地层参数校正后的中心管应力

        Public wd_dc_jz_xz_aqxs As Single  '温度_地层参数校正后的下拽油管安全系数
        Public wd_dc_jz_sj_aqxs As Single  '温度_地层参数校正后的上挤油管安全系数
        Public wd_dc_jz_zxgaqxs As Single  '温度_地层参数校正后的中心管安全系数
        '------------------------------
    End Structure
    '全井管柱应力参数结构体
    Public Structure qjGzYlPars
        '--------------------------------------元件
        Public yjxh As Integer '元件序号
        Public yjwj As Single  '外径mm
        Public yjmc As String
        Public yjnj As Single  '内径mm
        Public yjqfqd As Single  '屈服强度MPa
        '--------------------------------------工况参数
        Public gkxh As Integer
        Public gkmc As String
        Public iskymd As Single
        Public icsyl As Single
        Public obh_ydfjyl As Single   '井底爆轰附加压力MPa
        Public idzxs As Single
        '--------------------------------------计算输出
        Public qssd As Single
        Public zzsd As Single
        Public hsl As Single     '活塞力kN
        Public xsl As Single     '下甩力kN
        Public gzxz As Single    '管柱悬重kg
        Public gzfrz As Single     '浮容重kN
        Public xs_zzh As Single  '下甩时的总载荷kN
        Public yjyl As Single    '元件上部应力MPa
        Public ylaqxs As Single  '应力安全系数
        '--------------------------------------计算输出，2025年8月增加
        Public yjwzfjyl As Single '元件位置处的附加压力MPa
    End Structure
   

    '***********************************************************************************************************************
    '定义数据库及库表....
    '由于井斜、工况等数据在许多模块中使用，故设成全局变量，一次读取，多次使用
    '***********************************************************************************************************************
    Public dst As New DataSet("well_dst")
    Public Testwell_Table As DataTable = dst.Tables.Add("Testwell_Table")          '测井数据表               select  * from 测井数据表
    Public Ythshcsh_Table As DataTable = dst.Tables.Add("Ythshcsh_Table")          ' 三次井眼样条函数参数表

    Public gk_Table As DataTable = dst.Tables.Add("gk_Table")                      '工况参数表               select * from 工况参数表
    Public gk_xhmc_Table As DataTable = dst.Tables.Add("gk_xhmc_Table")            '工况参数表-              部分字段查询-工况序号、名称  "select 工况序号,工况名称 from 工况参数表
    'ipt_gkuang中用表
    Public gk_fgq_Table As DataTable = dst.Tables.Add("gk_fgq_Table")              '工况_封隔定位元件        部分字段查询-select 元件序号,元件名称,封隔器状态,定位方式 from 工况_封隔定位元件
    Public gk_mdgj_Table As DataTable = dst.Tables.Add("gk_mdgj_Table")            '工况_锚定元件            部分字段查询-select 元件序号,元件名称,定位方式  from 工况_锚定元件
    Public gk_kggj_Table As DataTable = dst.Tables.Add("gk_kggj_Table")            '工况_开关元件            部分字段查询-select 元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 工况_开关元件
    'FRM_TSM2中用表
    Public sp_gzh_Table As DataTable = dst.Tables.Add("sp_gzh_Table")              '管柱数据表-              字段重命名查询
    'Public gk_fgq_Table As DataTable = dst.Tables.Add("gk_fgq_Table")             '工况_封隔定位元件-       部分字段查询-select 元件序号,元件名称,封隔器状态,定位方式 from 工况_封隔定位元件
    'Public gk_mdgj_Table As DataTable = dst.Tables.Add("gk_mdgj_Table")           '工况_锚定元件-           部分字段查询-select 元件序号,元件名称,定位方式 " & " from 工况_锚定元件
    'Public gk_kggj_Table As DataTable = dst.Tables.Add("gk_kggj_Table")           '工况_开关元件-           部分字段查询-select 元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 工况_开关元件
    Public MaxMinCSTable As DataTable = dst.Tables.Add("MaxMinCSTable")            '节点计算参数表-          计算查询结果-计算参数极限值表，from 节点计算参数表
    Public gk_jd_jscs_Table As DataTable = dst.Tables.Add("gk_jd_jscs_Table")      '复合查询结果-选定工况不同节点计算参数列表，from 节点情况表,节点计算参数表
    Public gk_ssgzt_Table As DataTable = dst.Tables.Add("gk_ssgzt_Table")          '复合查询结果-选定工况伸缩管状态列表，from 工况_伸缩管状态 left Join 管柱数据表
    Public pg2_dxjdlb_Table As DataTable = dst.Tables.Add("pg2_dxjdlb_Table")      '复合查询结果-第2页即按节点显示页待选节点列表，from 节点情况表,节点计算参数表
    Public pg2_prtjdlb_Table As DataTable = dst.Tables.Add("pg2_prtjdlb_Table")    '复合查询结果-第2页即按节点显示页拟打印节点列表，from 节点情况表,节点计算参数表
    Public jd_gk_jscs_Table As DataTable = dst.Tables.Add("jd_gk_jscs_Table")      '复合查询结果-选定节点不同工况计算参数列表，from 工况参数表,节点情况表,节点计算参数表
    Public pg3_jdlb_Table As DataTable = dst.Tables.Add("pg3_jdlb_Table")          '复合查询结果-第3页各节点列表，from 节点情况表,节点计算参数表
    'lxfx_guanzhu中用表
    Public gz_shsg_Table As DataTable = dst.Tables.Add("gz_shsg_Table")            '管柱_伸缩元件            select * from 管柱_伸缩元件
    Public gk_fgqzt_Table As DataTable = dst.Tables.Add("gk_fgqzt_Table")          '工况_封隔定位元件        select * from 工况_封隔定位元件
    Public gk_mdgjzt_Table As DataTable = dst.Tables.Add("gk_mdgjzt_Table")        '工况_锚定元件            select * from 工况_锚定元件
    Public jdqk_Table As DataTable = dst.Tables.Add("jdqk_Table")                  '节点情况表               select * from 节点情况表
    Public yshu_Table As DataTable = dst.Tables.Add("yshu_Table")                  '约束点临时表             select * from 约束点临时表
    'ipt_gzhu中用表
    Public gzh_Table As DataTable = dst.Tables.Add("gzh_Table")                    '管柱数据表-              部分字段查询-select 元件序号,元件名称,元件性质,元件外径mm,元件内径mm,元件长度m from 管柱数据表
    'Frm_PRKM中用表
    Public gz_xhmc_Table As DataTable = dst.Tables.Add("gz_xhmc_Table")            '管柱数据表-              部分字段查询-select 元件序号,元件名称 from 管柱数据表
    Public fgqxfqx_Table As DataTable = dst.Tables.Add("fgqxfqx_Table")            '封隔器信封曲线参数表-    部分字段查询-select 序号,轴力kN,压差MPa from 封隔器信封曲线参数表 " _
    Public fgq_jd_Table As DataTable = dst.Tables.Add("fgq_jd_Table")              '复合查询结果-管柱中的封隔器，from 节点计算参数表 INNER JOIN 工况_封隔定位元件
    'Frm_KGGJM中用表
    'Public gz_xhmc_Table As DataTable = dst.Tables.Add("gz_xhmc_Table")           '管柱数据表-              部分字段查询-select 元件序号,元件名称 from 管柱数据表
    Public kggjxfqx_Table As DataTable = dst.Tables.Add("kggjxfqx_Table")          '开关元件载荷性能图参数表-部分字段查询-select 序号,轴力kN,压差MPa from 开关元件载荷性能图参数表
    Public kggj_jd_Table As DataTable = dst.Tables.Add("kggj_jd_Table")            '复合查询结果-管柱中的开关工具，from 节点计算参数表 INNER JOIN 工况_开关元件
    'Frm_MDGJM中用表
    'Public gz_xhmc_Table As DataTable = dst.Tables.Add("gz_xhmc_Table")           '管柱数据表-              部分字段查询-select 元件序号,元件名称 from 管柱数据表
    Public mdgjxfqx_Table As DataTable = dst.Tables.Add("mdgjxfqx_Table")          '锚定工具载荷性能图参数表-部分字段查询-sselect 序号,轴力kN,压差MPa from 锚定工具载荷性能图参数表
    Public mdgj_jd_Table As DataTable = dst.Tables.Add("mdgj_jd_Table")            '复合查询结果-管柱中的锚定工具，from 节点计算参数表 INNER JOIN 工况_锚定元件
    'L_JYGJ-井眼轨迹三维显示中用表
    Public JYGJ_3D_Table As DataTable = dst.Tables.Add("JYGJ_3D_Table")            '复合查询结果-井眼轨迹参数列表列表，from 测井数据表,井眼轨迹插值参数表
    'opt_xjzygzlxfx中用表
    Public MaxMinXJCSTable As DataTable = dst.Tables.Add("MaxMinXJCSTable")        '修井节点计算参数表-      计算查询结果-计算参数极限值表，from 修井节点计算参数表
    Public xjjd_jscs_Table As DataTable = dst.Tables.Add("xjjd_jscs_Table")        '复合查询结果-修井管柱力学分析不同节点计算参数列表，from  修井节点情况表, 修井节点计算参数表
    Public xj_kggj_Table As DataTable = dst.Tables.Add("xj_kggj_Table")            '修井摩阻分析_开关工具状态  部分字段查询-"select 元件序号,元件名称,开关状态,管内嘴损压差MPa,油套嘴损压差MPa from 修井摩阻分析_开关工具状态

    'M_lxfx_XiuJingGZH中用表
    Public xjjdqk_Table As DataTable = dst.Tables.Add("xjjdqk_Table")               '修井节点情况表            select * from 修井节点情况表
    'opt_kdwzhjs中用表
    Public gz_bx_Table As DataTable = dst.Tables.Add("gz_bx_Table")                 '提拉钩载-变形表
    Public kdfx_jdqk_Table As DataTable = dst.Tables.Add("kdfx_jdqk_Table")         '卡点分析节点情况表
    Public kdfxjg_table As DataTable = dst.Tables.Add("kdfxjg_table")               '卡点分析结果表
    'Frm_gzlsaqpj中用表
    Public zhl_lshaqxs_Table As DataTable = dst.Tables.Add("zhl_lshaqxs_Table")     '重量及拉伸应力安全系数   select * from 重量及拉伸应力安全系数
    'ipt_drilldata 中用表
    Public drilldata_Table As DataTable = dst.Tables.Add("drilldata_Table")         '钻井日志表               select  * from 钻井日志表
    'frmout_wear 中用表
    Public out_wear_Table As DataTable = dst.Tables.Add("out_wear_Table")           '磨损分析结果数据表       select * from 磨损分析结果数据表 where 井号='" & well_name & "' order by 井深m
    Public wear_canshu_Table As DataTable = dst.Tables.Add("wear_canshu_Table")     '套管磨损分析参数表       select 层,段,摩系系数,磨效系数 from  套管磨损分析参数表 where 井号='" & well_name & "' order by 层,段
    Public wear_zuoye_Table As DataTable = dst.Tables.Add("wear_zuoye_Table")       '磨损套管作业参数分析数据表  select * from 磨损套管作业参数分析数据表 where 井号='" & well_name & "'" & " order by 序号
    Public casing_Table As DataTable = dst.Tables.Add("casing_Table")               '套管数据表  "select * from 套管数据表 where 井号='" & well_name & "' order by 套管层数,套管段数"
    'mosunfx_taoguan.vb 中用表
    Public out_wear_Table2 As DataTable = dst.Tables.Add("out_wear_Table2")         '磨损分析结果数据表(数据导出用)       select * from 磨损分析结果数据表 where 井号='" & well_name & "' order by 井深m
    Public wearjd_ls_Table As DataTable = dst.Tables.Add("wearjd_ls_Table")         '磨损节点临时表            select * from 磨损节点临时表"
    Public BHA_Table As DataTable = dst.Tables.Add("BHA_Table")                     '组合序号表                select * from 钻具组合表"
    Public wjh_qxz_Table As DataTable = dst.Tables.Add("wjh_qxz_Table")             '完井后起下钻统计表        select * from 完井后起下钻统计表 where  井号='" & well_name & "' order by [序号]
    Public EBH_n_Table As DataTable = dst.Tables.Add("EBH_n_Table")                 '钻套磨损效率摩擦系数    从基础数据表中 select * from 钻套磨损效率摩擦系数 
    'frm_pjskd_l.vb 和frm_sk_qjzhyl.vb 中用表
    Public gz_yj_table As DataTable = dst.Tables.Add("gz_yj_table")                  '管柱元件起始、终止深度、重量等    从管柱数据表联合管柱_油井管等表 select * from 管柱数据表  李润洲2025年7月17日增加

    '宏狗运行需要的数据结构
    '***********************************************************************************************************************
    Structure RC_HARDWARE_INFO ' pre-defined type used by GrandDog Win32 API
        Dim lSerialNumber As Integer
        Dim lCurrentNumber As Integer
        Dim bDogType As Byte
        'Dim bDogModel(4) As Byte
        Dim bDogModel() As Byte
    End Structure
    Structure RC_ARRAY ' pre-defined type used by SafeSoft API
        'Dim buffer(1024) As Byte
        Dim buffer() As Byte
    End Structure
    '***********************************************************************************************************************
    '定义全局变量....
    '***********************************************************************************************************************
    Public dbname As String '基本数据数据库名称
    Public AdoConString As String '基本数据库ADO连接字符串
    Public use_dbname As String '用户数据库名称
    Public use_AdoConString As String '用户数据库ADO连接字符串
    Public sofe_name As String '软件名称全称
    Public sofe_name_l1 As String '软件名称第一行
    Public sofe_name_l2 As String '软件名称第二行
    Public sofe_danwei As String '软件单位全称
    Public sofe_danwei_l1 As String '软件单位第一行
    Public sofe_danwei_l2 As String '软件单位第二行
    Public sofe_ver As String '软件版本
    Public sofe_time As String '软件最后生成日期
    Public help_file As String '帮助文件名称及路径
    Public msg_prompt As String 'msgbox函数显示文本内容
    Public msg_buttons As Integer 'msgbox函数显示按钮形式
    Public msg_return As Integer 'msgbox函数返回值
    Public well_name As String '当前井号
    Public zuoye_name As String '当前管柱作业名称
    Public TSM_ver_switch As Integer '管柱力学分析软件版本控制开关。等于1时，垂直井节点法，井斜、方位均等于0，不判断、不读取井斜数据
    Public LoginSucceeded As Boolean '软件登录成功标志
    ' 宏狗判断需要的全局变量
    Public pDogHandle As Integer
    Public mPasswordType As Byte
    Public lOpenDogFlag As Integer
    Public bKeyType As Byte
    Public bDegree As Byte
    Public dog_prtsoftname As String
    Public dog_userpassword As String
    Public dog_msg_prompt1 As String
    Public dog_msg_prompt2 As String
    Public dog_msg_prompt3 As String
    Public dog_msg_prompt4 As String
    Public dog_msg_prompt5 As String
    Public dog_msg_prompt6 As String
    Public dog_msg_prompt7 As String
    Public dog_RetCode As Integer
    Public dog_DirID As Integer
    Public dog_FileID As Integer
    Public wdapp As Word.Application '定义word 应用程序变量
    Public wddoc As Word.Document '定义word 文档变量
    Public curAppPath As String
    Public userDbPath As String
    Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Integer)
    Declare Function HTMLhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Integer) As Integer
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Const SW_SHOWNORMAL As Short = 1
    Declare Function rc_OpenDog Lib "RCGrandDogW32.dll" (ByVal lFlag As Integer, ByVal pszProductName As String, ByRef pDogHandle As Integer) As Integer
    Declare Function rc_CloseDog Lib "RCGrandDogW32.dll" (ByVal lDogHandle As Integer) As Integer
    Declare Function rc_CheckDog Lib "RCGrandDogW32.dll" (ByVal lDogHandle As Integer) As Integer
    Declare Function rc_VerifyPassword Lib "RCGrandDogW32.dll" (ByVal lDogHandle As Integer, ByVal PasswordType As Byte, ByVal pszPassword As String, ByRef bDegree As Byte) As Integer
    Declare Function rc_VisitLicenseFile Lib "RCGrandDogW32.dll" (ByVal lDogHandle As Integer, ByVal iDirID As Integer, ByVal iFileID As Integer, ByVal lReserved As Integer) As Integer
    '*********************************************************************************************************************************************
    '对输入的文本框进行空、非数值判断
    '不通过判断时，返回-1
    '输入参数
    '   txtInpu               文本框名，如“Text1”
    '   prompt                文本框所输入值的名称，如“序号”
    '示例：
    '    If (isTextLegal(Text23, "序号")) = -1 Then Exit Sub
    '    If isTextLegal(Text24, "轴力") = -1 Then Exit Sub
    '    If isTextLegal(Text25, "压差") = -1 Then Exit Sub

    '*********************************************************************************************************************************************
    Public Function isTextLegal(ByVal txtInput As TextBox, ByVal prompt As String) As Short
        If Trim(txtInput.text) = "" Or Not IsNumeric(txtInput.text) Then
            msg_prompt = "请输入合法有效的" & prompt & "值。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, CType(msg_buttons, MsgBoxStyle), sofe_name)
            isTextLegal = -1
        Else
            isTextLegal = 0
        End If
    End Function
    '*********************************************************************************************************************************
    '                                                  主  程  序
    '此为主程序,所有全局变量在此赋初值,程序由这里开始执行!!!
    '1. 检查基本数据库文件及用户数据文件是否存在；
    '2. 建立用户数据库中必要的表。调用database_creat函数；
    '程序启动时，需确定本应用软件的路径，以及是否是第一次使用，如第一次使用应该检查一下软件的各个部分是否完整，是否存在一些错误。
    '如果不是第一次使用，也应该检查一下控制系统的数据文件是否正确，如不正确应该修改之，还应检查程序运行中的数据库的存在性和正确性。
    '*********************************************************************************************************************************
    Public Sub Main()
        Dim cn_basedb As System.Data.OleDb.OleDbConnection
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim SQL_command As String
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim rowfilter As String
        On Error GoTo errhandler
        '***************************************************************************************************************************
        '全局变量赋值
        '***************************************************************************************************************************
        well_name = ""
        zuoye_name = ""
        sofe_name = ""
        TSM_ver_switch = 0   '管柱力学分析软件版本控制开关。等于1时，垂直井节点法，井斜、方位均等于0，不判断、不读取井斜数据
        curAppPath = My.Application.Info.DirectoryPath
        '宏狗参数
        mPasswordType = 1 '用户密码验证
        bKeyType = 2 '
        bDegree = 0
        lOpenDogFlag = 1
        dog_RetCode = 1
        dog_msg_prompt1 = "没有发现合适的软件狗，核心计算将不能正常进行！"
        dog_msg_prompt2 = "软件狗密码有误！核心计算将不能正常进行！"
        dog_msg_prompt3 = "软件许可证时间不在有效期内，许可证已失效，核心计算不能正常进行！"
        dog_msg_prompt4 = "系统时间被篡改，许可证已失效，核心计算不能正常进行！"
        dog_msg_prompt5 = "许可证读取失败，核心计算不能正常进行！"
        dog_msg_prompt6 = "许可的使用次数已到，核心计算不能正常进行！"
        dog_msg_prompt7 = "许可的使用时间已到，核心计算不能正常进行！"
        '开发商工具中设定的参数： 产品名称
        dog_prtsoftname = "DHPS-ANS"
        '开发商工具中设定的参数： 用户口令
        dog_userpassword = "13909183925"
        '开发商工具中设定的参数： “狗”路径，不用设置
        dog_DirID = 16128
        '***********************************************************************************************************************
        '检测是否有基础数据库
        '***********************************************************************************************************************
        dbname = My.Application.Info.DirectoryPath & "\base_db.mdb"
        AdoConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbname & ";Persist Security Info=False"
        LoginSucceeded = False
        If Dir(dbname) = "" Then
            msg_prompt = "找不到基础数据库文件，程序结束。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, CType(msg_buttons, MsgBoxStyle), sofe_name)
            End
        End If
        '***********************************************************************************************************************
        '检测基础数据库中是否有软件信息表，若没有，结束软件运行
        '***********************************************************************************************************************
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        dbSchema = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='软件信息表'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length = 0 Then
            msg_prompt = "请升级基础数据库base_db.mdb，程序结束。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, CType(msg_buttons, MsgBoxStyle), sofe_name)
            dbSchema.Dispose()
            cn_basedb.Close()
            cn_basedb.Dispose()
            End
        End If
        dbSchema.Dispose()
        '***********************************************************************************************************************
        '软件信息赋值
        '***********************************************************************************************************************
        SQL_command = "select top 1 * from 软件信息表 where soft_carst=1"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.Read Then
            sofe_name = RECreader.Item("sofe_name").ToString()
            sofe_name_l1 = Trim(RECreader.Item("sofe_name_l1").ToString())
            sofe_name_l2 = Trim(RECreader.Item("sofe_name_l2").ToString())
            sofe_danwei = RECreader.Item("sofe_danwei").ToString()
            sofe_danwei_l1 = Trim(RECreader.Item("sofe_danwei_l1").ToString())
            sofe_danwei_l2 = Trim(RECreader.Item("sofe_danwei_l2").ToString())
            sofe_ver = Trim(RECreader.Item("sofe_ver").ToString())
            sofe_time = Trim(RECreader.Item("sofe_time").ToString())
            help_file = RECreader.Item("help_file").ToString()
        Else
            msg_prompt = "基础数据库信息错误，请升级基础数据库base_db.mdb，程序结束。"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, CType(msg_buttons, MsgBoxStyle), sofe_name)
            RECreader.Dispose()
            cn_basedb.Close()
            cn_basedb.Dispose()
            End
        End If
        RECreader.Dispose()
        RECreader.Close()
        cn_basedb.Close()
        cn_basedb.Dispose()

        '***********************************************************************************************************************
        '检测软件狗是否插上，李润洲2025年8月移到这里，先读取软件信息后，再判断软件狗
        '***********************************************************************************************************************
        dog_RetCode = rc_OpenDog(lOpenDogFlag, dog_prtsoftname, pDogHandle)
        If dog_RetCode <> 0 Then
            msg_buttons = 0 + 48
            msg_return = MsgBox(dog_msg_prompt1, CType(msg_buttons, MsgBoxStyle), sofe_name)
        End If
        dog_RetCode = rc_CloseDog(pDogHandle)

        '***********************************************************************************************************************
        '基础数据库升级
        '***********************************************************************************************************************
        Call base_database_update() '由于软件升级，数据结构改动，用程序自动升级相应数据表的数据结构。
        '***********************************************************************************************************************
        '设置帮助
        '***********************************************************************************************************************
        'UPGRADE_ISSUE: App 属性 App.HelpFile 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
        'App.HelpFile = My.Application.Info.DirectoryPath & "\" & help_file
        '***********************************************************************************************************************
        '显示登录界面，若登录失败，退出软件
        '***********************************************************************************************************************
        frmLogin1.ShowDialog()
        If LoginSucceeded = False Then
            End
        End If
        Exit Sub ' 退出程序，以避免进入错误处理程序。
errhandler:
        msg_prompt = "调用公共主程序时出错，可能是您的系统未装Microsoft Office Access，此为本软件运行必需，请确保安装了该软件！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
    End Sub

    '********************************************************************************************
    '生成汇总管柱元件起止深度、长度、重量的表  
    '用于射孔段管柱安全性评价
    '也用于管柱图绘制                             李润洲   2025年10月22日
    '********************************************************************************************
    Public Sub fill_gzyj_table()
        Dim SQL_command As String
        Dim gzyj_row As DataRow
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                '元件重量kg是整个长度的重量，不是单位长重量
                SQL_command = "select 元件序号,元件名称,元件性质,元件长度m,元件外径mm,元件内径mm,元件内径mm as 元件重量kg,元件长度m as 起始深度m, 元件长度m as 终止深度m,元件长度m as 屈服强度MPa from 管柱数据表 where 0>1"
                Using ad As New OleDbDataAdapter(SQL_command, cn_userdb)
                    gz_yj_table.Clear()
                    ad.Fill(gz_yj_table)
                End Using

                SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
                Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                    Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                        Dim s_length As Single = 0
                        While RECreader.Read
                            gzyj_row = gz_yj_table.NewRow()
                            gzyj_row.Item("元件序号") = RECreader.Item("元件序号").ToString()
                            gzyj_row.Item("元件名称") = RECreader.Item("元件名称").ToString()
                            gzyj_row.Item("元件性质") = RECreader.Item("元件性质").ToString()
                            gzyj_row.Item("元件长度m") = RECreader.Item("元件长度m").ToString()
                            gzyj_row.Item("元件外径mm") = RECreader.Item("元件外径mm").ToString()
                            gzyj_row.Item("元件内径mm") = RECreader.Item("元件内径mm").ToString()
                            gzyj_row.Item("起始深度m") = CStr(s_length)
                            s_length = s_length + Convert.ToSingle(RECreader.Item("元件长度m"))
                            gzyj_row.Item("终止深度m") = s_length.ToString()
                            Select Case RECreader.Item("元件性质").ToString()
                                Case "油井管"
                                    SQL_command = "select 单位长重kg╱m,屈服强度MPa from 管柱_油管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("单位长重kg╱m")) AndAlso Val(RECreader2.Item("单位长重kg╱m")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = (CSng(RECreader.Item("元件长度m")) * CSng(RECreader2.Item("单位长重kg╱m"))).ToString
                                                Else
                                                    Throw New InvalidDBException("油井管" & RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                If Not IsDBNull(RECreader2.Item("屈服强度MPa")) AndAlso Val(RECreader2.Item("屈服强度MPa")) <> 0 Then
                                                    gzyj_row.Item("屈服强度MPa") = RECreader2.Item("屈服强度MPa").ToString()
                                                Else
                                                    Throw New InvalidDBException("油井管" & RECreader.Item("元件名称").ToString() & "没有屈服强度参数")
                                                End If
                                            Else
                                                Throw New InvalidDBException("管柱数据没有建立" & RECreader.Item("元件名称").ToString() & "油井管")
                                            End If
                                        End Using
                                    End Using
                                Case "节流工具"
                                    SQL_command = "select 重量kg from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("重量kg")) AndAlso Val(RECreader2.Item("重量kg")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = RECreader2.Item("重量kg").ToString
                                                Else
                                                    Throw New InvalidDBException(RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                gzyj_row.Item("屈服强度MPa") = "0"
                                            Else
                                                Throw New InvalidDBException("管柱没有建立" & RECreader.Item("元件名称").ToString() & "数据")
                                            End If
                                        End Using
                                    End Using

                                Case "封隔器"
                                    SQL_command = "select 重量kg from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("重量kg")) AndAlso Val(RECreader2.Item("重量kg")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = RECreader2.Item("重量kg").ToString
                                                Else
                                                    Throw New InvalidDBException(RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                gzyj_row.Item("屈服强度MPa") = "0"
                                            Else
                                                Throw New InvalidDBException("管柱没有建立" & RECreader.Item("元件名称").ToString() & "数据")
                                            End If
                                        End Using
                                    End Using
                                Case "开关工具"
                                    gzyj_row.Item("元件重量kg") = "0"
                                    gzyj_row.Item("屈服强度MPa") = "0"
                                Case "伸缩管"
                                    SQL_command = "select 重量kg from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("重量kg")) AndAlso Val(RECreader2.Item("重量kg")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = RECreader2.Item("重量kg").ToString
                                                Else
                                                    Throw New InvalidDBException(RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                gzyj_row.Item("屈服强度MPa") = "0"
                                            Else
                                                Throw New InvalidDBException("管柱没有建立" & RECreader.Item("元件名称").ToString() & "数据")
                                            End If
                                        End Using
                                    End Using
                                Case "锚定工具"
                                    SQL_command = "select 重量kg from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("重量kg")) AndAlso Val(RECreader2.Item("重量kg")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = RECreader2.Item("重量kg").ToString
                                                Else
                                                    Throw New InvalidDBException(RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                gzyj_row.Item("屈服强度MPa") = "0"
                                            Else
                                                Throw New InvalidDBException("管柱没有建立" & RECreader.Item("元件名称").ToString() & "数据")
                                            End If
                                        End Using
                                    End Using
                                Case "普通钻杆"
                                    SQL_command = "select 单位长度质量kgpm,屈服强度MPa from 管柱_普通钻杆 where 井号='" & well_name & "' and 作业名称='" & zuoye_name _
                                                  & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("单位长度质量kgpm")) AndAlso Val(RECreader2.Item("单位长度质量kgpm")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = (CSng(RECreader.Item("元件长度m")) * CSng(RECreader2.Item("单位长度质量kgpm"))).ToString
                                                Else
                                                    Throw New InvalidDBException("钻杆：" & RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                If Not IsDBNull(RECreader2.Item("屈服强度MPa")) AndAlso Val(RECreader2.Item("屈服强度MPa")) <> 0 Then
                                                    gzyj_row.Item("屈服强度MPa") = RECreader2.Item("屈服强度MPa").ToString()
                                                Else
                                                    Throw New InvalidDBException("钻杆" & RECreader.Item("元件名称").ToString() & "没有屈服强度参数")
                                                End If
                                            Else
                                                Throw New InvalidDBException("管柱数据没有建立" & RECreader.Item("元件名称").ToString() & "钻杆")
                                            End If
                                        End Using
                                    End Using
                                Case "筛管"
                                    SQL_command = "select 单根质量kg,主材屈服强度MPa from 管柱_筛管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & RECreader.Item("元件序号").ToString()
                                    Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader2 As OleDbDataReader = EXECOleDbCommand2.ExecuteReader()
                                            If RECreader2.Read Then
                                                If Not IsDBNull(RECreader2.Item("单根质量kg")) AndAlso Val(RECreader2.Item("单根质量kg")) <> 0 Then
                                                    gzyj_row.Item("元件重量kg") = CSng(RECreader2.Item("单根质量kg")).ToString
                                                Else
                                                    Throw New InvalidDBException("筛管：" & RECreader.Item("元件名称").ToString() & "没有重量参数")
                                                End If
                                                If Not IsDBNull(RECreader2.Item("主材屈服强度MPa")) AndAlso Val(RECreader2.Item("主材屈服强度MPa")) <> 0 Then
                                                    gzyj_row.Item("屈服强度MPa") = RECreader2.Item("主材屈服强度MPa").ToString()
                                                Else
                                                    Throw New InvalidDBException("筛管" & RECreader.Item("元件名称").ToString() & "没有屈服强度参数")
                                                End If
                                            Else
                                                Throw New InvalidDBException("管柱数据没有建立" & RECreader.Item("元件名称").ToString() & "筛管")
                                            End If
                                        End Using
                                    End Using
                                Case Else
                                    gzyj_row.Item("元件重量kg") = "0"
                                    gzyj_row.Item("屈服强度MPa") = "0"
                            End Select
                            gz_yj_table.Rows.Add(gzyj_row)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As InvalidDBException
            msg_prompt = ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Catch ex As Exception
            msg_prompt = "计算管柱元件深度时出错！" & ex.Message
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        End Try
    End Sub
End Module

'20251124秦彦斌将Public Class InvalidDBException从pubFRm_call.vb移至main.vb
' 定义自定义异常类
Public Class InvalidDBException
    Inherits Exception
    Public Sub New()
        MyBase.New("数据库访问参数异常")
    End Sub
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub
End Class
