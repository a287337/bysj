Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_cal_zhijing
    '**********************************************************************************************************************************
    '                              节点套管内径、管柱内径、管柱外径,线重、油管钢级赋值子程序
    '                                                                                         2022年3月9日秦彦斌最后修订整理
    '输入参数：
    '  table_name              节点情况表名称                      类型：string
    '全局变量使用：
    '   well_name:             当前井号
    '   zuoye_name:            当前管柱作业名称
    '
    '编程思路：
    '   详见程序中的注释
    '
    '使用注意事项：
    '
    '运行结果：在给定库表名称如“节点情况表”中填入合适的数据，
    '          分别有：套管内径、管柱内径、管柱外径、线重、弹性模量、屈服强度、油管钢级、泊松比、热涨系数、抗挤强度MPa,管体抗内压MPa,接头抗内压MPa,管体抗拉kN,接头抗拉kN   
    '
    '修订情况：
    '    20200124除夕,程序迁移到VS2008完成，数据读取用ADO.NET提供的方法
    '    20211130，本模块针对节点生成的变化(详见jslb.vb模块)算法可用，原算法对计算点的节点类型赋值有误，对计算点的弹性模量、屈服强度、
    '              油管钢级、泊松比、热涨系数的赋值算法有误。
    '    20220306，实现往节点情况表中的抗挤强度MPa、管体抗内压MPa、接头抗内压MPa、管体抗拉kN、接头抗拉kN字段填写数据。
    '**********************************************************************************************************************************
    Sub cal_zhijing(ByVal table_name As String) '套管内径、管住内径、管住外径计算
        Dim SQL_command As String
        Dim jdqk_row As DataRow
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim ad As New System.Data.OleDb.OleDbDataAdapter
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim jd_count As Short
        Dim i As Short
        Dim gzh_waijing As Double
        Dim gzh_neijing As Double
        Dim gzh_xianzhong As Double
        Dim yj_changdu As Double
        Dim gzh_xgm_b As Double
        Dim gzh_yg_gj As String
        Dim gzh_e As Double
        Dim gzh_psb As Double
        Dim gzh_rzhxsh As Double
        Dim xgm_b As Double
        Dim yg_gj As String
        Dim kjqdu As Double '抗挤强度MPa
        Dim gtkny As Double '管体抗内压MPa
        Dim jtkny As Double '接头抗内压MPa
        Dim gtkla As Double '管体抗拉kN
        Dim jtkla As Double '接头抗拉kN 
        Dim gtd As Double   '曲率rad╱m
        Dim dlt_fwj As Double '单元上下方位角之差
        Dim dlt_jxj As Double '单元上下井斜角之差
        Dim k_alf As Double '井斜变化率
        Dim k_cta As Double '方位变化率
        Dim Ls As Double '单元长度
        Dim pj_alf As Double '单元上下端井斜角的平均值
        Dim jd_jdlx_arr() As String
        Dim jd_jdid_arr() As String
        Dim jd_jdxz_arr() As String
        Dim jd_jdbh_arr() As Short
        Dim jd_jdxs_arr() As Double
        Dim jd_jdcs_arr() As Double
        Dim jd_tgnj_arr() As Double
        Dim jd_gzwj_arr() As Double
        Dim jd_gznj_arr() As Double
        Dim jd_e_arr() As Double
        Dim jd_psb_arr() As Double
        Dim jd_xgm_b_arr() As Double '屈服强度
        Dim jd_yg_gj_arr() As String '油管钢级
        Dim jd_rzhxsh_arr() As Double '热膨胀系数(∕℃)
        Dim jd_kjqdu() As Double '抗挤强度MPa
        Dim jd_gtkny() As Double '管体抗内压MPa
        Dim jd_jtkny() As Double '接头抗内压MPa
        Dim jd_gtkla() As Double '管体抗拉kN
        Dim jd_jtkla() As Double '接头抗拉kN
        Dim jd_gtd() As Double '曲率rad╱m
        Dim jd_jxj() As Double '井斜角
        Dim jd_fwj() As Double '方位角
        Dim pi As Double 'I,圆周率

        pi = 3.14159265358979
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '***************************************************************************************************************************
        '读出该作业的节点参数
        '***************************************************************************************************************************
        SQL_command = "select * from  " & table_name & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 节点编号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
        jdqk_Table.Clear()
        ad.Fill(jdqk_Table)
        ad.Dispose()
        '节点数
        jd_count = jdqk_Table.Rows.Count
        ReDim jd_jdlx_arr(jd_count - 1)
        ReDim jd_jdid_arr(jd_count - 1)
        ReDim jd_jdxz_arr(jd_count - 1)
        ReDim jd_jdbh_arr(jd_count - 1)
        ReDim jd_jdxs_arr(jd_count - 1)
        ReDim jd_jdcs_arr(jd_count - 1)
        ReDim jd_tgnj_arr(jd_count - 1)
        ReDim jd_gzwj_arr(jd_count - 1)
        ReDim jd_gznj_arr(jd_count - 1)
        ReDim jd_e_arr(jd_count - 1)
        ReDim jd_psb_arr(jd_count - 1)
        ReDim jd_xgm_b_arr(jd_count - 1)
        ReDim jd_yg_gj_arr(jd_count - 1)
        ReDim jd_rzhxsh_arr(jd_count - 1)
        ReDim jd_kjqdu(jd_count - 1)
        ReDim jd_gtkny(jd_count - 1)
        ReDim jd_jtkny(jd_count - 1)
        ReDim jd_gtkla(jd_count - 1)
        ReDim jd_jtkla(jd_count - 1)
        ReDim jd_jxj(jd_count - 1)
        ReDim jd_fwj(jd_count - 1)
        ReDim jd_gtd(jd_count - 1)

        '***************************************************************************************************************************
        '正向循环，查对应下深处的套管内径
        '***************************************************************************************************************************
        gzh_waijing = 0.0#
        gzh_neijing = 0.0#
        yj_changdu = 0.0#
        gzh_xgm_b = 0.0#
        gzh_yg_gj = "\"
        gzh_e = 206842.72#
        gzh_psb = 0.3
        gzh_rzhxsh = 0.0000124
        kjqdu = 0.0#
        gtkny = 0.0#
        jtkny = 0.0#
        gtkla = 0.0#
        jtkla = 0.0#
        gtd = 0.0#
        i = 0
        For Each jdqk_row In jdqk_Table.Select
            jd_jdlx_arr(i) = jdqk_row.Item("节点类型").ToString
            jd_jdid_arr(i) = jdqk_row.Item("节点ID").ToString
            jd_jdxz_arr(i) = jdqk_row.Item("节点性质").ToString
            jd_jdbh_arr(i) = Val(jdqk_row.Item("节点编号").ToString)
            jd_jdxs_arr(i) = Val(jdqk_row.Item("节点下深m").ToString)
            jd_jdcs_arr(i) = Val(jdqk_row.Item("节点垂深m").ToString)
            jd_jxj(i) = Val(jdqk_row.Item("井斜角rad").ToString)
            jd_fwj(i) = Val(jdqk_row.Item("方位角rad").ToString)
            jd_tgnj_arr(i) = 0.0#
            jd_gzwj_arr(i) = 0.0#
            jd_gznj_arr(i) = 0.0#
            jd_e_arr(i) = gzh_e
            jd_psb_arr(i) = gzh_psb
            jd_xgm_b_arr(i) = gzh_xgm_b
            jd_yg_gj_arr(i) = gzh_yg_gj
            jd_rzhxsh_arr(i) = gzh_rzhxsh
            jd_kjqdu(i) = kjqdu
            jd_gtkny(i) = gtkny
            jd_jtkny(i) = jtkny
            jd_gtkla(i) = gtkla
            jd_jtkla(i) = jtkla
            jd_gtd(i) = gtd
            SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "' and 悬挂深度m<=" & CStr(jd_jdxs_arr(i)) & " and  套管下深m>=" & CStr(jd_jdxs_arr(i)) & " order by 套管外径mm"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            EXECOleDbCommand.Dispose()
            If RECreader.Read Then
                jd_tgnj_arr(i) = Val(RECreader.Item("套管外径mm").ToString) - 2.0# * Val(RECreader.Item("套管壁厚mm").ToString)
            Else
                jd_tgnj_arr(i) = 0
            End If
            i = i + 1
            RECreader.Close()
        Next
        ' 正向循环
        For i = 0 To jd_count - 1 Step 1
            '***************************************************************************************************************************
            ' 正向循环，若套管内径为0，则将上一节点的内径复制过来（裸眼井段，用上一段的套管内径作为井眼内径）
            ' 实际上裸眼直径等于此开钻进的钻头尺寸，应小于上一段的套管内径，但从井身结构数据表找不到此钻头尺寸，如此处理是没办法的办法。 
            ' ？？？？？？？？？？？        需要判断油管长于套管，且短于完钻井深        ？？？？？？？？？？？
            '***************************************************************************************************************************
            If jd_tgnj_arr(i) = 0 And i > 0 Then
                jd_tgnj_arr(i) = jd_tgnj_arr(i - 1)
            End If
            '***************************************************************************************************************************
            '                                                   计算管柱单元的全角变化
            '                                                                                       秦彦斌 2024年3月15日最后整理注释
            '    管柱单元的全角变化既节点处狗腿度,也即管柱单元的曲率，单位为弧度/m。
            '计算依据：
            '    管柱单元曲率的计算有2种方法，见稽国华《斜直井、定向井管柱屈曲分析与应用（最终）》式4-25,4-26，式4-26的计算方法称为最小
            '曲率法，式4-25的计算方法称为空间曲线法。“用最小曲率法计算的井眼曲率较小，用空间曲线法计算井眼曲率较为安全”。详见《钻井与
            '完井工程》陈平等主编 石油工业出版社 ISBN7502151990/TE。4014（课）250-252页，式6-6和式6-8。另见《油气井管柱力学与工程》 高
            '德利 中国石油大学出版社 ISBN：9787563621583，第85页式（3-1-21）-最小曲率法
            '***************************************************************************************************************************
            If i = 0 Then
                jd_gtd(i) = 0
            Else
                Ls = jd_jdxs_arr(i) - jd_jdxs_arr(i - 1)
                If Ls = 0 Then
                    jd_gtd(i) = 0
                Else
                    dlt_jxj = jd_jxj(i) - jd_jxj(i - 1)
                    dlt_fwj = jd_fwj(i) - jd_fwj(i - 1)
                    If System.Math.Abs(dlt_fwj) > pi Then
                        dlt_fwj = 2 * pi - System.Math.Abs(dlt_fwj)
                    End If
                    k_alf = dlt_jxj / Ls
                    k_cta = dlt_fwj / Ls
                    '管柱单元上下端井斜角的平均值
                    pj_alf = 0.5 * (jd_jxj(i) + jd_jxj(i - 1))
                    '空间曲线法，4-25式，单位为弧度/m，《钻井与完井工程》(6-8)式
                    'jd_gtd(i) = System.Math.Sqrt(k_alf * k_alf + k_cta * k_cta * System.Math.Sin(pj_alf) * System.Math.Sin(pj_alf)) / Ls
                    '最小曲率法，4-26式，单位为弧度/m ，《钻井与完井工程》(6-6)式
                    jd_gtd(i) = Arccos(System.Math.Cos(jd_jxj(i)) * System.Math.Cos(jd_jxj(i - 1)) + System.Math.Sin(jd_jxj(i)) * System.Math.Sin(jd_jxj(i - 1)) * System.Math.Cos(dlt_fwj)) / Ls
                End If
            End If
        Next i
        '***************************************************************************************************************************
        '反向循环，查对应下深处的管柱元件内、外径,顺便写入节点计算参数表
        '问题：如果是上下联通的开关元件关闭，要不要将其内径记为0？不同工况开关工具状态不同，要设为0也不应该在这里设。--20211130QYB答
        '***************************************************************************************************************************
        For i = jd_count - 1 To 0 Step -1
            If InStr(jd_jdlx_arr(i), "管柱") <> 0 And Trim(jd_jdid_arr(i)) <> "井口" Then
                SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                EXECOleDbCommand.Dispose()
                RECreader.Read()
                gzh_waijing = Val(RECreader.Item("元件外径mm").ToString)
                gzh_neijing = Val(RECreader.Item("元件内径mm").ToString)
                yj_changdu = Val(RECreader.Item("元件长度m").ToString)
                RECreader.Close()
                If jd_jdxz_arr(i) = "油井管" Then
                    SQL_command = "select * from 管柱_油管表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("单位长重kg╱m").ToString)
                    gzh_xgm_b = Val(RECreader.Item("屈服强度MPa").ToString)
                    gzh_yg_gj = RECreader.Item("油管钢级").ToString
                    gzh_e = Val(RECreader.Item("弹性模量MPa").ToString)
                    gzh_psb = Val(RECreader.Item("泊松比").ToString)
                    gzh_rzhxsh = Val(RECreader.Item("热膨胀系数").ToString)
                    kjqdu = Val(RECreader.Item("抗挤强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("接头抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("接头抗拉强度kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "开关工具" Then
                    SQL_command = "select * from 管柱_开关元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("重量kg").ToString) / yj_changdu
                    gzh_yg_gj = "\"
                    gzh_xgm_b = 0.0#
                    kjqdu = Val(RECreader.Item("抗外挤强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "封隔器" Then
                    SQL_command = "select * from 管柱_封隔定位元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("重量kg").ToString) / yj_changdu
                    gzh_yg_gj = "\"
                    gzh_xgm_b = Val(RECreader.Item("主材屈服强度MPa").ToString)
                    kjqdu = Val(RECreader.Item("抗外压强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("极限载荷kN").ToString)
                    jtkla = Val(RECreader.Item("极限载荷kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "节流工具" Then
                    SQL_command = "select * from 管柱_节流元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("重量kg").ToString) / yj_changdu
                    gzh_yg_gj = "\"
                    gzh_xgm_b = 0.0#
                    kjqdu = Val(RECreader.Item("抗外挤强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "伸缩管" Then
                    SQL_command = "select * from 管柱_伸缩元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("重量kg").ToString) / yj_changdu
                    gzh_yg_gj = "\"
                    gzh_xgm_b = 0.0#
                    kjqdu = Val(RECreader.Item("抗外挤强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "锚定工具" Then
                    SQL_command = "select * from 管柱_锚定元件 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("重量kg").ToString) / yj_changdu
                    gzh_yg_gj = "\"
                    gzh_xgm_b = 0.0#
                    kjqdu = Val(RECreader.Item("抗外压强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("抗拉强度kN").ToString)
                    RECreader.Close()
                End If
                If jd_jdxz_arr(i) = "普通钻杆" Then
                    SQL_command = "select * from 管柱_普通钻杆 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 元件序号=" & CStr(jd_jdid_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    RECreader.Read()
                    gzh_xianzhong = Val(RECreader.Item("单位长度质量kgpm").ToString)
                    gzh_xgm_b = Val(RECreader.Item("屈服强度MPa").ToString)
                    gzh_yg_gj = RECreader.Item("钢级").ToString
                    gzh_e = Val(RECreader.Item("弹性模量MPa").ToString)
                    gzh_psb = Val(RECreader.Item("泊松比").ToString)
                    gzh_rzhxsh = Val(RECreader.Item("热膨胀系数").ToString)
                    kjqdu = Val(RECreader.Item("抗挤强度MPa").ToString)
                    gtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    jtkny = Val(RECreader.Item("抗内压强度MPa").ToString)
                    gtkla = Val(RECreader.Item("管体抗拉强度kN").ToString)
                    jtkla = Val(RECreader.Item("接头抗拉强度kN").ToString)
                    RECreader.Close()
                End If
            End If
            jd_gzwj_arr(i) = gzh_waijing
            jd_gznj_arr(i) = gzh_neijing
            jd_xgm_b_arr(i) = gzh_xgm_b
            jd_yg_gj_arr(i) = gzh_yg_gj
            jd_e_arr(i) = gzh_e
            jd_psb_arr(i) = gzh_psb
            jd_rzhxsh_arr(i) = gzh_rzhxsh
            jd_kjqdu(i) = kjqdu
            jd_gtkny(i) = gtkny
            jd_jtkny(i) = jtkny
            jd_gtkla(i) = gtkla
            jd_jtkla(i) = jtkla
            '***************************************************************************************************************************
            ' 当Ls=[i]-[i-1]=0时，狗腿度取上面[i-1]的值，因同一深度
            '***************************************************************************************************************************
            If i <> 0 Then
                Ls = jd_jdxs_arr(i) - jd_jdxs_arr(i - 1)
                If Ls = 0 Then
                    jd_gtd(i) = jd_gtd(i - 1)
                End If
            End If
            SQL_command = "update  " & table_name & "  set " _
                & "套管内径mm=" & CStr(jd_tgnj_arr(i)) & "," & "油管外径mm=" & CStr(jd_gzwj_arr(i)) & "," & "油管内径mm=" & CStr(jd_gznj_arr(i)) & "," _
                & "线重kg╱m=" & CStr(gzh_xianzhong) & "," & "弹性模量MPa=" & CStr(jd_e_arr(i)) & "," & "泊松比=" & CStr(jd_psb_arr(i)) & "," _
                & "屈服强度MPa=" & CStr(jd_xgm_b_arr(i)) & "," & "油管钢级='" & jd_yg_gj_arr(i) & "'," & "热膨胀系数=" & CStr(jd_rzhxsh_arr(i)) & "," _
                & "抗挤强度MPa=" & CStr(jd_kjqdu(i)) & "," & "管体抗内压MPa=" & CStr(jd_gtkny(i)) & "," & "接头抗内压MPa=" & CStr(jd_jtkny(i)) & "," _
                & "管体抗拉kN=" & CStr(jd_gtkla(i)) & "," & "接头抗拉kN=" & CStr(jd_jtkla(i)) & "," & "曲率rad╱m=" & CStr(jd_gtd(i)) _
                & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点编号=" & CStr(jd_jdbh_arr(i))
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            EXECOleDbCommand = Nothing
        Next i
        '***************************************************************************************************************************
        '处理钢级、屈服极限没赋值的节点。方法，找最近的节点取值
        '***************************************************************************************************************************
        ' 反向循环
        xgm_b = 0.0#
        yg_gj = "\"
        For i = jd_count - 1 To 0 Step -1
            If jd_xgm_b_arr(i) <> 0.0# Then
                xgm_b = jd_xgm_b_arr(i)
            Else
                If xgm_b <> 0.0# Then
                    jd_xgm_b_arr(i) = xgm_b
                    SQL_command = "update  " & table_name & "  set " _
                    & "屈服强度MPa=" & CStr(jd_xgm_b_arr(i)) & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点编号=" & CStr(jd_jdbh_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                End If
            End If
            If jd_yg_gj_arr(i) <> "\" Then
                yg_gj = jd_yg_gj_arr(i)
            Else
                If yg_gj <> "\" Then
                    jd_yg_gj_arr(i) = yg_gj
                    SQL_command = "update  " & table_name & "  set " _
                    & "油管钢级='" & CStr(jd_yg_gj_arr(i)) & "' " & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点编号=" & CStr(jd_jdbh_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                End If
            End If
        Next i
        ' 正向循环
        For i = 0 To jd_count - 1 Step 1
            If jd_xgm_b_arr(i) <> 0.0# Then
                xgm_b = jd_xgm_b_arr(i)
            Else
                If xgm_b <> 0.0# Then
                    jd_xgm_b_arr(i) = xgm_b
                    SQL_command = "update  " & table_name & "  set " _
                    & "屈服强度MPa=" & CStr(jd_xgm_b_arr(i)) & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点编号=" & CStr(jd_jdbh_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                End If
            End If
            If jd_yg_gj_arr(i) <> "\" Then
                yg_gj = jd_yg_gj_arr(i)
            Else
                If yg_gj <> "\" Then
                    jd_yg_gj_arr(i) = yg_gj
                    SQL_command = "update  " & table_name & "  set " _
                    & "油管钢级='" & CStr(jd_yg_gj_arr(i)) & "' " & " where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点编号=" & CStr(jd_jdbh_arr(i))
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                End If
            End If
        Next i
        cn_userdb.Close()
        cn_userdb.Dispose()
        cn_userdb = Nothing
    End Sub
End Module