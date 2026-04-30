'**********************************************************************************************************************************
' 程序升级记事：
'                                                                                           秦彦斌 2020年01月12日最后整理
'  （1）20190717-开始从VB6升级到VS2008，20200112全部升级完成
'  （2）放弃没一条SQL语句就对数据库连接-打开-执行-关闭操作模式，变为一个模块开始时连接-打开，一系列执行完成后关闭数据库。减少
' 打开、关闭数据库次数。
'**********************************************************************************************************************************
Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_data_str
    Private b_dst As New DataSet("base_dst")
    Private color_Table As DataTable = b_dst.Tables.Add("color_Table")
    Private cn_userdb As System.Data.OleDb.OleDbConnection
    Private cn_basedb As System.Data.OleDb.OleDbConnection
    Private ad As New System.Data.OleDb.OleDbDataAdapter
    Private EXECOleDbCommand As OleDbCommand
    Private RECreader As OleDbDataReader
    Private SQL_command As String
    '**********************************************************************************************************************************
    '                                     建立数据库表过程，用以自动建立软件运行所需要的数据库表
    '参数说明：
    '   DBopt_select     数据库选择代号，取值为1时，连接操作基本数据库，取值为2时，连接操作用户数据库
    '   opt_code         用于指定要建立的数据库表，具体如下：
    '***********************************************************************************************************************************
    '  用户数据库库表建立  DBopt_select=2
    '***********************************************************************************************************************************
    '   通用基本数据
    '       opt_code=1   建立     油气井表
    '       opt_code=2   建立     套管数据表
    '       opt_code=3   建立     管柱数据表
    '       opt_code=4   建立     井斜数据表
    '   管柱力学分析数据
    '       opt_code=5   建立     管柱_封隔定位元件
    '       opt_code=6   建立     管柱_开关元件
    '       opt_code=7   建立     工况参数表
    '       opt_code=8   建立     工况_封隔定位元件
    '       opt_code=9   建立     工况_开关元件
    '       opt_code=10  建立     管柱_油管表
    '       opt_code=11  建立     管柱_节流元件
    '       opt_code=12  建立     节点情况表                     管柱力学计算用表
    '       opt_code=13  建立     节点计算参数表                 管柱力学计算用表
    '       opt_code=14  建立     管柱_伸缩元件
    '       opt_code=15  建立     简便TSM计算结果
    '       opt_code=16  建立     试油层参数
    '       opt_code=17  建立     工具参数
    '       opt_code=18  建立     试油层参数评估结果
    '       opt_code=19  建立     摩阻摩矩计算参数表           2018年8月7日，修井作业管柱力学分析模块用做摩阻摩矩分析功能实现。放弃原摩阻摩矩分析界面及计算代码。此表不再使用
    '       opt_code=20  建立     计算参数表                    用于保存软件运行时的一些重要参数，以利回放数据时显示
    '       opt_code=21  建立     三次井眼样条函数参数表
    '       opt_code=22  建立     试压情况数据，2个表，主表：试压参数，子表：试压套管。关系：1对多。建立主表：试压参数
    '       opt_code=23  建立     试压情况数据，2个表，主表：试压参数，子表：试压套管。关系：1对多。建立子表：试压套管
    '       opt_code=24  建立     固井质量数据
    '       opt_code=25  建立     完井设计报告
    '       opt_code=26  建立     管柱_油管初始损伤表           管柱动力学分析需要的原始数据,2016年10月增加，与管柱_油管表为1对多关系
    '       opt_code=27  建立     管柱_锚定元件                 2017年2月14日增加，与管柱数据表为1对多关系
    '       opt_code=28  建立     工况_锚定元件
    '       opt_code=29  建立     约束点临时表                  20170812增加，用于管柱力学分析中存储某工况下的约束点
    '       opt_code=30  建立     修井节点计算参数表            20171205增加，修井管柱力学计算用表
    '       opt_code=31  建立     管柱组合图表                  20171205增加，修井管柱管柱组合图表
    '       opt_code=32  建立     管柱_普通钻杆                 20180328增加，与管柱数据表为1对多关系
    '       opt_code=33  建立     卡点分析计算参数表            20180410增加，卡点计算用表：卡点分析节点情况表，卡点分析计算参数表
    '       opt_code=34  建立     工况_伸缩管状态               20180831增加，记录各工况下，各个伸缩管的拉伸、压缩状态
    '       opt_code=35  建立     重量及拉伸应力安全系数        20211111增加，记录管柱重量及拉伸应力安全系数计算结果
    '       opt_code=36  建立     修井摩阻分析_开关工具状态     20220325增加，记录修井摩阻分析时，各开关工具的状态
    '       opt_code=37  建立     修井及摩阻分析工况参数        20220329增加，保存某井某作业管柱修井及摩阻分析时的工况参数
    '       opt_code=38  建立     下入性分析工况参数        20221020增加，保存某井某作业管柱下入性分析时的工况参数及主要计算结果
    '       opt_code=39  建立     下入性分析_开关工具状态       20221020增加，记录管柱下入性分析时，各开关工具的状态
    '       opt_code=40  建立     下入性节点情况表、下入性节点计算参数表           20221020增加，管柱下入性分析计算用表

    '   套管磨损分析数据
    '       opt_code=62  建立     钻具组合-钻铤数据表
    '       opt_code=63  建立     钻具组合-加重钻杆数据表
    '       opt_code=64  建立     钻具组合-普通钻杆数据表
    '       opt_code=65  建立     钻井日志表
    '       opt_code=67  建立     钻具组合数据表
    '       opt_code=68  建立     磨损分析结果数据表
    '       opt_code=69  建立     磨损节点临时表                 套管磨损计算时的临时表
    '       opt_code=70  建立     套管磨损分析参数表             套管磨损计算时的摩擦系数、磨损效率修正系数
    '       opt_code=71  建立     磨损套管作业参数分析数据表
    '       opt_code=72  建立     完井后起下钻统计表

    '   管柱冲蚀预测数据
    '       opt_code=91  建立     冲蚀预测_压裂参数表
    '       opt_code=92  建立     冲蚀预测分析结果表

    '   套管振动疲劳寿命分析数据
    '       opt_code=501 建立     套管经历流体流动参数表
    '
    '   井眼轨迹三维显示
    '       opt_code=20001  建立  井眼轨迹插值参数表  李润洲2017年2月14日增加

    '***************************************************************************************************************************
    '  射孔相关参数表建立，从20010开始编号            李润洲2025年6月15日-2025年7月25日增加
    '***************************************************************************************************************************
    '  目的层参数表
    '       opt_code=20010  建立     目的层参数表，              call database_creat(20010,2)  '建立数据表  保存作业目的层压力、温度、岩石、地质参数，人工井底等
    '  射孔工况参数表
    '       opt_code=20011  建立     射孔工况参数表，            call database_creat(20011,2)  '建立数据表  射孔工况参数
    '  管柱-射孔枪弹表
    '       opt_code=20012  建立     管柱_射孔枪弹表，            call database_creat(20012,2)  '建立数据表  管柱_射孔枪弹参数
    '  工况_射孔夹层表
    '       opt_code=20013  建立     工况_射孔夹层表，            call database_creat(20013,2)  '建立数据表  工况_射孔夹层参数
    '  射孔段爆轰计算参数
    '       opt_code=20014  建立     射孔段爆轰计算参数，         call database_creat(20013,2)  '建立数据表  射孔段爆轰计算参数
    '  射孔段封隔器计算参数
    '       opt_code=20015  建立     射孔段封隔器计算参数，       call database_creat(20013,2)  '建立数据表  射孔段封隔器计算参数
    '  射孔全井油管应力计算参数
    '       opt_code=20016  建立     射孔油管应力计算参数，       call database_creat(20013,2)  '建立数据表  射孔油管应力计算参数
    '  管柱-筛管表
    '       opt_code=20017  建立     管柱_筛管表，                call database_creat(20017,2)  '建立数据表  管柱_筛管参数
    '***************************************************************************************************************************


    '***************************************************************************************************************************
    '  基础数据库库表建立 DBopt_select=1 ，从10001开始编号
    '***************************************************************************************************************************
    '       opt_code=10001  建立     封隔器信封曲线参数表，      Call database_creat(10001, 1)   '建立数据表  封隔器信封曲线参数表
    '       opt_code=10002  建立     锚定工具表          ，      Call database_creat(10002, 1)   '建立数据表  锚定工具表
    '       opt_code=10003  建立     开关元件载荷性能图参数表    Call database_creat(10003, 1)   '建立数据表  开关元件载荷性能图参数表
    '       opt_code=10004  建立     锚定工具载荷性能图参数表    Call database_creat(10004, 1)   '建立数据表  锚定工具载荷性能图参数表
    '***************************************************************************************************************************


    '***************************************************************************************************************************
    '  基础数据库库表射孔相关参数表建立，从10010开始编号            李润洲2025年7月15日增加
    '***************************************************************************************************************************
    '       opt_code=10010  建立     枪弹数据表，                Call database_creat(10010, 1)   '建立数据表  枪弹数据表   
    '       opt_code=10011  建立     炸药数据表，                Call database_creat(10011, 1)   '建立数据表  炸药数据表
    '       opt_code=10012  建立     岩石数据表，                Call database_creat(10012, 1)   '建立数据表  岩石数据表  
    '       opt_code=10013  建立     筛管表，                    Call database_creat(10013, 1)   '建立数据表  筛管
    '       opt_code=10014  建立     射孔施工方式表，            Call database_creat(10014, 1)   '建立数据表  射孔施工方式
    '***************************************************************************************************************************


    '变量AdoConString和use_AdoConString为全局变量，用于保存基本数据库ADO连接字符串和用户数据库ADO连接字符串
    '***************************************************************************************************************************
    Sub creat_tables()
        '建立计算所需要的库表
        use_AdoConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & use_dbname & ";Persist Security Info=False"
        'VS2008中下面的数据库引擎连接字符串用不成。
        '信息来自：https://blog.csdn.net/zyjq52uys/article/details/88545205
        '信息题目：VB.NET学习笔记：Microsoft Access数据库引擎（Jet、ACE）和数据库连接字符串
        'use_AdoConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & use_dbname & ";Persist Security Info=False" 

        Call database_creat(1, 2)    '建立数据表  油气井表
        Call database_creat(2, 2)    '建立数据表  套管数据表
        Call database_creat(3, 2)    '建立数据表  管柱数据表
        Call database_creat(4, 2)    '建立数据表  井眼轨道表
        Call database_creat(5, 2)    '建立数据表  管柱_封隔定位元件
        Call database_creat(6, 2)    '建立数据表  管柱_开关元件
        Call database_creat(7, 2)    '建立数据表  工况参数表
        Call database_creat(8, 2)    '建立数据表  工况-封隔器表
        Call database_creat(9, 2)    '建立数据表  工况_开关工具表
        Call database_creat(10, 2)    '建立数据表  管柱-油管表
        Call database_creat(11, 2)    '建立数据表  管柱-节流元件
        Call database_creat(12, 2)    '建立数据表  节点情况表
        Call database_creat(13, 2)    '建立数据表  节点计算参数表
        Call database_creat(14, 2)    '建立数据表  管柱-伸缩元件
        Call database_creat(15, 2)    '简便TSM计算结果
        Call database_creat(16, 2)    '试油层参数
        Call database_creat(17, 2)    '工具参数
        Call database_creat(18, 2)    '试油层参数评估结果
        'Call database_creat(19, 2)    '摩阻摩矩计算参数表,2018年8月7日，修井作业管柱力学分析模块用做摩阻摩矩分析功能实现。放弃原摩阻摩矩分析界面及计算代码。此表不再使用
        Call database_creat(20, 2)    '计算参数表
        Call database_creat(21, 2)    '三次井眼样条函数参数表
        Call database_creat(22, 2)    '试压情况数据，2个表，主表：试压参数，子表：试压套管。关系：1对多。建立主表：试压参数
        Call database_creat(23, 2)    '试压情况数据，2个表，主表：试压参数，子表：试压套管。关系：1对多。建立子表：试压套管
        Call database_creat(24, 2)    '固井质量表
        Call database_creat(25, 2)    '完井设计报告
        Call database_creat(26, 2)    '管柱_油管初始损伤表
        Call database_creat(27, 2)    '管柱_锚定元件
        Call database_creat(28, 2)    '工况_锚定元件
        Call database_creat(29, 2)    '建立数据表  约束点临时表
        Call database_creat(30, 2)    '建立数据表  修井节点计算参数表
        Call database_creat(31, 2)    '建立数据表  管柱组合图表
        Call database_creat(32, 2)    '建立数据表  管柱_普通钻杆表               20180328增加，与管柱数据表为1对多关系
        Call database_creat(33, 2)    '建立数据表  管柱_普通钻杆表               20180410增加，卡点计算用表：卡点分析节点情况表，卡点分析计算参数表
        Call database_creat(34, 2)    '建立数据表  工况_伸缩管状态               20180831增加，记录各工况下，各个伸缩管的拉伸、压缩状态
        Call database_creat(35, 2)    '建立数据表  重量及拉伸应力安全系数        20211111增加，记录管柱重量及拉伸应力安全系数计算结果
        Call database_creat(36, 2)    '建立数据表  修井摩阻分析_开关工具状态     20230325增加，记录修井摩阻分析时，各开关工具的状态
        Call database_creat(37, 2)    '建立数据表  修井及摩阻分析工况参数        20230329增加，保存某井某作业管柱修井及摩阻分析时的工况参数
        Call database_creat(38, 2)    '建立数据表  下入性分析工况参数            20221020增加，保存某井某作业管柱下入性分析时的工况参数及主要计算结果
        Call database_creat(39, 2)    '建立数据表  下入性分析_开关工具状态       20221020增加，记录管柱下入性分析时，各开关工具的状态
        Call database_creat(40, 2)    '建立数据表  下入性节点情况表、下入性节点计算参数表  20221020增加，管柱下入性分析计算用表
        Call database_creat(41, 2)    '建立数据表  作业地层参数表                20240711增加，保存作业所对应地层特性参数,   ‘**************************2025年7月22日李润洲修改，增加目的层名称列

        Call database_creat(62, 2)    '建立数据表   钻铤数据表
        Call database_creat(63, 2)    '建立数据表   加重钻杆数据表
        Call database_creat(64, 2)    '建立数据表   普通钻杆数据表
        Call database_creat(65, 2)    '建立数据表   钻井日志表
        Call database_creat(67, 2)    '建立数据表   钻具组合表

        Call database_creat(68, 2)    '建立数据表   磨损分析结果数据表
        Call database_creat(70, 2)    '建立数据表   套管磨损分析参数表             套管磨损计算时的摩擦系数、磨损效率
        Call database_creat(71, 2)    '建立数据表   磨损套管作业参数分析数据表
        Call database_creat(72, 2)    '建立数据表   完井后起下钻统计表
        Call database_creat(91, 2)    '建立数据表   冲蚀预测-压裂参数表
        Call database_creat(92, 2)    '建立数据表   冲蚀预测分析结果表
        Call database_creat(501, 2)   '建立数据表   套管经历流体流动参数表

        Call database_creat(20001, 2) '建立数据表，井眼轨迹插值数据表        李润洲2017年2月14日增加
        '*********************************************************************************************************************************
        Call database_creat(20010, 2) '建立数据表，目的层参数表              李润洲2025年6月15日增加
        Call database_creat(20011, 2) '建立数据表，射孔工况参数表            李润洲2025年6月18日增加
        Call database_creat(20012, 2) '建立数据表，管柱-射孔枪弹表           李润洲2025年6月18日增加
        Call database_creat(20013, 2) '建立数据表，工况_射孔夹层表           李润洲2025年6月18日增加
        Call database_creat(20014, 2) '建立数据表，射孔段爆轰计算参数        李润洲2025年7月14日增加   存放爆轰计算结果
        Call database_creat(20015, 2) '建立数据表，射孔段封隔器计算参数      李润洲2025年7月14日增加   存放射孔段封隔器应力计算结果
        Call database_creat(20016, 2) '建立数据表，射孔油管应力计算参数      李润洲2025年7月14日增加   存放射孔全井油管应力计算结果 
        Call database_creat(20017, 2) '建立数据表，管柱-筛管表               李润洲2025年7月24日增加   
    End Sub
    '*********************************************************************************************************************************
    '                                    基础数据数据结构升级与新建   开始
    '*********************************************************************************************************************************
    Sub base_database_update()
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim columnTable As DataTable
        Dim columnTab_row As DataRow
        Dim rowfilter As String
        Dim table_name As String
        Dim col_name As String
        Dim i As Short
        Dim field_finded() As Boolean
        Dim sj_strchg As Boolean '数据结构有升级标志

        '*********************************************************************************************************************************
        '2016年9月3日 基础数据表新增“封隔器信封曲线参数表”，用以保存封隔器信封曲线参数
        '*********************************************************************************************************************************
        Call database_creat(10001, 1) '建立数据表  封隔器信封曲线参数表
        '*********************************************************************************************************************************
        '2017年2月12日 基础数据表新增“锚定工具表”，用以保存锚定工具参数
        '*********************************************************************************************************************************
        Call database_creat(10002, 1) '建立数据表  锚定工具表
        '*********************************************************************************************************************************
        '2017年10月15日 基础数据表新增“开关元件载荷性能图参数表”，用以保存开关元件载荷性能图参数
        '*********************************************************************************************************************************
        Call database_creat(10003, 1) '建立数据表  开关元件载荷性能图参数表
        '*********************************************************************************************************************************
        '2017年10月24日 基础数据表新增“锚定工具载荷性能图参数表”，用以保存锚定工具载荷性能图参数
        '*********************************************************************************************************************************
        Call database_creat(10004, 1) '建立数据表  锚定工具载荷性能图参数表
        '*********************************************************************************************************************************
        '2025年7月22日 基础数据表新增“枪弹数据表”，用以保存岩石弹性模量、抗压强度等参数   李润洲
        '*********************************************************************************************************************************
        Call database_creat(10010, 1)
        '*********************************************************************************************************************************
        '2025年7月20日 基础数据表新增“炸药表”，用以保存岩石弹性模量、抗压强度等参数   李润洲
        '*********************************************************************************************************************************
        Call database_creat(10011, 1)
        '*********************************************************************************************************************************
        '2025年7月20日 基础数据表新增“岩石数据表”，用以保存岩石弹性模量、抗压强度等参数   李润洲
        '*********************************************************************************************************************************
        Call database_creat(10012, 1)
        '*********************************************************************************************************************************
        '2025年7月24日 基础数据表新增“筛管”，用以保存筛管-工具参数   李润洲
        '*********************************************************************************************************************************
        Call database_creat(10013, 1)
        '*********************************************************************************************************************************
        '2025年8月2日 基础数据表新增“射孔施工方式表”，用以保存射孔施工方式   李润洲
        '*********************************************************************************************************************************
        Call database_creat(10014, 1)
        '*********************************************************************************************************************************
        '                                                       基础数据库中开关元件表数据结构升级
        '2017年10月15日数据结构升级情况：
        '   （1）增加以下8个字段：
        '       01       温度范围下℃     FLOAT
        '       02       温度范围上℃     FLOAT
        '       03       压力等级MPa      FLOAT
        '       04       型号             TEXT(100)
        '       05       开关类型         TEXT(100)
        '       06       上端扣型         TEXT(100)
        '       07       下端扣型         TEXT(100)
        '       08       生产厂家         TEXT(100)
        '   （2）放弃并删除使用2个字段：开关方式，开关流向。注：开关流向在工况中考虑
        '    (3) 关键字：[外径(mm)]，[内径(mm)]，[长度)(m)],[型号]，[开关类型]
        '*********************************************************************************************************************************
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        dbSchema = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = dbSchema.Select("TABLE_NAME='开关元件'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(10)
            sj_strchg = False
            For i = 1 To 10
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If
                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If
                If col_name = "压力等级MPa" Then
                    field_finded(3) = True
                End If
                If col_name = "型号" Then
                    field_finded(4) = True
                End If
                If col_name = "开关类型" Then
                    field_finded(5) = True
                End If
                If col_name = "上端扣型" Then
                    field_finded(6) = True
                End If
                If col_name = "下端扣型" Then
                    field_finded(7) = True
                End If
                If col_name = "生产厂家" Then
                    field_finded(8) = True
                End If
                If col_name = "开关方式" Then
                    field_finded(9) = True
                End If
                If col_name = "开关流向" Then
                    field_finded(10) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 压力等级MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 压力等级MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 型号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 型号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 开关类型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型='配水器' where 名称 like '%配水器%'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型='球座' where 名称 like '%球座%'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型='安全阀' where 名称 like '%安全阀%'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型='滑套' where 名称 like '%滑套%'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 开关类型='滑套' where 名称 like '%测试阀%'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 开关元件 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 开关元件 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(9) = True Then
                SQL_command = "ALTER TABLE 开关元件 DROP COLUMN [开关方式]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(10) = True Then
                SQL_command = "ALTER TABLE 开关元件 DROP COLUMN [开关流向]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            End If
            If sj_strchg = True Then
                MsgBox("基础数据库开关元件数据表已升级。")
            End If
        Else
            MsgBox("基础数据库中未找到【开关元件】数据表，程序无法运行。")
            End
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中伸缩元件表数据结构升级
        '2016年10月9日数据结构升级情况：
        '   （1）增加以下12个字段：
        '       01       温度范围下℃     FLOAT
        '       02       温度范围上℃     FLOAT
        '       03       压力等级MPa      FLOAT
        '       04       抗拉强度kN       FLOAT
        '       05       上端扣型         TEXT(100)
        '       06       下端扣型         TEXT(100)
        '       07       主体材料         TEXT(100)
        '       08       主材屈服强度MPa  FLOAT
        '       09       零件号           TEXT(100)
        '       10       生产厂家         TEXT(100)
        '       11       备注             TEXT(200)
        '       12       全缩短长度m      FLOAT
        '   （2）放弃使用2个字段：中心管外径(mm)，中心管内径(mm)
        '   （3）注意：计算出的 全拉开长度m=全缩短长度m+伸缩行程；全缩短长度m<=下入长度<=全拉开长度m
        '   （4）原“长度(m)”系指“全缩短长度m”，在管柱-伸缩元件用户表中“长度”指下入长度，其值大于等于全缩短长度m。
        '    (5) 增加字段 “全缩短长度m”，FLOAT型;将字段 “长度(m)” 的值赋给字段“全缩短长度m”;删除字段“长度(m)”
        '    (6) 关键字：[外径(mm)]，[内径(mm)]，[伸缩行程(m)]，[全缩短长度m]，[零件号]
        '2016年10月15日数据结构升级情况：
        '   （1）删除放弃使用的2个字段：中心管外径(mm)，中心管内径(mm)
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='伸缩元件'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(14)
            sj_strchg = False
            For i = 1 To 14
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If
                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If
                If col_name = "压力等级MPa" Then
                    field_finded(3) = True
                End If
                If col_name = "抗拉强度kN" Then
                    field_finded(4) = True
                End If
                If col_name = "上端扣型" Then
                    field_finded(5) = True
                End If
                If col_name = "下端扣型" Then
                    field_finded(6) = True
                End If
                If col_name = "主体材料" Then
                    field_finded(7) = True
                End If
                If col_name = "主材屈服强度MPa" Then
                    field_finded(8) = True
                End If
                If col_name = "零件号" Then
                    field_finded(9) = True
                End If
                If col_name = "生产厂家" Then
                    field_finded(10) = True
                End If
                If col_name = "备注" Then
                    field_finded(11) = True
                End If
                If col_name = "全缩短长度m" Then
                    field_finded(12) = True
                End If
                If col_name = "中心管外径(mm)" Then
                    field_finded(13) = True
                End If
                If col_name = "中心管内径(mm)" Then
                    field_finded(14) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 伸缩元件 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 伸缩元件 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 压力等级MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 伸缩元件 set 压力等级MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 抗拉强度kN=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 伸缩元件 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 伸缩元件 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 主体材料 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 主体材料=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 主材屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 主材屈服强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 零件号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 零件号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 备注 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 伸缩元件 ADD COLUMN 全缩短长度m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 伸缩元件 set [全缩短长度m]=[长度(m)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "ALTER TABLE 伸缩元件 DROP COLUMN [长度(m)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(13) = True Then
                SQL_command = "ALTER TABLE 伸缩元件 DROP COLUMN [中心管外径(mm)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(14) = True Then
                SQL_command = "ALTER TABLE 伸缩元件 DROP COLUMN [中心管内径(mm)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("基础数据库伸缩元件数据表已升级。")
            End If
        Else
            MsgBox("基础数据库中未找到【 伸缩元件】数据表，程序无法运行。")
            End
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中封隔器表数据结构升级
        '2016年8月28日-9月2日数据结构升级情况：
        '   （1）增加以下13个字段：
        '       01       温度范围下℃     FLOAT
        '       02       温度范围上℃     FLOAT
        '       03       抗内压强度MPa    FLOAT
        '       04       抗外压强度MPa    FLOAT
        '       05       型号             TEXT(100)
        '       06       生产厂家         TEXT(100)
        '       07       上端扣型         TEXT(100)
        '       08       下端扣型         TEXT(100)
        '       09       主体材料         TEXT(100)
        '       10       主材屈服强度MPa  FLOAT
        '       11       最小坐封压力MPa  FLOAT
        '       12       最大坐封压力MPa  FLOAT
        '       13       备注             TEXT(200)
        '
        '2016年9月2日
        '   还应该删除下列字段
        '    'If Not IsNull(Adodc1.Recordset.Fields("定位方式")) Then Combo2.text = Adodc1.Recordset.Fields("定位方式")
        '    'Text9.text = Adodc1.Recordset.Fields("剪销承载(kN)")
        '    'Text11.text = Adodc1.Recordset.Fields("中心管外径(mm)")
        '    'Text12.text = Adodc1.Recordset.Fields("中心管内径(mm)")
        '    '封隔方式
        '2016年9月16日
        '   （1）不用的几个字段暂时不删除，以免老程序运行出现错误。
        '   （2）关键字为：[型号][最大外径(mm)][最小通径(mm)][长度(m)][坐封方式]
        '2016年9月18日
        '    (1) 增加字段 “重量kg”，FLOAT型
        '    (2) 将字段 “重量(kN)” 的值赋给字段“重量kg”
        '    (3) 删除字段“重量(kN)”
        '2017年2月4日
        '    (1) 增加字段 “能否反洗井”，TEXT(10)
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='封隔器'")
        If foundRows.Length <> 0 Then '有封隔器
            ReDim field_finded(15)
            sj_strchg = False
            For i = 1 To 15
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If
                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If
                If col_name = "抗内压强度MPa" Then
                    field_finded(3) = True
                End If
                If col_name = "抗外压强度MPa" Then
                    field_finded(4) = True
                End If
                If col_name = "型号" Then
                    field_finded(5) = True
                End If
                If col_name = "生产厂家" Then
                    field_finded(6) = True
                End If
                If col_name = "上端扣型" Then
                    field_finded(7) = True
                End If
                If col_name = "下端扣型" Then
                    field_finded(8) = True
                End If
                If col_name = "主体材料" Then
                    field_finded(9) = True
                End If
                If col_name = "主材屈服强度MPa" Then
                    field_finded(10) = True
                End If
                If col_name = "最小坐封压力MPa" Then
                    field_finded(11) = True
                End If
                If col_name = "最大坐封压力MPa" Then
                    field_finded(12) = True
                End If
                If col_name = "备注" Then
                    field_finded(13) = True
                End If
                If col_name = "重量kg" Then
                    field_finded(14) = True
                End If
                If col_name = "能否反洗井" Then
                    field_finded(15) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 抗内压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 抗外压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 抗外压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 型号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 型号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 主体材料 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 主体材料=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 主材屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 主材屈服强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 最小坐封压力MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 最小坐封压力MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 最大坐封压力MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 最大坐封压力MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(13) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 备注 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(14) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 重量kg  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set [重量kg]=[重量(kN)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "ALTER TABLE 封隔器 DROP COLUMN [重量(kN)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(15) = False Then
                SQL_command = "ALTER TABLE 封隔器 ADD COLUMN 能否反洗井 TEXT(10)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 封隔器 set 能否反洗井='不能'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("基础数据库封隔器数据表已升级。")
            End If
        Else
            MsgBox("基础数据库中未找到【封隔器】数据表，程序无法运行。")
            End
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中 管柱元件性质 表数据升级
        '2017年2月3日
        '     考虑用户习惯，将管柱元件分类重新定义，管柱元件由原来的：油管、开关元件、封隔定位元件、节流元件、伸缩元件5种，调整修改为:
        ' 油井管、开关工具、封隔器、节流工具、伸缩管、锚定工具6种。为此，原来基础数据库管柱元件性质表中管柱元件性质字段值需做相应的修改：
        '将“油管”改为“油井管”；
        '将“开关元件”改为“开关工具”；
        '将“封隔定位元件”改为“封隔器”；
        '将“节流元件”改为“节流工具”；
        '将“伸缩元件”改为“伸缩管”。
        '增加一条记录，其管柱元件性质字段值为“锚定工具”
        '
        '2017年11月14日
        '   考虑不同的用户，所用的管柱元件不同，增加“Enabled”字段，整型，为1选取，为0不选取
        '   考虑修井管柱使用钻具，增加“井下钻具”、“射孔枪”、“普通钻杆”、“加重钻杆”、“钻铤”记录
        '
        '2025年7月24日         李润洲
        '  用于射孔作业，修改“射孔枪”的Enabled=1；增加“筛管”记录       
        '*********************************************************************************************************************************
        SQL_command = "update 管柱元件性质 set 管柱元件性质='油井管' where 管柱元件性质='油管'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        SQL_command = "update 管柱元件性质 set 管柱元件性质='开关工具' where 管柱元件性质='开关元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        SQL_command = "update 管柱元件性质 set 管柱元件性质='封隔器' where 管柱元件性质='封隔定位元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        SQL_command = "update 管柱元件性质 set 管柱元件性质='节流工具' where 管柱元件性质='节流元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        SQL_command = "update 管柱元件性质 set 管柱元件性质='伸缩管' where 管柱元件性质='伸缩元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='锚定工具'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质) values ('锚定工具')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        foundRows = dbSchema.Select("TABLE_NAME='管柱元件性质'")
        If foundRows.Length <> 0 Then '有管柱元件性质
            ReDim field_finded(1)
            sj_strchg = False
            field_finded(0) = False
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "Enabled" Then
                    field_finded(0) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(0) = False Then
                SQL_command = "ALTER TABLE 管柱元件性质 ADD COLUMN Enabled INTEGER"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 管柱元件性质 set Enabled=1"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
        Else
            MsgBox("基础数据库中未找到【管柱元件性质】数据表，程序无法运行。")
            End
        End If

        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='井下钻具'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('井下钻具',0)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='普通钻杆'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('普通钻杆',0)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='加重钻杆'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('加重钻杆',0)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='钻铤'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('钻铤',0)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='射孔枪'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            '李润洲2025年7月24日修改enabled=1
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('射孔枪',1)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        Else
            SQL_command = "update 管柱元件性质 set Enabled=1 where 管柱元件性质='射孔枪'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        '2025年7月24日增加“筛管”
        SQL_command = "select * from 管柱元件性质 where 管柱元件性质='筛管'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows = False Then
            SQL_command = "insert into 管柱元件性质(管柱元件性质,Enabled) values ('筛管',1)"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中 封隔器定位方式 表数据升级
        '2017年2月18日：
        '     考虑用户习惯，考虑封隔器和锚定工具均有锚定定位功能，为此，原来字段值“双向移动”改为“未(无)锚定”，以便封隔器、锚定工具均能
        ' 从此表中选择定位情况。
        '*********************************************************************************************************************************
        SQL_command = "update 封隔器定位方式 set 定位方式='未(无)锚定' where 定位方式='双向移动'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        '*********************************************************************************************************************************
        '                                                       基础数据库中 套管 表数据升级
        '2017年9月3日：
        '     考虑套管力学分析需要增加以下1个字段：
        '       01       线重(kg/m)  Float
        '2018年7月18日：
        '     考虑套管工作能力分析时应取管体及接头强度较低者参与计算的可能，以及保证套管数据的完整，增加以下字段：
        '       01       接箍外径mm               Float
        '       02       接箍长度mm               Float
        '2018年7月20日：
        '     2015年2015年8月程序自动升级后所发布的基础数据库里的套管表中，已经存在“单位长质量(kg/m)”、“接头抗内压强度(MPa)”、
        '“接头抗拉强度(kN)” 字段，增加“线重(kg/m)”字段为重复之举。
        '    增加“接箍外径mm”、“接箍长度mm”字段于套管强度计算目前看来，无关。所以，先放着。
        '*********************************************************************************************************************************
        '    ReDim field_finded(2)
        '    sj_strchg = False
        '    For i = 1 To 2
        '        field_finded(i) = False
        '    Next i
        '    adoCN.Open AdoConString
        '    Set rstSchema = adoCN.OpenSchema(adSchemaColumns)
        '    rstSchema.MoveFirst
        '    Do
        '        table_name = rstSchema.Fields("TABLE_NAME")
        '        col_name = rstSchema.Fields("COLUMN_NAME")
        '        If table_name = "套管" And col_name = "接箍外径mm" Then
        '            field_finded(1) = True
        '        End If
        '        If table_name = "套管" And col_name = "接箍长度mm" Then
        '            field_finded(2) = True
        '        End If
        '        rstSchema.MoveNext
        '    Loop Until rstSchema.EOF
        '    If field_finded(1) = False Then
        '        SQL_command = "ALTER TABLE 套管 ADD COLUMN [接箍外径mm]  FLOAT"
        '        Set SQL_rst = ExecuteSQL(SQL_command, msg_prompt, 1)
        '        SQL_command = "update 套管 set [接箍外径mm]=0.0"
        '        Set SQL_rst = ExecuteSQL(SQL_command, msg_prompt, 1)
        '        sj_strchg = True
        '    End If
        '    If field_finded(2) = False Then
        '        SQL_command = "ALTER TABLE 套管 ADD COLUMN [接箍长度mm]  FLOAT"
        '        Set SQL_rst = ExecuteSQL(SQL_command, msg_prompt, 1)
        '        SQL_command = "update 套管 set [接箍长度mm]=0.0"
        '        Set SQL_rst = ExecuteSQL(SQL_command, msg_prompt, 1)
        '        sj_strchg = True
        '    End If
        '    adoCN.Close
        '*********************************************************************************************************************************
        '                                                       基础数据库中 油管 表数据升级
        ' 2018年7月30日：
        '     油管数据库字段规范：加 扣型、接头抗内压强度、接头抗拉强度、备注 字段，为判断管柱管体、接头工作能力做准备。
        ' 油管关键字应该为：外径、壁厚、钢级、扣型。
        ' 2018年7月30日前字段内容：
        '     油管规格 TEXT(50)     ,材料  TEXT(50)    ,油管外径mm    FLOAT,油管壁厚mm FLOAT   ,加厚情况   TEXT(20),
        '     单位长度质量kg/m FLOAT,屈服应力MPa  FLOAT,抗内压强度MPa FLOAT,抗外挤强度MPa FLOAT,抗拉强度kN FLOAT   ,
        '     弹性模量MPa FLOAT     ,泊松比  FLOAT
        '
        ' 关键字：材料,油管外径mm,油管壁厚mm,单位长度质量kg/m
        '
        ' 2018年7月30日增加字段
        '   (01)扣型              TEXT(50)
        '   (02)接头抗内压强度MPa FLOAT
        '   (03)接头抗拉强度kN    FLOAT
        '   (04)备注              TEXT(250)
        ' 2018年7月30日删除字段
        '   (05)加厚情况   TEXT(20)
        ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
        '   (06)热膨胀系数    FLOAT 
        ' 关键字：油管外径mm,油管壁厚mm,材料（钢级）,扣型，单位长度质量kg/m
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='油管'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(6)
            sj_strchg = False
            For i = 1 To 5
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "扣型" Then
                    field_finded(1) = True
                End If
                If col_name = "接头抗内压强度MPa" Then
                    field_finded(2) = True
                End If
                If col_name = "接头抗拉强度kN" Then
                    field_finded(3) = True
                End If
                If col_name = "备注" Then
                    field_finded(4) = True
                End If
                If col_name = "加厚情况" Then
                    field_finded(5) = True
                End If
                If col_name = "热膨胀系数" Then
                    field_finded(6) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 油管 ADD COLUMN 扣型 TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 扣型=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 扣型='EU' where 加厚情况='加厚'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 扣型='NU' where 加厚情况='不加厚'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 油管 ADD COLUMN 接头抗内压强度MPa  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 接头抗内压强度MPa=抗内压强度MPa"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 油管 ADD COLUMN 接头抗拉强度kN  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 接头抗拉强度kN=抗拉强度kN"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 油管 ADD COLUMN 备注 TEXT(250)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 备注=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(5) = True Then
                SQL_command = "ALTER TABLE 油管 DROP COLUMN 加厚情况"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            '*********************************************************************************************************************************************
            ' 其中材料的线膨胀系数取值依据见：化学工业出版社，成大先主编《机械设计手册》第四版 第一卷 1-11页表1-1-14。
            ' 在本程序中采用定值，取各温度段的平均值的平均值。由于手册中取值范围较大，故：
            '    按窦老师2015年2月10日提供的“00 高难度复杂井完井（试油）油套管柱力学分析要点及若干工程问题简析（2011年5月11日西安完井会议）.ppt”中“温度
            ' 效应”公式及热膨胀系数（1.2*10^(-5)）计算公式取默认值。
            '    对于钢，WELLCAT中为12.4*10E-6/℃。
            '*********************************************************************************************************************************************
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 油管 ADD COLUMN 热膨胀系数  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 油管 set 热膨胀系数=1.24E-5"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
        Else
            MsgBox("基础数据库中未找到【油管】数据表，程序无法运行。")
            End
        End If
        '*********************************************************************************************************************************
        '                                            基础数据库中 钻杆（即普通钻杆） 表数据升级
        '2018年3月18日：
        '     考虑修井等作业时，井筒内可能下入钻杆，为进行管柱力学，钻杆基础数据中应增加必要的字段。
        '2018年3月18日前字段内容：
        '     钻杆规格 TEXT(50),钻杆外径(mm)     FLOAT,钻杆壁厚(mm) FLOAT,单根长度(m) FLOAT,单位长度质量(kg/m) FLOAT,
        '     加厚型式 TEXT(50),最小抗拉强度(kN) FLOAT,备注      TEXT(50)
        '  关键字：钻杆规格，钻杆外径(mm)，钻杆壁厚(mm)，单根长度(m)
        '2018年3月18日增加字段
        '   (01)钢级              TEXT(50)
        '   (02)扣型              TEXT(50)
        '   (03)屈服强度MPa       FLOAT
        '   (04)弹性模量MPa       FLOAT                   钢制和铝合金的不同
        '   (05)泊松比            FLOAT                   钢制和铝合金的不同
        '   (06)管体抗扭强度Nm    FLOAT
        '   (07)接头抗扭强度Nm    FLOAT
        '   (08)管体抗拉强度kN    FLOAT
        '   (09)接头抗拉强度kN    FLOAT
        '   (10)抗内压强度MPa     FLOAT
        '   (11)抗挤强度MPa       FLOAT
        '   (12)接头外径mm        FLOAT
        '   (13)将“最小抗拉强度(kN)”值赋给“管体抗拉强度kN”后删除
        ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
        '   (14)热膨胀系数    FLOAT
        '  关键字：钻杆规格，钻杆外径(mm)，钻杆壁厚(mm)，单根长度(m)
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='钻杆'")
        If foundRows.Length <> 0 Then '有钻杆表
            ReDim field_finded(14)
            sj_strchg = False
            For i = 1 To 13
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "钢级" Then
                    field_finded(1) = True
                End If
                If col_name = "扣型" Then
                    field_finded(2) = True
                End If
                If col_name = "屈服强度MPa" Then
                    field_finded(3) = True
                End If
                If col_name = "弹性模量MPa" Then
                    field_finded(4) = True
                End If
                If col_name = "泊松比" Then
                    field_finded(5) = True
                End If
                If col_name = "管体抗扭强度Nm" Then
                    field_finded(6) = True
                End If
                If col_name = "接头抗扭强度Nm" Then
                    field_finded(7) = True
                End If
                If col_name = "管体抗拉强度kN" Then
                    field_finded(8) = True
                End If
                If col_name = "接头抗拉强度kN" Then
                    field_finded(9) = True
                End If
                If col_name = "抗内压强度MPa" Then
                    field_finded(10) = True
                End If
                If col_name = "抗挤强度MPa" Then
                    field_finded(11) = True
                End If
                If col_name = "接头外径mm" Then
                    field_finded(12) = True
                End If
                If col_name = "最小抗拉强度(kN)" Then
                    field_finded(13) = True
                End If
                If col_name = "热膨胀系数" Then
                    field_finded(14) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 钢级 TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 钢级=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 扣型  TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 扣型=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 屈服强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 弹性模量MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 弹性模量MPa=206000"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 泊松比 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 泊松比=0.3"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 管体抗扭强度Nm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 管体抗扭强度Nm=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 接头抗扭强度Nm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 接头抗扭强度Nm=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 管体抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 管体抗拉强度kN=[最小抗拉强度(kN)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 接头抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 接头抗拉强度kN=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 抗内压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 抗挤强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 抗挤强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 接头外径mm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 接头外径mm=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(13) = True Then
                SQL_command = "ALTER TABLE 钻杆 DROP COLUMN [最小抗拉强度(kN)]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            '*********************************************************************************************************************************************
            ' 其中材料的线膨胀系数取值依据见：化学工业出版社，成大先主编《机械设计手册》第四版 第一卷 1-11页表1-1-14。
            ' 在本程序中采用定值，取各温度段的平均值的平均值。由于手册中取值范围较大，故：
            '    按窦老师2015年2月10日提供的“00 高难度复杂井完井（试油）油套管柱力学分析要点及若干工程问题简析（2011年5月11日西安完井会议）.ppt”中“温度
            ' 效应”公式及热膨胀系数（1.2*10^(-5)）计算公式取默认值。
            '    对于钢，WELLCAT中为12.4*10E-6/℃。
            '*********************************************************************************************************************************************
            If field_finded(14) = False Then
                SQL_command = "ALTER TABLE 钻杆 ADD COLUMN 热膨胀系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 钻杆 set 热膨胀系数=1.24E-5"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
        Else
            MsgBox("基础数据库中未找到【钻杆】数据表，程序无法运行。")
            End
        End If
        '    If sj_strchg = True Then
        '        MsgBox "基础数据库钻杆表已升级。"
        '    End If
        '*********************************************************************************************************************************
        '                                            基础数据库中 增加并填充绘图颜色表
        '                                                                              秦彦斌 2020年1月3日最后更新
        '2020年1月3日：
        '     用iPlotX绘制曲线时，该控件默认最多自动给出7种颜色，而且有白色及黄色太浅，分辨度不够。为此：
        ' (1) 在基础数据库中增加("绘图颜色表")，自动加入经反复对比找出的26种分辨度高颜色。详见文档《RGB颜色对照表.docx》
        ' (2) 提供颜色查取函数，传入颜色序号，1开头，返回Uint32值
        ' (3) iPlotX使用颜色时，次序时B.G.R，故库里次序也是B.G.R
        '*********************************************************************************************************************************
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        dbSchema = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='绘图颜色表'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length = 0 Then
            SQL_command = "CREATE TABLE 绘图颜色表(序号 INTEGER,英文名称 TEXT(50),红色值 INTEGER, 绿色值 INTEGER,蓝色值 INTEGER,HEXBGR色名 TEXT(10))"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (1,'Red',255,0,0,'0000FF')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (2,'Blue',0,0,255,'FF0000')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (3,'Orange',255,165,0,'00A5FF')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (4,'Green',0,255,0,'00FF00')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (5,'Magenta',255,0,255,'FF00FF')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (6,'Cyan1',0,255,255,'FFFF00')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (7,'Black',0,0,0,'000000')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (8,'DarkRed',139,0,0,'00008B')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (9,'Blue4',0,0,139,'8B0000')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (10,'Chocolate3',205,102,29,'1D66CD')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (11,'DarkGreen',0,100,0,'006400')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (12,'Purple',160,32,240,'F020A0')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (13,'DarkCyan',0,139,139,'8B8B00')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (14,'grey41',105,105,105,'696969')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (15,'Tomato',255,99,71,'4763FF')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (16,'Goldenrod4',139,105,20,'14698B')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (17,'DodgerBlue1',30,144,255,'FF901E')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (18,'Green4',0,139,0,'008B00')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (19,'DeepPink1',255,20,147,'9314FF')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (20,'DarkMagenta',139,0,139,'8B008B')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (21,'RosyBrown',188,143,143,'8F8FBC')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (22,'Gold3',205,173,0,'00ADCD')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (23,'GreenYellow',173,255,47,'2FFFAD')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (24,'MediumOrchid1',224,102,255,'FF66E0')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (25,'DarkOliveGreen4',110,139,61,'3D8B6E')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
            SQL_command = "insert into 绘图颜色表(序号,英文名称,红色值,绿色值,蓝色值,HEXBGR色名) values (26,'MedSpringGreen',0,250,154,'9AFA00')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
            EXECOleDbCommand.ExecuteNonQuery()
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中 炸药 表数据升级   2025年7月18日李润洲
        ' 2025年7月18日：
        '     将炸药表和物质生成热表合并为一个表，物质生成热表中的两个字段"相对分子质量"、“生成热kj_mol"合并到炸药表中
        ' 修改字段名：生成热kj_mol->生成热kJ╱mol
        '增加备注字段
        ' 2025年7月18日前物质生成热表字段内容：
        '     物质名 TEXT(50)       ,分子式  TEXT(50)  ,相对分子质量  FLOAT,生成热kj_mol   
        ' 关键字：物质名
        '2025年7月18日前炸药表表字段内容：
        '     炸药名 TEXT(50)       ,分子式  TEXT(50)  ,C数      FLOAT,H数      FLOAT,N数      FLOAT,O数      FLOAT,爆热KJ_mol  FOLAT, 爆温K  FOLAT  
        ' 关键字：炸药名
        '
        '
        ' 2025年7月18日合并后 炸药 表字段
        ' 炸药名 TEXT(50)       ,分子式  TEXT(50)  ,C数      FLOAT,H数      FLOAT,N数      FLOAT,O数      FLOAT,爆热KJ_mol  FOLAT, 爆温K  FOLAT 
        ' 相对分子质量  FLOAT   ,生成热kJ╱mol  FLOAT
        ' 关键字：炸药名
        '2025年7月18日合并后,将物质生成热表中的两列字段值，读取并写入炸药表
        '2025年7月18日删除物质生成热表
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='炸药表'")
        If foundRows.Length <> 0 Then '有炸药表
            ReDim field_finded(4)
            sj_strchg = False
            For i = 1 To 3
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "生成热kJ╱mol" Then
                    field_finded(1) = True
                End If
                If col_name = "相对分子质量" Then
                    field_finded(2) = True
                End If
                If col_name = "备注" Then
                    field_finded(3) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                '增加字段
                SQL_command = "ALTER TABLE 炸药表 ADD COLUMN 生成热kJ╱mol  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 炸药表 set 生成热kJ╱mol=100"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 炸药表 ADD COLUMN 相对分子质量  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 炸药表 set 相对分子质量=100"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 炸药表 ADD COLUMN 备注 TEXT(250)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 炸药表 set 备注=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If Not field_finded(1) And Not field_finded(2) Then
                Dim wz_Rows() As DataRow = dbSchema.Select("TABLE_NAME='物质生成热表'")
                If wz_Rows.Length > 0 Then
                    SQL_command = "UPDATE 炸药表, 物质生成热表 SET 炸药表.生成热kJ╱mol = 物质生成热表.生成热kj_mol," _
                                  & "炸药表.相对分子质量 = 物质生成热表.相对分子质量" _
                                  & " WHERE 炸药表.炸药名 = 物质生成热表.物质名"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                    SQL_command = "DROP TABLE 物质生成热表"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                    sj_strchg = True
                Else
                    MsgBox("基础数据库中未找到【物质生成热】数据表，程序无法运行。")
                End If
            Else
                If field_finded(3) Then
                    SQL_command = "select * from 炸药表 where 炸药名='HMX'"
                    Using cmd As New OleDbCommand(SQL_command, cn_basedb)
                        Using reader As OleDbDataReader = cmd.ExecuteReader()
                            If Not reader.Read Then
                                SQL_command = "insert into 炸药表( 炸药名,分子式,C数,H数,N数,O数,相对分子质量,[生成热kJ╱mol],[爆热KJ_mol],[爆温_K],[备注])" _
                                             & " Values('HMX','C4H808N8',4,8,8,8,296,-74.89,832.75,4816,'') "
                                Using cmd2 As New OleDbCommand(SQL_command, cn_basedb)
                                    cmd2.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End Using
                    SQL_command = "select * from 炸药表 where 炸药名='PETN'"
                    Using cmd As New OleDbCommand(SQL_command, cn_basedb)
                        Using reader As OleDbDataReader = cmd.ExecuteReader()
                            If Not reader.Read Then
                                SQL_command = "insert into 炸药表( 炸药名,分子式,C数,H数,N数,O数,相对分子质量,[生成热kJ╱mol],[爆热KJ_mol],[爆温_K],[备注])" _
                                             & " Values('PETN','C5H8012N4',5,8,4,12,100,100,2007.44,3651,'') "
                                Using cmd2 As New OleDbCommand(SQL_command, cn_basedb)
                                    cmd2.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End Using
                    SQL_command = "select * from 炸药表 where 炸药名='RDX'"
                    Using cmd As New OleDbCommand(SQL_command, cn_basedb)
                        Using reader As OleDbDataReader = cmd.ExecuteReader()
                            If Not reader.Read Then
                                SQL_command = "insert into 炸药表( 炸药名,分子式,C数,H数,N数,O数,相对分子质量,[生成热kJ╱mol],[爆热KJ_mol],[爆温_K],[备注])" _
                                             & " Values('RDX','C3H606N6',3,6,6,6,222,-65.44,1383.835,4030,'') "
                                Using cmd2 As New OleDbCommand(SQL_command, cn_basedb)
                                    cmd2.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End Using
                    SQL_command = "select * from 炸药表 where 炸药名='TNT'"
                    Using cmd As New OleDbCommand(SQL_command, cn_basedb)
                        Using reader As OleDbDataReader = cmd.ExecuteReader()
                            If Not reader.Read Then
                                SQL_command = "insert into 炸药表( 炸药名,分子式,C数,H数,N数,O数,相对分子质量,[生成热kJ╱mol],[爆热KJ_mol],[爆温_K],[备注])" _
                                             & " Values('TNT','C7H506N3',7,5,3,6,227,73.22,1223.157,3050,'') "
                                Using cmd2 As New OleDbCommand(SQL_command, cn_basedb)
                                    cmd2.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End Using
                End If
            End If
        Else
            MsgBox("基础数据库中未找到【炸药】数据表，程序无法运行。")
        End If
        '*********************************************************************************************************************************
        '                                                       基础数据库中 枪弹数据 表数据升级   2025年8月2日李润洲
        ' 2025年8月2日：
        '     重命名枪弹数据表中的“射孔密度孔/m”，“装药密度g/cm3”字段，改为“射孔密度孔╱m”，“装药密度g╱cm3”,全角斜杠

        ' 升级后 枪弹数据表 字段
        '(射孔器分类 TEXT(50)  ,[射孔器名称] TEXT(50),[外径mm] FLOAT       ,[壁厚mm] FLOAT     ,[耐压MPa] FLOAT," _
        '[射孔密度孔╱m] FLOAT ,[射孔相位°] FLOAT   ,[射孔弹名称] TEXT(50),[炸药名称] TEXT(50),[装药量g] FLOAT," _
        '[装药密度g╱cm3] FLOAT,[耐温°C] FLOAT      ,[套管外径mm] FLOAT   ,[平均孔径mm] FLOAT ,[平均穿深mm] FLOAT," _
        ' [流动效率] FLOAT      ,[钢靶孔径mm] FLOAT   ,[钢靶穿深mm] FLOAT   ,[备注] TEXT(250))"
        ' 关键字：射孔器名称、射孔弹名称、炸药名称
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='枪弹数据表'")
        If foundRows.Length <> 0 Then '有枪弹数据表
            ReDim field_finded(2)
            sj_strchg = False
            For i = 1 To 2
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_basedb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "射孔密度孔/m" Then
                    field_finded(1) = True
                End If
                If col_name = "装药密度g/cm3" Then
                    field_finded(2) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) Then
                SQL_command = "ALTER TABLE 枪弹数据表 ADD COLUMN 射孔密度孔╱m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 枪弹数据表 set 射孔密度孔╱m=[射孔密度孔/m]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "ALTER TABLE 枪弹数据表 DROP COLUMN [射孔密度孔/m]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
            If field_finded(2) Then
                SQL_command = "ALTER TABLE 枪弹数据表 ADD COLUMN 装药密度g╱cm3 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "update 枪弹数据表 set 装药密度g╱cm3=[装药密度g/cm3]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                SQL_command = "ALTER TABLE 枪弹数据表 DROP COLUMN [装药密度g/cm3]"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_basedb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                sj_strchg = True
            End If
        End If


        dbSchema.Clear()
        dbSchema.Dispose()
        cn_basedb.Close()
        cn_basedb.Dispose()
    End Sub
    '*********************************************************************************************************************************
    '                                    基础数据数据结构升级与新建   结束
    '*********************************************************************************************************************************

    '*********************************************************************************************************************************
    '                                                       用户数据数据结构升级  开始
    '*********************************************************************************************************************************
    Sub user_database_update()
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim columnTable As DataTable
        Dim columnTab_row As DataRow
        Dim table_name As String
        Dim col_name As String
        Dim i As Short
        Dim field_finded() As Boolean
        Dim sj_strchg As Boolean '数据结构有升级标志
        Dim SQL_command As String
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        '*********************************************************************************************************************************
        '                                                       用户数据库中 油气井表数据结构升级   李润洲2025年7月21日
        '1.为计算裸眼射孔时的爆轰参数，2025年7月21日增加“完钻钻头尺寸mm FLOAT”1个字段。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='油气井表'")
        If foundRows.Length <> 0 Then '有工况_开关元件表
            ReDim field_finded(1)
            sj_strchg = False
            For i = 1 To 1
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "完钻钻头尺寸mm" Then
                    field_finded(1) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 油气井表 ADD COLUMN 完钻钻头尺寸mm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 油气井表 set 完钻钻头尺寸mm=168.3"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("油气井基本数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查工况数据中有关开关工具的参数是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 作业地层参数表数据结构升级   李润洲2025年10月23日
        '2025年7月改进这个表时没有写升级代码，增加“目的层名称” 1个字段
        '2025年10月23日增加“人工井底m FLOAT”1个字段。因为“人工井底m”字段放在目的层参数表中不合理，迁移到作业地层参数表中。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='作业地层参数表'")
        If foundRows.Length <> 0 Then '有作业地层参数表
            ReDim field_finded(2)
            sj_strchg = False
            For i = 1 To 2
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "目的层名称" Then
                    field_finded(1) = True
                End If
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "人工井底m" Then
                    field_finded(2) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 作业地层参数表 ADD COLUMN 目的层名称 TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 作业地层参数表 set 目的层名称='暂定名称'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 作业地层参数表 ADD COLUMN 人工井底m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 作业地层参数表 set 人工井底m=12000"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("作业地层数据已升级，新增的数据项可能不准确。计算前，请一一检查参数是否合理。")
            End If
        End If

        '*********************************************************************************************************************************
        '                                                       用户数据库中 目的层参数表数据结构升级   李润洲2025年10月23日
        '2025年10月，将表中的“人工井底m”迁移到作业地层参数表中，删除目的层参数表的该字段。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='目的层参数表'")
        If foundRows.Length <> 0 Then '有作业地层参数表
            sj_strchg = False
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "人工井底m" Then
                    SQL_command = "ALTER TABLE 目的层参数表 DROP COLUMN 人工井底m"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    sj_strchg = True
                    Exit For
                End If
            Next
            columnTable.Dispose()
            If sj_strchg = True Then
                MsgBox("目的层参数表中的人工井底迁移到作业地层界面，迁移后的数据不准确。计算前，请检查参数是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       工况参数升级
        ' 1. 工况参数变化情况
        '    2015年2月22日再次修订程序，增“井口管压开关”、“井口套压开关”两字段，用以描述井口压力计算方式。
        '    2015年2月27日再次修订程序，改字段“井底深度m”为“管压井底深度m”，增“套压井底深度m”，用以描述井底压力对应的下深。
        '    2015年4月1日再次修订程序，增“环空流阻模型”、“管内流阻模型”、“环空稠剂浓度”、“管内稠剂浓度”、
        '                                “环空撑剂浓度”、“管内撑剂浓度”、“环空流变指数n”、“管内流变指数n”、
        '                                “环空稠度系数K”、“管内稠度系数K”10个字段用以进行流体摩阻计算。
        '    2015年7月19日再次修订程序，增“井口加扭Nm”、 “井底约束情况”、“井底约束位置m”、“约束定位方式”4个字段用以解卡、钻磨计算。
        '                               改“动作参数”为“井口加力kN”
        '                               删“井口情况”
        '    2015年7月26日再次修订程序，增“管牛模折减系数”、 “环牛模折减系数”2个字段用以摩阻计算。
        '    2016年10月24日再次修订程序，增“管流时间h”字段，用以疲劳寿命分析
        '
        '2. 工况参数升级内容：
        '  （1）删除“工况参数表”中“井口情况”字段
        '  （2）删除“工况参数表”中“动作参数”字段
        '  （3）增加21个新字段
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='工况参数表'")
        If foundRows.Length <> 0 Then '有工况参数表
            ReDim field_finded(22)
            sj_strchg = False
            For i = 1 To 22
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "井底深度m" Then
                    SQL_command = "ALTER TABLE 工况参数表 DROP COLUMN 井底深度m"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    sj_strchg = True
                End If

                If col_name = "井口情况" Then
                    SQL_command = "ALTER TABLE 工况参数表 DROP COLUMN 井口情况"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    sj_strchg = True
                End If

                If col_name = "动作参数" Then
                    SQL_command = "ALTER TABLE 工况参数表 DROP COLUMN 动作参数"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    sj_strchg = True
                End If
                If col_name = "井口管压开关" Then
                    field_finded(1) = True
                End If
                If col_name = "井口套压开关" Then
                    field_finded(2) = True
                End If
                If col_name = "管压井底深度m" Then
                    field_finded(3) = True
                End If
                If col_name = "套压井底深度m" Then
                    field_finded(4) = True
                End If
                If col_name = "环空流阻模型" Then
                    field_finded(5) = True
                End If
                If col_name = "管内流阻模型" Then
                    field_finded(6) = True
                End If
                If col_name = "环空稠剂浓度" Then
                    field_finded(7) = True
                End If
                If col_name = "管内稠剂浓度" Then
                    field_finded(8) = True
                End If
                If col_name = "环空撑剂浓度" Then
                    field_finded(9) = True
                End If
                If col_name = "管内撑剂浓度" Then
                    field_finded(10) = True
                End If
                If col_name = "环空流变指数n" Then
                    field_finded(11) = True
                End If
                If col_name = "管内流变指数n" Then
                    field_finded(12) = True
                End If
                If col_name = "环空稠度系数K" Then
                    field_finded(13) = True
                End If
                If col_name = "管内稠度系数K" Then
                    field_finded(14) = True
                End If
                If col_name = "井口加力kN" Then
                    field_finded(15) = True
                End If
                If col_name = "井口加扭Nm" Then
                    field_finded(16) = True
                End If
                If col_name = "井底约束情况" Then
                    field_finded(17) = True
                End If
                If col_name = "井底约束位置m" Then
                    field_finded(18) = True
                End If
                If col_name = "约束定位方式" Then
                    field_finded(19) = True
                End If
                If col_name = "管牛模折减系数" Then
                    field_finded(20) = True
                End If
                If col_name = "环牛模折减系数" Then
                    field_finded(21) = True
                End If
                If col_name = "管流时间h" Then
                    field_finded(22) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井口管压开关 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井口管压开关='输入'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井口套压开关 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井口套压开关='输入'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管压井底深度m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管压井底深度m=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 套压井底深度m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 套压井底深度m=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环空流阻模型 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环空流阻模型='无'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管内流阻模型 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管内流阻模型='无'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环空稠剂浓度 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环空稠剂浓度=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管内稠剂浓度    FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管内稠剂浓度=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环空撑剂浓度  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环空撑剂浓度=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管内撑剂浓度    FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管内撑剂浓度=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环空流变指数n   FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环空流变指数n=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管内流变指数n   FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管内流变指数n=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(13) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环空稠度系数K   FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环空稠度系数K=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(14) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管内稠度系数K FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管内稠度系数K=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(15) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井口加力kN      FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井口加力kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(16) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井口加扭Nm      FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井口加扭Nm=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(17) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井底约束情况 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井底约束情况='无'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(18) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 井底约束位置m   FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 井底约束位置m=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(19) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 约束定位方式 TEXT(20)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 约束定位方式=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(20) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管牛模折减系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管牛模折减系数=0.35"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(21) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 环牛模折减系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 环牛模折减系数=0.35"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(22) = False Then
                SQL_command = "ALTER TABLE 工况参数表 ADD COLUMN 管流时间h FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况参数表 set 管流时间h=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("工况参数数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查工况参数数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 工况_开关元件表数据结构升级
        '1.为计算流体过开关元件后的压力，2017年7月17日增加“管内嘴损压差MPa FLOAT,油套嘴损压差MPa FLOAT”2个字段。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='工况_开关元件'")
        If foundRows.Length <> 0 Then '有工况_开关元件表
            ReDim field_finded(2)
            sj_strchg = False
            For i = 1 To 2
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "管内嘴损压差MPa" Then
                    field_finded(1) = True
                End If
                If col_name = "油套嘴损压差MPa" Then
                    field_finded(2) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 工况_开关元件 ADD COLUMN 管内嘴损压差MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况_开关元件 set 管内嘴损压差MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 工况_开关元件 ADD COLUMN 油套嘴损压差MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 工况_开关元件 set 油套嘴损压差MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("工况_开关工具数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查工况数据中有关开关工具的参数是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 井身数据结构升级
        '1. 井身结构数据变化情况
        '    2015年10月31日，修订程序，增加了“钻头尺寸mm”、“扣型”、“套管名称”3个字段。
        '    2017年07月21日，修订程序，增加了“单位长质量kgpm”、“接头抗内压强度MPa”、“接头抗拉强度kN”、“备注”4个字段。
        '    2023年12月25日，修订程序，增加了“固井时井液密度“、”固井水泥浆密度”2个字段
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='套管数据表'")
        If foundRows.Length <> 0 Then '有套管数据表
            ReDim field_finded(9)
            sj_strchg = False
            For i = 1 To 9
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "钻头尺寸mm" Then
                    field_finded(1) = True
                End If

                If col_name = "扣型" Then
                    field_finded(2) = True
                End If

                If col_name = "套管名称" Then
                    field_finded(3) = True
                End If
                If col_name = "单位长质量kgpm" Then
                    field_finded(4) = True
                End If

                If col_name = "接头抗内压强度MPa" Then
                    field_finded(5) = True
                End If
                If col_name = "接头抗拉强度kN" Then
                    field_finded(6) = True
                End If

                If col_name = "备注" Then
                    field_finded(7) = True
                End If
                If col_name = "固井时井液密度" Then
                    field_finded(8) = True
                End If
                If col_name = "固井水泥浆密度" Then
                    field_finded(9) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 钻头尺寸mm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 钻头尺寸mm=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 扣型 TEXT(30)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 套管名称 TEXT(30)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 套管名称=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 单位长质量kgpm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 单位长质量kgpm=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 接头抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 接头抗内压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 接头抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 接头抗拉强度kN=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 备注 TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 固井时井液密度 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 固井时井液密度=1.3"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 套管数据表 ADD COLUMN 固井水泥浆密度 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 套管数据表 set 固井水泥浆密度=1.8"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("井身结构数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查井身结构数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱_油管表 结构升级
        ' 2018年7月30日增加字段
        '   (01)扣型              TEXT(50)
        '   (02)接头抗内压强度MPa FLOAT
        '   (03)接头抗拉强度kN    FLOAT
        '   (04)备注              TEXT(250)
        ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
        '   (05)热膨胀系数    FLOAT
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='管柱_油管表'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(5)
            sj_strchg = False
            For i = 1 To 5
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "扣型" Then
                    field_finded(1) = True
                End If
                If col_name = "接头抗内压强度MPa" Then
                    field_finded(2) = True
                End If
                If col_name = "接头抗拉强度kN" Then
                    field_finded(3) = True
                End If
                If col_name = "备注" Then
                    field_finded(4) = True
                End If
                If col_name = "热膨胀系数" Then
                    field_finded(5) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 管柱_油管表 ADD COLUMN 扣型  TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_油管表 set 扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 管柱_油管表 ADD COLUMN 接头抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_油管表 set 接头抗内压强度MPa=抗内压强度MPa"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 管柱_油管表 ADD COLUMN 接头抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_油管表 set 接头抗拉强度kN=抗拉强度kN"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 管柱_油管表 ADD COLUMN 备注 TEXT(250)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_油管表 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            '*********************************************************************************************************************************************
            ' 其中材料的线膨胀系数取值依据见：化学工业出版社，成大先主编《机械设计手册》第四版 第一卷 1-11页表1-1-14。
            ' 在本程序中采用定值，取各温度段的平均值的平均值。由于手册中取值范围较大，故：
            '    按窦老师2015年2月10日提供的“00 高难度复杂井完井（试油）油套管柱力学分析要点及若干工程问题简析（2011年5月11日西安完井会议）.ppt”中“温度
            ' 效应”公式及热膨胀系数（1.2*10^(-5)）计算公式取默认值。
            '    对于钢，WELLCAT中为12.4*10E-6/℃。
            '*********************************************************************************************************************************************
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 管柱_油管表 ADD COLUMN 热膨胀系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_油管表 set 热膨胀系数=1.24E-5"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("管柱中管柱_油管表数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查管柱中油井管数据是否合理。")
            End If
        End If

        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱_普通钻杆 结构升级
        ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
        '   (05)热膨胀系数    FLOAT
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='管柱_普通钻杆'")
        If foundRows.Length <> 0 Then '有管柱_普通钻杆
            ReDim field_finded(1)
            sj_strchg = False
            For i = 1 To 1
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "热膨胀系数" Then
                    field_finded(1) = True
                End If
            Next
            columnTable.Dispose()
            '*********************************************************************************************************************************************
            ' 其中材料的线膨胀系数取值依据见：化学工业出版社，成大先主编《机械设计手册》第四版 第一卷 1-11页表1-1-14。
            ' 在本程序中采用定值，取各温度段的平均值的平均值。由于手册中取值范围较大，故：
            '    按窦老师2015年2月10日提供的“00 高难度复杂井完井（试油）油套管柱力学分析要点及若干工程问题简析（2011年5月11日西安完井会议）.ppt”中“温度
            ' 效应”公式及热膨胀系数（1.2*10^(-5)）计算公式取默认值。
            '    对于钢，WELLCAT中为12.4*10E-6/℃。
            '*********************************************************************************************************************************************
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 管柱_普通钻杆 ADD COLUMN 热膨胀系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_普通钻杆 set 热膨胀系数=1.24E-5"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("管柱中管柱_普通钻杆数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查管柱中油井管数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 计算参数表数据结构升级
        '1. 计算参数表结构数据变化情况
        '    2015年11月15日，修订程序，将“参数值”字段长度由20更改为200。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='计算参数表'")
        If foundRows.Length <> 0 Then '有计算参数表
            ReDim field_finded(1)
            sj_strchg = False
            For i = 1 To 1
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "参数值" Then
                    field_finded(1) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = True Then
                SQL_command = "ALTER TABLE 计算参数表 ALTER COLUMN 参数值 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱_开关元件表数据结构升级
        '2017年10月28日：
        '    由于基础数据库中开关元件数据结构变化，需对用户数据库中 管柱_开关元件表进行升级。
        '   （1）增加以下9个字段：
        '       01       温度范围下℃     FLOAT
        '       02       温度范围上℃     FLOAT
        '       03       压力等级MPa      FLOAT
        '       04       型号             TEXT(100)
        '       05       开关类型         TEXT(100)
        '       06       上端扣型         TEXT(100)
        '       07       下端扣型         TEXT(100)
        '       08       生产厂家         TEXT(100)
        '       09       备注             TEXT(100)
        '   （2）放弃并删除使用2个字段：开关方式，开关流向。注：开关流向在工况中考虑
        '    (3) 关键字：[外径(mm)]，[内径(mm)]，[长度)(m)],[型号]，[开关类型]
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='管柱_开关元件'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(11)
            sj_strchg = False
            For i = 1 To 11
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If
                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If
                If col_name = "压力等级MPa" Then
                    field_finded(3) = True
                End If
                If col_name = "型号" Then
                    field_finded(4) = True
                End If
                If col_name = "开关类型" Then
                    field_finded(5) = True
                End If
                If col_name = "上端扣型" Then
                    field_finded(6) = True
                End If
                If col_name = "下端扣型" Then
                    field_finded(7) = True
                End If
                If col_name = "生产厂家" Then
                    field_finded(8) = True
                End If
                If col_name = "备注" Then
                    field_finded(9) = True
                End If
                If col_name = "开关方式" Then
                    field_finded(10) = True
                End If
                If col_name = "开关流向" Then
                    field_finded(11) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 压力等级MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 压力等级MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 型号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 型号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 开关类型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 开关类型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 管柱_开关元件 ADD COLUMN 备注 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_开关元件 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = True Then
                SQL_command = "ALTER TABLE 管柱_开关元件 DROP COLUMN 开关方式"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = True Then
                SQL_command = "ALTER TABLE 管柱_开关元件 DROP COLUMN 开关流向"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
            End If
            If sj_strchg = True Then
                MsgBox("管柱_开关元件数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查管柱中开关元件数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱_伸缩元件表数据结构升级
        '2016年10月10日：
        '    由于基础数据库中伸缩元件数据结构变化，需对用户数据库中 管柱_伸缩元件表进行升级。
        '   （1）增加以下13个字段：
        '       01       温度范围下℃     FLOAT
        '       02       温度范围上℃     FLOAT
        '       03       压力等级MPa      FLOAT
        '       04       抗拉强度kN       FLOAT
        '       05       上端扣型         TEXT(100)
        '       06       下端扣型         TEXT(100)
        '       07       主体材料         TEXT(100)
        '       08       主材屈服强度MPa  FLOAT
        '       09       零件号           TEXT(100)
        '       10       生产厂家         TEXT(100)
        '       11       备注             TEXT(200)
        '       12       全缩短长度m      FLOAT
        '       13       伸缩动作载荷kN   FLOAT
        '   （2）放弃使用2个字段：中心管外径(mm)，中心管内径(mm)
        '   （3）在管柱-伸缩元件用户表中“长度”指下入长度，其值大于等于全缩短长度m。
        '        注意：计算出的 全拉开长度m=全缩短长度m+伸缩行程；全缩短长度m<=下入长度<=全拉开长度m
        '    (4) 由于同一伸缩管上装不同的剪切销钉其动作载荷不同，故将“伸缩动作载荷kN”字段放在管柱_伸缩元件表中。
        '2024年9月30日：
        '   （1）增加以下1个字段：
        '       14       伸缩动作压力MPa   FLOAT
        '    (2) 由于同一伸缩管上装不同的剪切销钉其动作压力不同，故将“伸缩动作压力MPa”字段放在管柱_伸缩元件表中。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='管柱_伸缩元件'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(14)
            sj_strchg = False
            For i = 1 To 14
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If

                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If

                If col_name = "压力等级MPa" Then
                    field_finded(3) = True
                End If

                If col_name = "抗拉强度kN" Then
                    field_finded(4) = True
                End If

                If col_name = "上端扣型" Then
                    field_finded(5) = True
                End If

                If col_name = "下端扣型" Then
                    field_finded(6) = True
                End If

                If col_name = "主体材料" Then
                    field_finded(7) = True
                End If

                If col_name = "主材屈服强度MPa" Then
                    field_finded(8) = True
                End If

                If col_name = "零件号" Then
                    field_finded(9) = True
                End If

                If col_name = "生产厂家" Then
                    field_finded(10) = True
                End If

                If col_name = "备注" Then
                    field_finded(11) = True
                End If
                If col_name = "全缩短长度m" Then
                    field_finded(12) = True
                End If
                If col_name = "伸缩动作载荷kN" Then
                    field_finded(13) = True
                End If
                If col_name = "伸缩动作压力MPa" Then
                    field_finded(14) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 压力等级MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 压力等级MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 抗拉强度kN=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 主体材料 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 主体材料=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 主材屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 主材屈服强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 零件号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 零件号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 备注 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 全缩短长度m FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 全缩短长度m=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(13) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 伸缩动作载荷kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 伸缩动作载荷kN=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(14) = False Then
                SQL_command = "ALTER TABLE 管柱_伸缩元件 ADD COLUMN 伸缩动作压力MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_伸缩元件 set 伸缩动作压力MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("管柱_伸缩元件数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查管柱中伸缩元件数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱_封隔定位元件表数据结构升级
        '2016年9月16日：
        '    由于基础数据库中封隔器数据结构变化，需对用户数据库中 管柱_封隔定位元件表进行升级。
        '   （1）增加以下13个字段：
        '       温度范围下℃     FLOAT
        '       温度范围上℃     FLOAT
        '       抗内压强度MPa    FLOAT
        '       抗外压强度MPa    FLOAT
        '       型号             TEXT(100)
        '       生产厂家         TEXT(100)
        '       上端扣型         TEXT(100)
        '       下端扣型         TEXT(100)
        '       主体材料         TEXT(100)
        '       主材屈服强度MPa  FLOAT
        '       最小坐封压力MPa  FLOAT
        '       最大坐封压力MPa  FLOAT
        '       备注             TEXT(200)
        '2017年2月4日
        '    (1) 增加字段 “能否反洗井”，TEXT(10)
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='管柱_封隔定位元件'")
        If foundRows.Length <> 0 Then '有开关元件表
            ReDim field_finded(14)
            sj_strchg = False
            For i = 1 To 14
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "温度范围下℃" Then
                    field_finded(1) = True
                End If

                If col_name = "温度范围上℃" Then
                    field_finded(2) = True
                End If

                If col_name = "抗内压强度MPa" Then
                    field_finded(3) = True
                End If

                If col_name = "抗外压强度MPa" Then
                    field_finded(4) = True
                End If

                If col_name = "型号" Then
                    field_finded(5) = True
                End If

                If col_name = "生产厂家" Then
                    field_finded(6) = True
                End If

                If col_name = "上端扣型" Then
                    field_finded(7) = True
                End If

                If col_name = "下端扣型" Then
                    field_finded(8) = True
                End If

                If col_name = "主体材料" Then
                    field_finded(9) = True
                End If

                If col_name = "主材屈服强度MPa" Then
                    field_finded(10) = True
                End If

                If col_name = "最小坐封压力MPa" Then
                    field_finded(11) = True
                End If

                If col_name = "最大坐封压力MPa" Then
                    field_finded(12) = True
                End If

                If col_name = "备注" Then
                    field_finded(13) = True
                End If

                If col_name = "能否反洗井" Then
                    field_finded(14) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 温度范围下℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 温度范围下℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 温度范围上℃ FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 温度范围上℃=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 抗内压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 抗外压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 抗外压强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 型号 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 型号=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 生产厂家 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 生产厂家=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 上端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 上端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 下端扣型 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 下端扣型=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 主体材料 TEXT(100)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 主体材料=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 主材屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 主材屈服强度MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 最小坐封压力MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 最小坐封压力MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(12) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 最大坐封压力MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 最大坐封压力MPa=0.0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(13) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 备注 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 备注=' '"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(14) = False Then
                SQL_command = "ALTER TABLE 管柱_封隔定位元件 ADD COLUMN 能否反洗井 TEXT(10)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 管柱_封隔定位元件 set 能否反洗井='不能'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("管柱_封隔定位元件数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查管柱中封隔器数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 管柱数据表数据升级
        '2017年春节期间（2月3日开始）
        '     考虑用户习惯，将管柱元件分类重新定义，管柱元件由原来的：油管、开关元件、封隔定位元件、节流元件、伸缩元件5种，调整修改为:
        ' 油井管、开关工具、封隔器、节流工具、伸缩管、锚定工具6种。
        '     (1)油井管  --原来类型名为 油管，范围小，没有涵盖套管、钻杆等管材，现改为油井管。压力、轴力传递连续
        '     (2)节流工具--原来类型名为 节流元件。考虑到现场习惯改为节流工具，该类工具保持管内或内外常同，流体流过时压力有
        ' 显著的降低。压力传递有突变，轴力传递连续。
        '     (3)封隔器----原来类型名为 封隔定位元件。考虑到封隔器的特殊性，单独设类。坐封后，隔断环空压力，还可能有锚定功能。
        ' 坐封后，环空压力传递不连续或（有洗井功能的洗井时）突变，轴力传递突变（无锚定，考虑胶筒摩擦）或不连续（有锚定）
        '     (4)开关工具--原来类型名为 开关元件。开关工具在不同工况下可有8种状态，以连通或断开压力传递，还有节流功能。压力传递
        ' （通道打开节流突变）连续或（通道关闭）不连续。
        '     (5)伸缩管----原来类型名为 伸缩工具。压力传递连续，但动作时，轴力传递不连续。
        '     (6)锚定工具--新增，压力连续，坐卡后轴力不连续。
        '为此，原来用户数据库管柱数据表中元件性质字段值需做相应的修改：
        '将“油管”改为“油井管”；
        '将“开关元件”改为“开关工具”；
        '将“封隔定位元件”改为“封隔器”；
        '将“节流元件”改为“节流工具”；
        '将“伸缩元件”改为“伸缩管”。
        '                                                            --2017年2月14日 秦彦斌
        '*********************************************************************************************************************************
        SQL_command = "update 管柱数据表 set 元件性质='油井管' where 元件性质='油管'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        SQL_command = "update 管柱数据表 set 元件性质='开关工具' where 元件性质='开关元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        SQL_command = "update 管柱数据表 set 元件性质='封隔器' where 元件性质='封隔定位元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        SQL_command = "update 管柱数据表 set 元件性质='节流工具' where 元件性质='节流元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        SQL_command = "update 管柱数据表 set 元件性质='伸缩管' where 元件性质='伸缩元件'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        '*********************************************************************************************************************************
        '                                                       用户数据库中 钻具组合表升级
        '1. 钻具组合表变化情况
        '    2017年5月6日，修订程序，增加了“用途 TEXT(10)”、“起下次数  INTEGER”2个字段，用以计算起下钻、钻塞磨铣时钻柱对套管的磨损。
        '    程序中,用途值为：“正常钻进”、“钻塞磨铣”、“开窗侧钻？--想想”
        '    2021年5月30日，修订程序：增加“已固井深 FLOAT”字段，用以计算起下钻、钻塞磨铣时钻柱对套管的磨损。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='钻具组合表'")
        If foundRows.Length <> 0 Then '有钻具组合表
            ReDim field_finded(3)
            sj_strchg = False
            For i = 1 To 3
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "用途" Then
                    field_finded(1) = True
                End If
                If col_name = "起下次数" Then
                    field_finded(2) = True
                End If
                If col_name = "已固井深" Then
                    field_finded(3) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 钻具组合表 ADD COLUMN 用途 TEXT(10)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合表 set 用途='正常钻进'"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 钻具组合表 ADD COLUMN 起下次数  INTEGER"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合表 set 起下次数=1"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 钻具组合表 ADD COLUMN 已固井深 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合表 set 已固井深=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            SQL_command = "update 钻具组合表 set set 用途='正常钻进' where 用途='钻井'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            If sj_strchg = True Then
                MsgBox("钻具组合表数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查钻具组合数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 钻井日志表升级
        '1. 钻井日志表表变化情况
        '    2021年5月28日，修订程序，增加了“已固井深 FLOAT”字段，用以计算套管磨损时，判断套管磨损属于正常钻进磨损还是钻塞、修井作业
        ' 钻柱对套管的磨损。
        '    钻井日志中，当钻头深度大于已固井深，表示此条日志为正常钻进，磨损的套管为下入深度小于钻头深度的最内层套管。
        '    钻井日志中，当钻头深度小于等于已固井深，表示此条日志为钻塞磨铣，磨损的套管为下入深度小于等于已固井深的最内层套管。
        '*********************************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='钻井日志表'")
        If foundRows.Length <> 0 Then '钻井日志表
            ReDim field_finded(2)
            sj_strchg = False
            For i = 1 To 2
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "已固井深" Then
                    field_finded(1) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 钻井日志表 ADD COLUMN 已固井深 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻井日志表 set 已固井深=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("钻井日志表中已加【已固井深】字段，该字段用于判断该条日志属于正常钻进还是钻塞磨铣钻进。若进行套管磨损分析，请修改该井的钻井日志导入表，按模板格式增加【已固井深】数据，重新导入钻井日志，程序方可正常运行。")
            End If
        End If



        '*********************************************************************************************************************************
        '                                                       用户数据库中 钻具组合_普通钻杆表升级
        ' 根据2018年3月18日钻杆基础数据变化情况：
        '2018年3月22日数据结构升级情况：
        '增加字段
        '   (01)扣型              TEXT(50)
        '   (02)管体抗扭强度Nm    FLOAT
        '   (03)接头抗扭强度Nm    FLOAT
        '   (04)管体抗拉强度kN    FLOAT
        '   (05)接头抗拉强度kN    FLOAT
        '   (06)抗内压强度MPa     FLOAT
        '   (07)抗挤强度MPa       FLOAT
        '   (08)接头外径mm        FLOAT
        '   (09)屈服强度MPa       FLOAT
        '   (10)备注              TEXT(200)
        '   (11)将“最小抗拉强度kN”值赋给“管体抗拉强度kN”后删除
        '   (12)将“最小屈服强度MPa”值赋给“屈服强度MPa”后删除
        ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
        '   (13)热膨胀系数    FLOAT
        '****************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='钻具组合_普通钻杆表'")
        If foundRows.Length <> 0 Then '钻具组合_普通钻杆表
            ReDim field_finded(13)
            sj_strchg = False
            For i = 1 To 13
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "扣型" Then
                    field_finded(1) = True
                End If
                If col_name = "管体抗扭强度Nm" Then
                    field_finded(2) = True
                End If
                If col_name = "接头抗扭强度Nm" Then
                    field_finded(3) = True
                End If
                If col_name = "管体抗拉强度kN" Then
                    field_finded(4) = True
                End If
                If col_name = "接头抗拉强度kN" Then
                    field_finded(5) = True
                End If
                If col_name = "抗内压强度MPa" Then
                    field_finded(6) = True
                End If
                If col_name = "抗挤强度MPa" Then
                    field_finded(7) = True
                End If
                If col_name = "接头外径mm" Then
                    field_finded(8) = True
                End If
                If col_name = "屈服强度MPa" Then
                    field_finded(9) = True
                End If
                If col_name = "备注" Then
                    field_finded(10) = True
                End If
                If col_name = "最小抗拉强度kN" Then
                    field_finded(11) = True
                End If
                If col_name = "最小屈服强度MPa" Then
                    field_finded(12) = True
                End If
                If col_name = "热膨胀系数" Then
                    field_finded(13) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 扣型 TEXT(50)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 扣型=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 管体抗扭强度Nm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 管体抗扭强度Nm=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 接头抗扭强度Nm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 接头抗扭强度Nm=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 管体抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 管体抗拉强度kN=最小抗拉强度kN"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 接头抗拉强度kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 接头抗拉强度kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 抗内压强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 抗内压强度MPa=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(7) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 抗挤强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 抗挤强度MPa=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(8) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 接头外径mm FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 接头外径mm=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(9) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 屈服强度MPa FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 屈服强度MPa=最小屈服强度MPa"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(10) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 备注 TEXT(200)"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 备注=''"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(11) = True Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 DROP COLUMN 最小抗拉强度kN"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(12) = True Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 DROP COLUMN 最小屈服强度MPa"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(13) = False Then
                SQL_command = "ALTER TABLE 钻具组合_普通钻杆表 ADD COLUMN 热膨胀系数 FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 钻具组合_普通钻杆表 set 热膨胀系数=1.24E-5"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("钻具组合_普通钻杆表数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确。计算前，请一一检查钻具组合中普通钻杆数据是否合理。")
            End If
        End If
        '*********************************************************************************************************************************
        '                                                       用户数据库中 重量及拉伸应力安全系数表升级
        '2023年11月19日数据结构升级情况：
        '增加字段
        '   (01)空气中累重kN          FLOAT
        '   (02)液体中累重kN          FLOAT
        '   (03)空气中管体剩余拉力kN  FLOAT
        '   (04)液体中管体剩余拉力kN  FLOAT
        '   (05)空气中接头剩余拉力kN  FLOAT
        '   (06)液体中接头剩余拉力kN  FLOAT
        '****************************************************************************************************************
        foundRows = dbSchema.Select("TABLE_NAME='重量及拉伸应力安全系数'")
        If foundRows.Length <> 0 Then '有“重量及拉伸应力安全系数”表
            ReDim field_finded(6)
            sj_strchg = False
            For i = 1 To 6
                field_finded(i) = False
            Next i
            table_name = foundRows(0).Item("TABLE_NAME").ToString
            columnTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
            For Each columnTab_row In columnTable.Rows
                col_name = columnTab_row.Item("COLUMN_NAME").ToString
                If col_name = "空气中累重kN" Then
                    field_finded(1) = True
                End If
                If col_name = "液体中累重kN" Then
                    field_finded(2) = True
                End If
                If col_name = "空气中管体剩余拉力kN" Then
                    field_finded(3) = True
                End If
                If col_name = "液体中管体剩余拉力kN" Then
                    field_finded(4) = True
                End If
                If col_name = "空气中接头剩余拉力kN" Then
                    field_finded(5) = True
                End If
                If col_name = "液体中接头剩余拉力kN" Then
                    field_finded(6) = True
                End If
            Next
            columnTable.Dispose()
            If field_finded(1) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 空气中累重kN  FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 空气中累重kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(2) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 液体中累重kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 液体中累重kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(3) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 空气中管体剩余拉力kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 空气中管体剩余拉力kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(4) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 液体中管体剩余拉力kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 液体中管体剩余拉力kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(5) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 空气中接头剩余拉力kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 空气中接头剩余拉力kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If field_finded(6) = False Then
                SQL_command = "ALTER TABLE 重量及拉伸应力安全系数 ADD COLUMN 液体中接头剩余拉力kN FLOAT"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                SQL_command = "update 重量及拉伸应力安全系数 set 液体中接头剩余拉力kN=0"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                sj_strchg = True
            End If
            If sj_strchg = True Then
                MsgBox("重量及拉伸应力安全系数表数据已升级，升级程序最大限度保存已有数据，但新增的数据项有可能不准确，可重新计算涮新不准确的数据项。")
            End If
        End If
        dbSchema.Clear()
        dbSchema.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Sub
    '*********************************************************************************************************************************
    '                                                       用户数据数据结构升级  结束
    '*********************************************************************************************************************************

    '*********************************************************************************************************************************
    ' 建立数据表
    '*********************************************************************************************************************************
    Sub database_creat(ByVal opt_code As Short, ByVal DBopt_select As Short)
        Dim cn As System.Data.OleDb.OleDbConnection
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        If Not (DBopt_select = 1 Or DBopt_select = 2) Then
            Exit Sub
        End If
        cn = Nothing
        Select Case DBopt_select
            Case 1
                cn = New System.Data.OleDb.OleDbConnection(AdoConString)
            Case 2
                cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        End Select
        cn.Open()
        dbSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        Select Case opt_code
            Case 1
                foundRows = dbSchema.Select("TABLE_NAME='油气井表'")
                If foundRows.Length = 0 Then '没有油气井表
                    '*********************************************************************************************************************************
                    '2025年7月21日增加“完钻钻头尺寸mm FLOAT”1个字段。                        李润洲
                    '*********************************************************************************************************************************
                    SQL_command = "CREATE TABLE 油气井表 " _
                        & "(井号 TEXT(50)," & " 井别 TEXT(50), " & " 地理位置 TEXT(200), " & " 构造位置 TEXT(200), " & " 完钻层位 TEXT(200), " _
                        & " 设计井深m FLOAT, " & " 完钻井深m FLOAT, " & "完钻钻头尺寸mm FLOAT, " & " 开钻日期 DATETIME, " & " 完钻日期 DATETIME," & " 地温梯度℃pbm  FLOAT, " _
                        & " 地压梯度MPapbm FLOAT)"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 2
                foundRows = dbSchema.Select("TABLE_NAME='套管数据表'")
                If foundRows.Length = 0 Then '没有套管数据表
                    '*************************************************************************************************************************
                    '    2015年10月31日，修订程序，增加了“钻头尺寸mm”、“扣型”、“套管名称”3个字段。
                    '    2017年07月21日，修订程序，增加了“单位长质量kgpm”、“接头抗内压强度MPa”、“接头抗拉强度kN”、“备注”4个字段。
                    '    2023年12月25日，修订程序，增加了“固井时井液密度“、”固井水泥浆密度”2个字段
                    '*************************************************************************************************************************
                    SQL_command = "CREATE TABLE 套管数据表 " _
                        & "(套管规格       TEXT(50),套管类型       TEXT(30),套管层数       INTEGER ,套管段数    INTEGER ,是否回接      TEXT(20)," _
                        & " 套管外径mm     FLOAT   ,套管壁厚mm     FLOAT   ,悬挂深度m         FLOAT,套管下深m      FLOAT,水泥返深m     FLOAT   ," _
                        & " 完钻深度m      FLOAT   ,套管钢级       TEXT(50),弹性模量MPa       FLOAT,抗内压强度MPa  FLOAT,抗拉强度kN    FLOAT   ," _
                        & " 抗挤强度MPa    FLOAT   ,泊松比         FLOAT   ,屈服极限MPa       FLOAT,钻头尺寸mm     FLOAT,扣型          TEXT(30)," _
                        & " 套管名称       TEXT(30),单位长质量kgpm FLOAT   ,接头抗内压强度MPa FLOAT,接头抗拉强度kN FLOAT,备注          TEXT(50)," _
                        & " 固井时井液密度 FLOAT   ,固井水泥浆密度 FLOAT   ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 3
                foundRows = dbSchema.Select("TABLE_NAME='管柱数据表'")
                If foundRows.Length = 0 Then '没有管柱数据表
                    SQL_command = "CREATE TABLE 管柱数据表 " _
                        & "(作业名称 TEXT(50)," & "元件序号 INTEGER , " & "元件名称 TEXT(50)," & "元件性质 TEXT(30)," & "元件长度m FLOAT, " _
                        & "元件外径mm FLOAT, " & "元件内径mm FLOAT, " & "井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 4
                foundRows = dbSchema.Select("TABLE_NAME='测井数据表'")
                If foundRows.Length = 0 Then '没有井眼轨道表
                    SQL_command = "CREATE TABLE 测井数据表 " & "([序   号] integer,[井  深(m)] float,[井斜角(°)] FLOAT,[方位角(°)] FLOAT, [全角变化率(°)/25m] float," & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 5
                '****************************************************************************************************************
                ' 根据基础数据变化情况：
                '2016年8月28日-9月2日数据结构升级情况：
                '   （1）增加以下10个字段：
                '       温度范围下℃     FLOAT
                '       温度范围上℃     FLOAT
                '       抗内压强度MPa    FLOAT
                '       抗外压强度MPa    FLOAT
                '       型号             TEXT(100)
                '       生产厂家         TEXT(100)
                '       上端扣型         TEXT(100)
                '       下端扣型         TEXT(100)
                '       主体材料         TEXT(100)
                '       主材屈服强度MPa  FLOAT
                '       最小坐封压力MPa  FLOAT
                '       最大坐封压力MPa  FLOAT
                '       备注             TEXT(200)
                '              & "剪销承载kN      FLOAT,插管外径mm     FLOAT,插管内径mm    FLOAT,坐封压力MPa   FLOAT,坐封力kN       FLOAT," _
                '
                ' 2016年9月16日，修改生成程序。
                '2017年2月4日
                '    (1) 增加字段 “能否反洗井”，TEXT(10)
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_封隔定位元件'")
                If foundRows.Length = 0 Then '没有管柱_封隔定位元件表
                    SQL_command = "CREATE TABLE 管柱_封隔定位元件 " _
                        & "(作业名称    TEXT(50),元件序号     INTEGER,元件名称   TEXT(50),极限压差MPa   FLOAT,极限载荷kN     FLOAT," _
                        & "剪销承载kN      FLOAT,插管外径mm     FLOAT,插管内径mm    FLOAT,坐封压力MPa   FLOAT,坐封力kN       FLOAT," _
                        & "重量kg          FLOAT,坐封方式    text(50),封隔方式   text(50),定位方式   text(50),井号        TEXT(50)," _
                        & "温度范围下℃    Float,温度范围上℃    Float,抗内压强度MPa Float,抗外压强度MPa Float,型号       TEXT(100)," _
                        & "生产厂家    text(100),上端扣型   TEXT(100),下端扣型  text(100),主体材料  text(100),主材屈服强度MPa Float," _
                        & "最小坐封压力MPa Float,最大坐封压力MPa Float,备注     text(200),能否反洗井 TEXT(10))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 6
                '****************************************************************************************************************
                '2017年10月28日：
                '    由于基础数据库中开关元件数据结构变化，需对用户数据库中 管柱_开关元件表进行升级。
                '   （1）增加以下8个字段：
                '       01       温度范围下℃     FLOAT
                '       02       温度范围上℃     FLOAT
                '       03       压力等级MPa      FLOAT
                '       04       型号             TEXT(100)
                '       05       开关类型         TEXT(100)
                '       06       上端扣型         TEXT(100)
                '       07       下端扣型         TEXT(100)
                '       08       生产厂家         TEXT(100)
                '       09       备注             TEXT(100)
                '   （2）放弃并删除使用2个字段：开关方式，开关流向。注：开关流向在工况中考虑
                '    (3) 关键字：[外径(mm)]，[内径(mm)]，[长度)(m)],[型号]，[开关类型]
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_开关元件'")
                If foundRows.Length = 0 Then '没有管柱_开关元件表
                    SQL_command = "CREATE TABLE 管柱_开关元件 " _
                        & "(作业名称  TEXT(50),元件序号   INTEGER ,元件名称   TEXT(50), 重量kg      FLOAT,抗拉强度kN   FLOAT, " _
                        & " 极限压差MPa  FLOAT,抗外挤强度MPa FLOAT,抗内压强度MPa FLOAT,温度范围下℃ FLOAT,温度范围上℃ FLOAT, " _
                        & " 压力等级MPa  FLOAT,型号      TEXT(100),开关类型  TEXT(100),上端扣型 TEXT(100),下端扣型 TEXT(100), " _
                        & " 生产厂家 TEXT(100),备注      TEXT(100),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 7
                foundRows = dbSchema.Select("TABLE_NAME='工况参数表'")
                If foundRows.Length = 0 Then '没有工况参数表
                    '*************************************************************************************************************************
                    '    2013年2月修订程序，增加了“井底套压MPa、井底管压MPa、压力计算开关”3个字段。
                    '    2013年4月1日再次修订程序，删“压力计算开关”字段，增“井底管压开关”、“井底套压开关”两字段，用以描述井底压力计算方式。
                    '    2015年2月22日再次修订程序，增“井口管压开关”、“井口套压开关”两字段，用以描述井口压力计算方式。
                    '    2015年2月27日再次修订程序，改字段“井底深度m”为“管压井底深度m”，增“套压井底深度m”，用以描述井底压力对应的下深。
                    '    2015年4月1日再次修订程序，增“环空流阻模型”、“管内流阻模型”、“环空稠剂浓度”、“管内稠剂浓度”、
                    '                                “环空撑剂浓度”、“管内撑剂浓度”、“环空流变指数n”、“管内流变指数n”、
                    '                                “环空稠度系数K”、“管内稠度系数K”10个字段用以进行流体摩阻计算。
                    '    2015年7月19日再次修订程序，增“井口加扭Nm”、 “井底约束情况”、“井底约束位置m”、“约束定位方式”4个字段用以解卡、钻磨计算。
                    '                               改“动作参数”为“井口加力kN”
                    '                               删“井口情况”
                    '    2015年7月26日再次修订程序，增“管牛模折减系数”、 “环牛模折减系数”2个字段用以摩阻计算。
                    '    2016年10月24日再次修订程序，增“管流时间h”字段，用以疲劳寿命分析
                    '
                    ' 因修改数据结构而软件运行出错提示及处理方法：
                    '    若在工况参数查看或修改界面出现错误提示："读取工况数据时出错，可能是因工况数据结构升级所致，请手动删除工况数据或修改数
                    ' 据结构！",则表示工况参数表为老的结构，可删除工况参数表，让程序自动重新建立之，但数据需要重新输入。若不想删除工况参数表，
                    ' 则需手动增加成程序最后字段情况并输入数据。
                    '    从2015年7月20日起，所有数据结构升级实现程序自动升级。
                    '*************************************************************************************************************************
                    SQL_command = "CREATE TABLE 工况参数表 " _
                        & "(作业名称     TEXT(50),工况序号      INTEGER,工况名称    TEXT(50),环空流体流向 TEXT(10),管内流体流向 TEXT(10), " _
                        & " 井口温度℃      FLOAT,井底温度℃      FLOAT,管压井底深度m  FLOAT,井口环压MPa     FLOAT,环液深度m       FLOAT, " _
                        & " 环液密度g╱cm3  FLOAT,环液粘度mPaS    FLOAT,环流流量m3╱m  FLOAT,井口管压MPa     FLOAT,管液深度m       FLOAT, " _
                        & " 管液密度g╱cm3  FLOAT,管流粘度mPaS    FLOAT,管流流量m3╱m  FLOAT,库摩系数        FLOAT,井口加力kN      FLOAT, " _
                        & " 井口加扭Nm      FLOAT,井底套压MPa     FLOAT,井底管压MPa    FLOAT,井底套压开关 TEXT(20),井底管压开关 TEXT(20), " _
                        & " 井口套压开关 TEXT(20),井口管压开关 TEXT(20),套压井底深度m  FLOAT,环空流阻模型 TEXT(20),管内流阻模型 TEXT(20), " _
                        & " 环空稠剂浓度    FLOAT,管内稠剂浓度    FLOAT,环空撑剂浓度   FLOAT,管内撑剂浓度    FLOAT,环空流变指数n   FLOAT, " _
                        & " 管内流变指数n   FLOAT,环空稠度系数K   FLOAT,管内稠度系数K  FLOAT,井底约束情况 TEXT(20),井底约束位置m   FLOAT, " _
                        & " 约束定位方式 TEXT(20),管牛模折减系数  FLOAT,环牛模折减系数 FLOAT,管流时间h       FLOAT,井号         TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 8
                foundRows = dbSchema.Select("TABLE_NAME='工况_封隔定位元件'")
                If foundRows.Length = 0 Then '没有工况_封隔定位元件
                    SQL_command = "CREATE TABLE 工况_封隔定位元件 " _
                        & "(作业名称 TEXT(50)," & " 工况序号   INTEGER, " & " 工况名称   TEXT(50), " & " 元件序号 INTEGER , " & " 元件名称 TEXT(50)," _
                        & " 封隔器状态 TEXT(10), " & " 定位方式 text(50), " & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 9
                foundRows = dbSchema.Select("TABLE_NAME='工况_开关元件'")
                If foundRows.Length = 0 Then '没有工况_开关元件
                    SQL_command = "CREATE TABLE 工况_开关元件 " _
                        & "(作业名称 TEXT(50),工况序号   INTEGER   ,工况名称 TEXT(50)   ,元件序号 INTEGER,元件名称 TEXT(50)," _
                        & " 开关状态 TEXT(20),管内嘴损压差MPa FLOAT,油套嘴损压差MPa FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 10
                '*************************************************************************************************************************
                ' 2018年7月30日增加字段
                '   (01)扣型              TEXT(50)
                '   (02)接头抗内压强度MPa FLOAT
                '   (03)接头抗拉强度kN    FLOAT
                '   (04)备注              TEXT(250)
                ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
                '   (05)热膨胀系数    FLOAT
                '*************************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_油管表'")
                If foundRows.Length = 0 Then '没有管柱-油管表
                    SQL_command = "CREATE TABLE 管柱_油管表(" _
                        & "作业名称 TEXT(50)  ,元件序号 INTEGER     ,元件名称 TEXT(50),油管壁厚mm FLOAT       ,油管钢级 TEXT(30)   ," _
                        & "单位长重kg╱m FLOAT,抗拉强度kN FLOAT     ,屈服强度MPa FLOAT,弹性模量MPa FLOAT      ,泊松比 FLOAT        , " _
                        & "抗内压强度MPa FLOAT,抗挤强度MPa FLOAT    ,扣型 TEXT(50)    ,接头抗内压强度MPa FLOAT,接头抗拉强度kN FLOAT, " _
                        & "备注 TEXT(250)     ,热膨胀系数 FLOAT       ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 11
                foundRows = dbSchema.Select("TABLE_NAME='管柱_节流元件'")
                If foundRows.Length = 0 Then '没有管柱-节流元件表
                    SQL_command = "CREATE TABLE 管柱_节流元件(" _
                        & "作业名称 TEXT(50),元件序号 INTEGER , 元件名称 TEXT(50),   重量kg FLOAT,        流孔内径mm FLOAT, " _
                        & "流孔长度mm FLOAT, 抗拉强度kN FLOAT,  抗外挤强度MPa FLOAT, 抗内压强度MPa FLOAT, 节流流向 TEXT(10), " _
                        & "井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 12
                foundRows = dbSchema.Select("TABLE_NAME='节点情况表'")
                If foundRows.Length = 0 Then '没有找到节点情况表
                    '**************************************************************************************************************************
                    '节点类型包括：井口、管柱、套管、井眼、管内变液面、环空变液面（其中井口、管内变液面、环空变液面在节点计算参数表用）
                    '节点性质包括：对井口：
                    '              对管柱有：油管、开关元件、伸缩元件、封隔定位元件、节流元件、管柱扶正元件共6种管柱元件性质
                    '              对套管有：导管、表层套管、技术套管、生产套管、生产尾管、筛管共6种套管类型
                    '              对变液有：
                    '              对井眼有：
                    '节点下深m：   对变液面不填，其余为下入深度
                    '2016年10月24日，为进行疲劳寿命分析，增加：初始损伤长mm，初始损伤宽mm，初始损伤深mm，初始损伤角°，剩余强度系数 五个字段
                    ' 2020年12月20日考虑使用非碳钢油井管，增加字段：热膨胀系数     FLOAT
                    ' 2022年03月06日考虑抗挤、抗拉、抗内压强度校核，增加字段：抗挤强度MPa、管体抗内压MPa、接头抗内压MPa、管体抗拉kN、接头抗拉kN
                    ' 2024年03月15日考虑到液体中管柱单位长度重量qe与环空、管内流体密度有关，即同一节点，不同工况的qe是不同的，由此算得的Fcrh是
                    '               不同的，故屈曲临界载荷Fcrh字段不应放在此表中。
                    '**************************************************************************************************************************
                    SQL_command = "CREATE TABLE 节点情况表(" _
                            & "作业名称  TEXT(50),节点类型    TEXT(50),节点ID    TEXT(50),节点性质  TEXT(30),节点下深m    FLOAT," _
                            & "节点编号   INTEGER,节点垂深m      FLOAT,井斜角rad    FLOAT,方位角rad    FLOAT,套管内径mm   FLOAT," _
                            & "油管外径mm   FLOAT,油管内径mm     FLOAT,线重kg╱m    FLOAT,弹性模量MPa  FLOAT,泊松比       FLOAT," _
                            & "屈服强度MPa   FLOAT,油管钢级   TEXT(30),曲率rad╱m   FLOAT,初始损伤长mm FLOAT,初始损伤宽mm FLOAT," _
                            & "初始损伤深mm  FLOAT,初始损伤角°  FLOAT,剩余强度系数 FLOAT,热膨胀系数   FLOAT,抗挤强度MPa  FLOAT," _
                            & "管体抗内压MPa FLOAT,接头抗内压MPa FLOAT,管体抗拉kN   FLOAT,接头抗拉kN   FLOAT," _
                            & "井号      TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 13
                foundRows = dbSchema.Select("TABLE_NAME='节点计算参数表'")
                If foundRows.Length = 0 Then '没有节点计算参数表
                    '**************************************************************************************************************************
                    '2016年10月24日，为进行疲劳寿命分析，增加：TSM_OK，纵振频率Hz，横振频率Hz，纵振损伤深mm，横振损伤深mm 五个字段
                    '2024年3月15日，增加：Fecrs_N，Fecrh_N，油套摩擦系数 三个字段
                    '2024年3月24日，增加：环液密度g╱cm3， 管液密度g╱cm3 两个字段
                    '2025年8月24日，考虑到同一规格的油管，因轴向应力和内压的不同，其抗挤强度也不同，增加：抗挤强度MPa 字段
                    '**************************************************************************************************************************
                    SQL_command = "CREATE TABLE 节点计算参数表(" _
                            & "作业名称  TEXT(50),工况序号   INTEGER,节点编号     INTEGER,节点下深m      FLOAT," _
                            & "管内压力MPa  FLOAT,管内液压OK TINYINT,管外压力MPa    FLOAT,管外液压OK     TINYINT," _
                            & "真实轴力N    FLOAT,真实轴力OK TINYINT,等效轴力N      FLOAT,等效轴力OK     TINYINT," _
                            & "接触力N      FLOAT,合弯矩Nm     FLOAT,合弯矩OK     TINYINT,节点温度℃     FLOAT ," _
                            & "温度变形m    FLOAT,轴力变形m    FLOAT,鼓胀变形m      FLOAT,螺旋变形m      FLOAT," _
                            & "综合变形m    FLOAT,合成应力MPa  FLOAT,安全系数       FLOAT,等效悬持力N    FLOAT," _
                            & "温度效应m    FLOAT,轴力效应m    FLOAT,鼓胀效应m      FLOAT,螺旋效应m      FLOAT," _
                            & "综合效应m    FLOAT,扭矩Nm       FLOAT,TSM_OK       TINYINT,纵振频率Hz     FLOAT," _
                            & "横振频率Hz   FLOAT,纵振损伤深mm FLOAT,横振损伤深mm   FLOAT,Fecrs_N        FLOAT," _
                            & "Fecrh_N      FLOAT,油套摩擦系数 FLOAT,环液密度g╱cm3 FLOAT,管液密度g╱cm3 FLOAT," _
                            & "抗挤强度MPa  FLOAT,井号      TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 14
                '****************************************************************************************************************
                ' 根据基础数据变化情况：
                '2016年10月10日数据结构升级情况：
                '   （1）增加以下10个字段：
                '       01       温度范围下℃     FLOAT
                '       02       温度范围上℃     FLOAT
                '       03       压力等级MPa      FLOAT
                '       04       抗拉强度kN       FLOAT
                '       05       上端扣型         TEXT(100)
                '       06       下端扣型         TEXT(100)
                '       07       主体材料         TEXT(100)
                '       08       主材屈服强度MPa  FLOAT
                '       09       零件号           TEXT(100)
                '       10       生产厂家         TEXT(100)
                '       11       备注             TEXT(200)
                '       12       全缩短长度m      FLOAT
                '       13       伸缩动作载荷kN   FLOAT
                '2024年9月30日数据结构升级情况：
                '   （1）增加以下1个字段：
                '       14       伸缩动作压力MPa   FLOAT
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_伸缩元件'")
                If foundRows.Length = 0 Then '没有管柱-节流元件表
                    SQL_command = "CREATE TABLE 管柱_伸缩元件(" _
                            & "作业名称    TEXT(50)  ,元件序号 INTEGER  ,元件名称 TEXT(100) ,重量kg FLOAT       ,伸缩行程m FLOAT   , " _
                            & "抗外挤强度MPa FLOAT   ,抗内压强度MPa FLOAT,井号 TEXT(50)     ,温度范围下℃ FLOAT ,温度范围上℃ FLOAT, " _
                            & "压力等级MPa FLOAT     ,抗拉强度kN FLOAT   ,上端扣型 TEXT(100),下端扣型 TEXT(100) ,主体材料 TEXT(100), " _
                            & "主材屈服强度MPa FLOAT ,零件号 TEXT(100)   ,生产厂家 TEXT(100),备注 TEXT(200)     ,全缩短长度m  FLOAT, " _
                            & "伸缩动作载荷kN  FLOAT ,伸缩动作压力MPa FLOAT)"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 15
                foundRows = dbSchema.Select("TABLE_NAME='简便TSM计算结果'")
                If foundRows.Length = 0 Then '没有简便TSM计算结果表
                    SQL_command = "CREATE TABLE 简便TSM计算结果(" _
                            & "作业名称 TEXT(50)  , 工况序号 INTEGER     , 井口轴力kN  FLOAT    , 封隔器处轴力kN  FLOAT ,井口安全系数 FLOAT, " _
                            & "是否螺旋弯曲 TINYINT, 轴力变形m FLOAT      , 温度变形m   FLOAT    , 鼓胀变形m FLOAT       ,螺旋弯曲变形m FLOAT," _
                            & "自由变形m  FLOAT    , 轴力效应m FLOAT      , 温度效应m   FLOAT    , 鼓胀效应m FLOAT       ,螺旋弯曲效应m FLOAT," _
                            & "自由变形效应m FLOAT , 封隔器安全系数 FLOAT , 井口应力MPa FLOAT    , 封隔器处应力MPa  FLOAT,井口内压MPa   FLOAT," _
                            & "井口外压MPa FLOAT   , 封隔器处内压MPa FLOAT, 封隔器处外压MPa FLOAT, 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 16
                foundRows = dbSchema.Select("TABLE_NAME='试油层参数'")
                If foundRows.Length = 0 Then '没有试油层参数表
                    SQL_command = "CREATE TABLE 试油层参数(" _
                            & "层数  INTEGER  ,起始深度 FLOAT     , 结束深度 FLOAT, 压力系数 FLOAT  ,压力 FLOAT, " _
                            & " 温度 FLOAT      ,产出物 TEXT(50)    , 油产量 FLOAT  , 气产量 float    ,水产量 float," _
                            & " 硫化氢含量 FLOAT,二氧化碳含量 FLOAT , 氯根含量 FLOAT, 钻井漏失量 FLOAT," _
                            & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 17
                foundRows = dbSchema.Select("TABLE_NAME='工具参数'")
                If foundRows.Length = 0 Then '没有简便TSM计算结果表
                    SQL_command = "CREATE TABLE 工具参数(" _
                            & "序号  INTEGER   ,工具名称 TEXT(50) ,外径 FLOAT        ,内径 FLOAT       , 材料 TEXT(50)    ," _
                            & " 屈服强度 FLOAT   ,轴向力kN FLOAT    ,内压MPa FLOAT     ,外压MPa FLOAT    ,弯矩Nm FLOAT      ," _
                            & " 剪力N FLOAT      ,扭矩Nm FLOAT      ,轴向应力MPa FLOAT ,径向应力MPa FLOAT,环向应力MPa FLOAT ," _
                            & " 弯矩应力MPa FLOAT,扭矩应力MPa FLOAT ,剪力应力MPa FLOAT ,等效应力MPa FLOAT,安全系数 FLOAT    ," _
                            & " 挤毁临界载荷 FLOAT,等效压力MPa FLOAT,挤毁安全系数 FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 18
                foundRows = dbSchema.Select("TABLE_NAME='试油层参数评估结果'")
                If foundRows.Length = 0 Then '没有简便TSM计算结果表
                    SQL_command = "CREATE TABLE 试油层参数评估结果 " & " (层数  INTEGER,参数名称 TEXT(50),参数值 TEXT(30),风险情况 TEXT(50), 解决方案 TEXT(255))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 19
                '***********************************************************************************************************************
                '     2018年8月7日，修井作业管柱力学分析模块用做摩阻摩矩分析功能实现。放弃原摩阻摩矩分析界面及计算代码。此表不再使用
                '***********************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='摩阻摩矩计算参数表'")
                If foundRows.Length = 0 Then '摩阻摩矩计算参数表
                    SQL_command = "CREATE TABLE 摩阻摩矩计算参数表(" _
                        & "作业名称  TEXT(50)   ,节点编号 integer,    节点下深m FLOAT, " & " 单元全角变化   FLOAT ,悬持等效轴力N  FLOAT,上提等效轴力N  FLOAT," _
                        & "下放等效轴力N  FLOAT ,上提侧向力N    float,下放侧向力N    float," & " 旋转摩擦扭矩Nm FLOAT ,悬持真实轴力N  FLOAT,上提真实轴力N  FLOAT," _
                        & " 下放真实轴力N  FLOAT," & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 20
                foundRows = dbSchema.Select("TABLE_NAME='计算参数表'")
                If foundRows.Length = 0 Then '计算参数表
                    SQL_command = "CREATE TABLE 计算参数表(参数名称  TEXT(50),参数值 TEXT(200),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 21
                foundRows = dbSchema.Select("TABLE_NAME='三次井眼样条函数参数表'")
                If foundRows.Length = 0 Then '没有三次井眼样条函数参数表
                    SQL_command = "CREATE TABLE 三次井眼样条函数参数表 " & "([序   号] integer,[井  深(m)] float,[D_M] FLOAT,[x_m] FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 22
                foundRows = dbSchema.Select("TABLE_NAME='试压参数'")
                If foundRows.Length = 0 Then '没有试压参数表
                    SQL_command = "CREATE TABLE 试压参数( " _
                            & "序号    INTEGER,井段上深m   FLOAT,井段下深m      FLOAT,试压值Mpa   FLOAT,稳压时间min   FLOAT," _
                            & " 稳压值MPa FLOAT,试压介质 TEXT(30),介质密度g╱cm3 FLOAT,水泥塞深度m FLOAT,试压日期   DATETIME," _
                            & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 23
                foundRows = dbSchema.Select("TABLE_NAME='试压套管'")
                If foundRows.Length = 0 Then '没有试压参数表
                    SQL_command = "CREATE TABLE 试压套管 " & "(序号 INTEGER , 套管层数 INTEGER , 套管段数 INTEGER , 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 24
                foundRows = dbSchema.Select("TABLE_NAME='固井质量表'")
                If foundRows.Length = 0 Then '没有固井质量表
                    SQL_command = "CREATE TABLE 固井质量表 " & "(套管层数 INTEGER , 井段上深m FLOAT , 井段下深m FLOAT , 固井质量 TEXT(20) , 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 25
                foundRows = dbSchema.Select("TABLE_NAME='完井设计报告'")
                If foundRows.Length = 0 Then '没有固井质量表
                    SQL_command = "CREATE TABLE 完井设计报告 " & "(一级标题 INTEGER , 二级标题 INTEGER , 三级标题 INTEGER , 标题内序号 INTEGER , 类型 TEXT(20) , 内容 TEXT , 备注 TEXT(200) , 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 26
                '****************************************************************************************************************
                '     为进行管柱疲劳寿命分析，需知道管柱在什么位置，探伤查出有多大的裂纹，据此进行疲劳损伤累积计算。为此，设计
                ' "管柱_油管初始损伤表"它与“管柱_油管表”为1对多关系
                '                                                            --2016年10月23日 秦彦斌
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_油管初始损伤表'")
                If foundRows.Length = 0 Then '没有管柱_油管初始损伤表
                    SQL_command = "CREATE TABLE 管柱_油管初始损伤表(" _
                            & "作业名称 TEXT(50),元件序号 INTEGER,参数序号 INTEGER,损伤位置m FLOAT,损伤长度mm FLOAT," _
                            & "损伤宽度mm FLOAT ,损伤深度mm FLOAT,损伤角度° FLOAT  ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 27
                foundRows = dbSchema.Select("TABLE_NAME='管柱_锚定元件'")
                If foundRows.Length = 0 Then '没有管柱_封隔定位元件表
                    SQL_command = "CREATE TABLE 管柱_锚定元件( " _
                            & "作业名称  TEXT(50),元件序号    INTEGER,元件名称    TEXT(50),型号      TEXT(100),重量kg        FLOAT," _
                            & "工作压力MPa   FLOAT,温度范围下℃  FLOAT,温度范围上℃   FLOAT,坐卡方式   TEXT(50),小坐卡压力MPa FLOAT," _
                            & "大坐卡压力MPa FLOAT,上端扣型  TEXT(100),下端扣型   TEXT(100),生产厂家  TEXT(100),最大锚定力kN  FLOAT," _
                            & "最大解锚力kN  FLOAT,抗内压强度MPa FLOAT,抗外压强度MPa  FLOAT,抗拉强度kN    FLOAT,备注      TEXT(200)," _
                            & "井号       TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 28
                foundRows = dbSchema.Select("TABLE_NAME='工况_锚定元件'")
                If foundRows.Length = 0 Then '没有工况_锚定元件
                    SQL_command = "CREATE TABLE 工况_锚定元件(" _
                            & "作业名称 TEXT(50),工况序号   INTEGER,工况名称   TEXT(50),元件序号 INTEGER ,元件名称 TEXT(50)," _
                            & "定位方式 text(50),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 29
                foundRows = dbSchema.Select("TABLE_NAME='约束点临时表'")
                If foundRows.Length = 0 Then '
                    SQL_command = "CREATE TABLE 约束点临时表 " & "(元件序号 INTEGER,约束点位置m  FLOAT,约束点类型 TEXT(50),约束方式 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 30
                '**************************************************************************************************************************
                '2024年3月15日，增加：Fecrs_N，Fecrh_N，油套摩擦系数 三个字段
                '2024年3月24日，增加：环液密度g╱cm3， 管液密度g╱cm3 两个字段
                '2025年8月24日，考虑到同一规格的油管，因轴向应力和内压的不同，其抗挤强度也不同，增加：抗挤强度MPa 字段
                '**************************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='修井节点计算参数表'")
                If foundRows.Length = 0 Then '没有修井节点计算参数表
                    SQL_command = "CREATE TABLE 修井节点计算参数表(" _
                            & "作业名称    TEXT(50),节点编号  integer,节点下深m    FLOAT,管内压力MPa    FLOAT," _
                            & "管内液压OK   TINYINT,管外压力MPa FLOAT,管外液压OK TINYINT,真实轴力N      FLOAT," _
                            & "真实轴力OK   TINYINT,等效轴力N   FLOAT,等效轴力OK TINYINT,等效悬持力N    FLOAT," _
                            & "真实悬持力N    FLOAT,接触力N     FLOAT,库伦摩擦力N  FLOAT,扭矩Nm         FLOAT," _
                            & "摩擦扭矩Nm     FLOAT,合弯矩Nm    FLOAT,合弯矩OK   TINYINT,节点温度℃     FLOAT," _
                            & "温度变形m      FLOAT,轴力变形m   FLOAT,鼓胀变形m    FLOAT,螺旋变形m      FLOAT," _
                            & "综合变形m      FLOAT,合成应力MPa FLOAT,安全系数     FLOAT,控制安全系数   FLOAT," _
                            & "Fecrs_N        FLOAT,Fecrh_N     FLOAT,油套摩擦系数 FLOAT,环液密度g╱cm3 FLOAT," _
                            & "管液密度g╱cm3 FLOAT,抗挤强度MPa FLOAT,井号     TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                foundRows = dbSchema.Select("TABLE_NAME='修井节点情况表'")
                If foundRows.Length = 0 Then '没有修井节点情况表
                    SQL_command = "select * into 修井节点情况表 from 节点情况表 where 节点下深m<0  and 井号='" & well_name & "'"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 31
                foundRows = dbSchema.Select("TABLE_NAME='管柱组合图表'")
                If foundRows.Length = 0 Then '没有管柱组合图表
                    SQL_command = "CREATE TABLE 管柱组合图表" _
                        & "(序号 INTEGER," _
                        & " 工具编号 TEXT(50), " _
                        & " 工具描述 TEXT(120), " _
                        & " 工具类型 TEXT(30), " _
                        & " 工具型号 TEXT(100), " _
                        & " 工具长度 FLOAT, " _
                        & " 套油管嵌套 TEXT(5), " _
                        & " 图片类型 TEXT(10), " _
                        & " 旋转角度 FLOAT, " _
                        & " 造斜 TEXT(5), " _
                        & " 图元长度na FLOAT, " _
                        & " 图元宽度na FLOAT, " _
                        & " 标注名称 TEXT(30)," _
                        & " 标注深度m FLOAT," _
                        & " 标注层位 TEXT(30)," _
                        & " [层位位置(%)] FLOAT," _
                        & " 层位顶界深m FLOAT," _
                        & " [顶界位置(%)] FLOAT," _
                        & " 层位底界深m FLOAT," _
                        & " [底界位置(%)] FLOAT," _
                        & " 管柱组合表序号 INTEGER, " _
                        & " 井号 TEXT(50), " _
                        & " 作业名称 TEXT(50)," _
                        & " 备注 TEXT(100))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 32
                '****************************************************************************************************************
                '2018年3月28日，考虑修井等作业需要下入钻杆，增加 管柱_普通钻杆 表
                ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
                ' 热膨胀系数     FLOAT
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='管柱_普通钻杆'")
                If foundRows.Length = 0 Then '没有管柱_普通钻杆表
                    SQL_command = "CREATE TABLE 管柱_普通钻杆(" _
                            & "作业名称 TEXT(50)  ,元件序号     INTEGER,规格名称 TEXT(50)     ,钻杆外径mm FLOAT ,钻杆壁厚mm FLOAT    ," _
                            & "钢级 TEXT(50)       ,管体抗拉强度kN FLOAT,屈服强度MPa FLOAT     ,弹性模量MPa FLOAT,泊松比 FLOAT       ," _
                            & "加厚型式  TEXT(50)  ,单根长度m FLOAT     ,单位长度质量kgpm FLOAT,扣型    TEXT(50) ,管体抗扭强度Nm FLOAT," _
                            & "接头抗扭强度Nm FLOAT,接头抗拉强度kN FLOAT ,抗内压强度MPa  FLOAT ,抗挤强度MPa FLOAT ,接头外径mm    FLOAT," _
                            & "备注      TEXT(200) ,热膨胀系数 FLOAT     ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 33
                '****************************************************************************************************************
                '2018年4月10日，为计算卡点位置，建立数据表
                ' (1) 卡点分析节点情况表                      保存管柱节点数据
                '（2）卡点分析计算参数表                      保存两两变形相减数据
                ' (3) 卡点位置计算提拉变形数据                保存输入界面所输入的提拉力及对应的变形量
                ' (4) 卡点分析结果表                          保存卡点分析结果
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='卡点分析计算参数表'")
                If foundRows.Length = 0 Then '没有卡点分析计算参数表
                    SQL_command = "CREATE TABLE 卡点分析计算参数表 " & "(拉力_变形编号 TEXT(50),节点编号 integer,节点下深m FLOAT,dlt_节点变形m FLOAT,作业名称 TEXT(50),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                foundRows = dbSchema.Select("TABLE_NAME='卡点分析结果表'")
                If foundRows.Length = 0 Then '卡点分析结果表
                    SQL_command = "CREATE TABLE 卡点分析结果表(拉力_变形编号 TEXT(50),卡点位置m TEXT(50),dlt_总变形m FLOAT,作业名称 TEXT(50),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                foundRows = dbSchema.Select("TABLE_NAME='卡点分析节点情况表'")
                If foundRows.Length = 0 Then '卡点分析节点情况表
                    SQL_command = "select * into 卡点分析节点情况表 from 节点情况表 where 节点下深m<0  and 井号='" & well_name & "'"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                foundRows = dbSchema.Select("TABLE_NAME='卡点位置计算提拉变形数据'")
                If foundRows.Length = 0 Then '卡点位置计算提拉变形数据
                    SQL_command = "CREATE TABLE 卡点位置计算提拉变形数据(提拉力kN FLOAT,变形量m FLOAT,作业名称 TEXT(50),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 34
                '****************************************************************************************************************
                '2018年8月31日，为记录各工况下，各个伸缩管的拉伸、压缩状态，建立“工况_伸缩管状态”数据表
                '调用方法：Call database_creat(34, 2)  '工况_伸缩管状态
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='工况_伸缩管状态'")
                If foundRows.Length = 0 Then '没有工况_伸缩管状态
                    SQL_command = "CREATE TABLE 工况_伸缩管状态(" _
                            & "作业名称 TEXT(50)    ,工况序号 INTEGER   ,工况名称 TEXT(50)    ,元件序号 INTEGER     ,元件名称 TEXT(50)," _
                            & " 伸缩动作载荷kN FLOAT,伸缩行程m FLOAT    ,全缩短长度m FLOAT    ,剪切力kN FLOAT       ,剪销状态 text(50)," _
                            & " 伸缩长度m FLOAT     ,伸缩状态 text(50),本工况初始长度m FLOAT,伸缩动作压力MPa FLOAT,不伸缩等效轴力kN FLOAT ," _
                            & " 内压MPa FLOAT       ,外压MPa FLOAT      ,压差剪切力kN FLOAT   ,等效轴力kN FLOAT     ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 35
                '****************************************************************************************************************
                '2021年11月11日，为记录管柱重量及拉伸应力安全系数计算结果，建立“重量及拉伸应力安全系数”数据表
                '2023年11月19日，增加空气中累重kN、液体中累重kN、空气中管体剩余拉力kN、液体中管体剩余拉力kN、空气中接头剩余拉力kN、液体中接头剩余拉力kN 六个字段。
                '调用方法：Call database_creat(35, 2)  '重量及拉伸应力安全系数
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='重量及拉伸应力安全系数'")
                If foundRows.Length = 0 Then '没有重量及拉伸应力安全系数
                    SQL_command = "CREATE TABLE 重量及拉伸应力安全系数(" _
                    & "作业名称 TEXT(50)             ,序号 INTEGER                  ,名称规格 TEXT(50)             ,类型 TEXT(30)                 ,下深m TEXT(30)                ," _
                    & "段长m FLOAT                   ,空气中重量kN FLOAT            ,液体中重量kN FLOAT            ,空气中累重kN FLOAT            ,液体中累重kN FLOAT            ," _
                    & "空气管体屈服安全系数 TEXT(20) ,液体管体屈服安全系数 TEXT(20) ,空气管体抗拉安全系数 TEXT(20) ,液体管体抗拉安全系数 TEXT(20) ,空气接头抗拉安全系数 TEXT(20) ," _
                    & "液体接头抗拉安全系数 TEXT(20) ,空气中管体剩余拉力kN FLOAT    ,液体中管体剩余拉力kN FLOAT    ,空气中接头剩余拉力kN FLOAT    ,液体中接头剩余拉力kN FLOAT    ," _
                    & "外径mm FLOAT                  ,内径mm FLOAT                  ,壁厚mm TEXT(20)               ,钢级 TEXT(30)                 ,屈服强度MPa  TEXT(20)         ," _
                    & "管体抗拉强度kN TEXT(20)       ,接头抗拉强度kN TEXT(20)       ,液体密度g╱cm3 FLOAT          ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 36
                '****************************************************************************************************************
                '2022年3月25日，修井及摩阻计算时，开关工具的状态会影响压力及相关计算，为此建立“修井摩阻分析_开关工具状态”数据表
                '调用方法：Call database_creat(36, 2)  '修井摩阻分析_开关工具状态
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='修井摩阻分析_开关工具状态'")
                If foundRows.Length = 0 Then '没有修井摩阻分析_开关工具状态
                    SQL_command = "CREATE TABLE 修井摩阻分析_开关工具状态 " _
                        & "(作业名称 TEXT(50),元件序号 INTEGER,元件名称 TEXT(50),开关状态 TEXT(20),管内嘴损压差MPa FLOAT,油套嘴损压差MPa FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 37
                '****************************************************************************************************************
                '    2022年3月29日，考虑到修井及摩阻计算时，工况参数较多，原来存放在“计算参数表”中存在参数多、不区分“作业”的
                '问题。故建立“修井及摩阻分析工况参数”表，以保存某井某作业管柱修井及摩阻分析时的工况参数。
                '调用方法：Call database_creat(37, 2)  '修井及摩阻分析工况参数
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='修井及摩阻分析工况参数'")
                If foundRows.Length = 0 Then '没有修井及摩阻分析工况参数
                    SQL_command = "CREATE TABLE 修井及摩阻分析工况参数 " _
                        & "(作业名称     TEXT(50),管柱下入长度m   FLOAT,环空流体流向 TEXT(10),管内流体流向  TEXT(10), " _
                        & " 井口温度℃      FLOAT,井底温度℃      FLOAT,管压井底深度m   FLOAT,井口环压MPa    FLOAT,环液深度m       FLOAT, " _
                        & " 环液密度g╱cm3  FLOAT,环液粘度mPaS    FLOAT,环流流量m3╱m   FLOAT,井口管压MPa    FLOAT,管液深度m       FLOAT, " _
                        & " 管液密度g╱cm3  FLOAT,管流粘度mPaS    FLOAT,管流流量m3╱m   FLOAT,库摩系数       FLOAT,井口加力kN      FLOAT, " _
                        & " 井口加扭Nm      FLOAT,井底套压MPa     FLOAT,井底管压MPa     FLOAT,井底套压开关 TEXT(20),井底管压开关 TEXT(20), " _
                        & " 井口套压开关 TEXT(20),井口管压开关 TEXT(20),套压井底深度m   FLOAT,环空流阻模型 TEXT(20),管内流阻模型 TEXT(20), " _
                        & " 环空稠剂浓度    FLOAT,管内稠剂浓度    FLOAT,环空撑剂浓度    FLOAT,管内撑剂浓度    FLOAT,环空流变指数n   FLOAT, " _
                        & " 管内流变指数n   FLOAT,环空稠度系数K   FLOAT,管内稠度系数K   FLOAT,轴力加载方式 TEXT(20),扭矩加载方式 TEXT(20), " _
                        & " 管牛模折减系数  FLOAT,环牛模折减系数  FLOAT,管流时间h       FLOAT,井号         TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 38
                '****************************************************************************************************************
                '    2022年10月19日，考虑到管柱下入性分析时，（1）工况参数较多，原来存放在“计算参数表”中存在参数多、不区分“作业”
                '的问题；（2）下入性分析结果数据存放问题。建立“下入性分析工况参数”表，以保存某井某作业管柱下入性分析时的工况
                '参数及主要计算结果。
                '调用方法：Call database_creat(38, 2)  '下入性分析工况参数
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='下入性分析工况参数'")
                If foundRows.Length = 0 Then '没有下入性分析工况参数
                    SQL_command = "CREATE TABLE 下入性分析工况参数 " _
                        & "(作业名称     TEXT(50),管柱下入长度m   FLOAT,环空流体流向 TEXT(10),管内流体流向  TEXT(10), " _
                        & " 井口温度℃      FLOAT,井底温度℃      FLOAT,管压井底深度m   FLOAT,井口环压MPa    FLOAT,环液深度m       FLOAT, " _
                        & " 环液密度g╱cm3  FLOAT,环液粘度mPaS    FLOAT,环流流量m3╱m   FLOAT,井口管压MPa    FLOAT,管液深度m       FLOAT, " _
                        & " 管液密度g╱cm3  FLOAT,管流粘度mPaS    FLOAT,管流流量m3╱m   FLOAT,库摩系数       FLOAT,井口加力kN      FLOAT, " _
                        & " 井口加扭Nm      FLOAT,井底套压MPa     FLOAT,井底管压MPa     FLOAT,井底套压开关 TEXT(20),井底管压开关 TEXT(20), " _
                        & " 井口套压开关 TEXT(20),井口管压开关 TEXT(20),套压井底深度m   FLOAT,环空流阻模型 TEXT(20),管内流阻模型 TEXT(20), " _
                        & " 环空稠剂浓度    FLOAT,管内稠剂浓度    FLOAT,环空撑剂浓度    FLOAT,管内撑剂浓度    FLOAT,环空流变指数n   FLOAT, " _
                        & " 管内流变指数n   FLOAT,环空稠度系数K   FLOAT,管内稠度系数K   FLOAT,轴力加载方式 TEXT(20),扭矩加载方式 TEXT(20), " _
                        & " 管牛模折减系数  FLOAT,环牛模折减系数  FLOAT,管流时间h       FLOAT,井口下压力否 TEXT(20),可下入深度m     FLOAT, " _
                        & " 循环迭代次数 TEXT(20),井号         TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 39
                '****************************************************************************************************************
                '2022年10月20日，下入性分析时，开关工具的状态会影响压力及相关计算，为此建立“下入性分析_开关工具状态”数据表
                '调用方法：Call database_creat(39, 2)  '下入性分析_开关工具状态
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='下入性分析_开关工具状态'")
                If foundRows.Length = 0 Then '没有修井摩阻分析_开关工具状态
                    SQL_command = "CREATE TABLE 下入性分析_开关工具状态 " _
                        & "(作业名称 TEXT(50),元件序号 INTEGER,元件名称 TEXT(50),开关状态 TEXT(20),管内嘴损压差MPa FLOAT,油套嘴损压差MPa FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 40
                '**************************************************************************************************************************
                '2024年10月20日，为了数据独立,建立“下入性分析_开关工具状态”数据表
                '调用方法：Call database_creat(40, 2)  '建立下入性节点情况表，下入性节点计算参数表
                '2024年3月15日，增加：Fecrs_N，Fecrh_N，油套摩擦系数 三个字段
                '2024年3月24日，增加：环液密度g╱cm3， 管液密度g╱cm3 两个字段
                '2025年8月24日，考虑到同一规格的油管，因轴向应力和内压的不同，其抗挤强度也不同，增加：抗挤强度MPa 字段
                '**************************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='下入性节点计算参数表'")
                If foundRows.Length = 0 Then '没有下入性节点计算参数表
                    SQL_command = "CREATE TABLE 下入性节点计算参数表(" _
                            & "作业名称    TEXT(50),节点编号  integer,节点下深m    FLOAT,管内压力MPa    FLOAT," _
                            & "管内液压OK   TINYINT,管外压力MPa FLOAT,管外液压OK TINYINT,真实轴力N      FLOAT," _
                            & "真实轴力OK   TINYINT,等效轴力N   FLOAT,等效轴力OK TINYINT,等效悬持力N    FLOAT," _
                            & "真实悬持力N    FLOAT,接触力N     FLOAT,库伦摩擦力N  FLOAT,扭矩Nm         FLOAT," _
                            & "摩擦扭矩Nm     FLOAT,合弯矩Nm    FLOAT,合弯矩OK   TINYINT,节点温度℃     FLOAT," _
                            & "温度变形m      FLOAT,轴力变形m   FLOAT,鼓胀变形m    FLOAT,螺旋变形m      FLOAT," _
                            & "综合变形m      FLOAT,合成应力MPa FLOAT,安全系数     FLOAT,控制安全系数   FLOAT," _
                            & "Fecrs_N        FLOAT,Fecrh_N     FLOAT,油套摩擦系数 FLOAT,环液密度g╱cm3 FLOAT," _
                            & "管液密度g╱cm3 FLOAT,抗挤强度MPa FLOAT,井号      TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                foundRows = dbSchema.Select("TABLE_NAME='下入性节点情况表'")
                If foundRows.Length = 0 Then '没有下入性节点情况表
                    SQL_command = "select * into 下入性节点情况表 from 节点情况表 where 节点下深m<0  and 井号='" & well_name & "'"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 41
                '**************************************************************************************************************************
                '2024年7月11日，为解决坐封且井下关井工况下封隔器以下环空及关井阀以下管内的压力计算有误问题，建立“作业地层参数表”数据表
                '调用方法：Call database_creat(41, 2)  
                '作业地层参数表，李润洲2025年7月5日增加字段：目的层名称,
                '                李润洲2025年10月23日增加字段：人工井底m
                '**************************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='作业地层参数表'")
                If foundRows.Length = 0 Then '没有作业地层参数表
                    SQL_command = "CREATE TABLE 作业地层参数表 " _
                        & "(作业名称 TEXT(50),压力系数 FLOAT,地破压力MPa FLOAT,地破垂深m FLOAT,目的层温度℃ FLOAT," _
                        & " 目的层垂深m FLOAT,人工井底m FLOAT,目的层名称  TEXT(50),井号  TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If

            Case 62
                foundRows = dbSchema.Select("TABLE_NAME='钻具组合_钻铤表'")
                If foundRows.Length = 0 Then '没有钻具组合_钻铤表
                    SQL_command = "CREATE TABLE 钻具组合_钻铤表(" _
                            & "组合序号 INTEGER,元件序号 INTEGER    ,钻铤规格 TEXT(50),钻铤外径mm FLOAT      ,钻铤内径mm FLOAT," _
                            & "螺纹类型 TEXT(30),最小抗拉强度kN FLOAT,屈服强度MPa FLOAT,弹性模量MPa FLOAT     ,泊松比 FLOAT    ," _
                            & "井深m FLOAT      ,元件个数 integer    ,单根长度m FLOAT  ,单位长度质量kgpm FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 63
                foundRows = dbSchema.Select("TABLE_NAME='钻具组合_加重钻杆表'")
                If foundRows.Length = 0 Then '没有钻具组合_加重钻杆表
                    SQL_command = "CREATE TABLE 钻具组合_加重钻杆表 " _
                            & "(组合序号 INTEGER     ,元件序号 INTEGER  ,井深m FLOAT      ,元件个数 integer    ,加重钻杆规格 TEXT(50)," _
                            & "接头外径mm FLOAT      ,接头内径mm FLOAT  ,螺纹类型 TEXT(30),最小抗拉强度kN FLOAT,最小屈服强度MPa FLOAT," _
                            & "弹性模量MPa FLOAT     ,泊松比 FLOAT      ,管体外径mm FLOAT ,管体内径mm FLOAT    ,单根长度m FLOAT      ," _
                            & "单位长度质量kgpm FLOAT,加厚形式  TEXT(50),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 64
                '****************************************************************************************************************
                ' 根据2018年3月18日钻杆基础数据变化情况：
                '2018年3月22日数据结构升级情况：
                '增加字段
                '   (01)扣型              TEXT(50)
                '   (02)管体抗扭强度Nm    FLOAT
                '   (03)接头抗扭强度Nm    FLOAT
                '   (04)管体抗拉强度kN    FLOAT
                '   (05)接头抗拉强度kN    FLOAT
                '   (06)抗内压强度MPa     FLOAT
                '   (07)抗挤强度MPa       FLOAT
                '   (08)接头外径mm        FLOAT
                '   (09)屈服强度MPa       FLOAT
                '   (10)备注              TEXT(200)
                '   (11)将“最小抗拉强度kN”值赋给“管体抗拉强度kN”后删除
                '   (12)将“最小屈服强度MPa”值赋给“屈服强度MPa”后删除
                ' 2020年12月20日考虑使用非碳钢油井管，增加字段：
                '   (13)热膨胀系数     FLOAT
                '****************************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='钻具组合_普通钻杆表'")
                If foundRows.Length = 0 Then '没有钻具组合_普通钻杆表
                    SQL_command = "CREATE TABLE 钻具组合_普通钻杆表 " _
                            & "(组合序号 INTEGER     ,元件序号 INTEGER    ,井深m FLOAT         ,元件个数 integer    ,钻杆规格 TEXT(50)     ," _
                            & "钻杆外径mm FLOAT      ,钻杆壁厚mm FLOAT    ,钢级 TEXT(50)       ,管体抗拉强度kN FLOAT,屈服强度MPa FLOAT ," _
                            & "弹性模量MPa FLOAT     ,泊松比 FLOAT        ,加厚型式  TEXT(50)  ,单根长度m FLOAT     ,单位长度质量kgpm FLOAT," _
                            & "扣型          TEXT(50),管体抗扭强度Nm FLOAT,接头抗扭强度Nm FLOAT,接头抗拉强度kN FLOAT,抗内压强度MPa  FLOAT," _
                            & "抗挤强度MPa      FLOAT,接头外径mm     FLOAT,备注      TEXT(200) ,热膨胀系数  FLOAT     ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 65
                '20210528 加“已固井深”字段
                foundRows = dbSchema.Select("TABLE_NAME='钻井日志表'")
                If foundRows.Length = 0 Then '没有钻井记录表
                    SQL_command = "CREATE TABLE 钻井日志表(" _
                            & "序号 integer  ,起深 FLOAT,进尺 FLOAT, 纯钻时间 FLOAT , 转速 FLOAT , " _
                            & "钻压 FLOAT    ,泥浆类型 TEXT(30) , 泥浆密度 FLOAT , 划扩眼时间 FLOAT , 泵压 FLOAT , " _
                            & "已固井深 FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 67
                '20210530 加“已固井深”字段
                foundRows = dbSchema.Select("TABLE_NAME='钻具组合表'")
                If foundRows.Length = 0 Then '钻具组合表
                    SQL_command = "CREATE TABLE 钻具组合表 " & "(组合序号 INTEGER ,元件序号 INTEGER , 井深 FLOAT ,用途 TEXT(50),起下次数  INTEGER,个数  INTEGER,已固井深 FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 68
                foundRows = dbSchema.Select("TABLE_NAME='磨损分析结果数据表'")
                If foundRows.Length = 0 Then '套管磨损计算节点参数表
                    SQL_command = "CREATE TABLE 磨损分析结果数据表 " _
                            & "(序号 INTEGER    ,井深m FLOAT        ,层数 INTEGER     ,段数 INTEGER       ,是否回接   TEXT(50)," _
                            & " 套管外径mm FLOAT,初始壁厚mm    FLOAT,悬挂深度m   FLOAT,套管下深m     FLOAT,完钻深度m     FLOAT," _
                            & " 抗内压MPa  FLOAT,抗挤强度MPa   FLOAT,抗拉强度kN FLOAT ,屈服极限MPa   FLOAT,狗腿度        FLOAT," _
                            & " 磨损深度mm FLOAT,磨损面积m2    FLOAT,剩余抗内压MPa FLOAT,剩余抗挤MPa FLOAT,剩余抗拉kN    FLOAT," _
                            & " 钻杆外径mm FLOAT,钻套磨损次数  FLOAT,套管钢级   TEXT(50),最大侧向力  FLOAT,钻套磨损时间  FLOAT," _
                            & " 井号    TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 69
                foundRows = dbSchema.Select("TABLE_NAME='磨损节点临时表'")
                If foundRows.Length = 0 Then '磨损节点临时表
                    SQL_command = "CREATE TABLE 磨损节点临时表(序号 INTEGER,井深m FLOAT,层数 INTEGER,段数 INTEGER,套管外径mm FLOAT," _
                    & " 初始壁厚mm FLOAT,磨损深度mm FLOAT,狗腿度 FLOAT,磨损面积m2 FLOAT,钻杆外径mm  FLOAT,钻套磨损次数  FLOAT," _
                    & " 套管钢级 TEXT(50),屈服极限MPa FLOAT,最大侧向力  FLOAT,钻套磨损时间  FLOAT)"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 70
                foundRows = dbSchema.Select("TABLE_NAME='套管磨损分析参数表'")
                If foundRows.Length = 0 Then '套管磨损分析参数表
                    SQL_command = "CREATE TABLE 套管磨损分析参数表(层 INTEGER,段 INTEGER,摩系系数 FLOAT,磨效系数 FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 71
                foundRows = dbSchema.Select("TABLE_NAME='磨损套管作业参数分析数据表'")
                If foundRows.Length = 0 Then '磨损套管作业参数分析数据表
                    SQL_command = "CREATE TABLE 磨损套管作业参数分析数据表 " _
                                & "(序号 INTEGER       ,井深m FLOAT         ,层数 INTEGER ,段数 INTEGER  ,是否回接   TEXT(50)," _
                                & " 剩余抗内压MPa FLOAT,剩余抗挤MPa FLOAT   ,垂深m FLOAT  ,泥浆密度 FLOAT,最低替液密度 FLOAT ," _
                                & " 最高环空压力  FLOAT,最高环空压力2  FLOAT,最大掏空深度 FLOAT,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 72
                '**********************************************************************************************************
                '  2017年5月5-7日，为评价老井套管磨损情况，设“完井后起下钻统计表”
                '**********************************************************************************************************
                foundRows = dbSchema.Select("TABLE_NAME='完井后起下钻统计表'")
                If foundRows.Length = 0 Then '完井后起下钻统计表
                    SQL_command = "CREATE TABLE 完井后起下钻统计表 " _
                            & "(序号      INTEGER,下入日期  DATETIME,下入深度m FLOAT      ,下入井液密度 FLOAT,下入井液类型 TEXT(30)," _
                            & " 起出日期 DATETIME,起出井液密度 FLOAT,起出井液类型 TEXT(30),工作内容 TEXT(254),井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 91
                foundRows = dbSchema.Select("TABLE_NAME='冲蚀预测_压裂参数表'")
                If foundRows.Length = 0 Then '冲蚀预测-压裂参数表
                    SQL_command = "CREATE TABLE 冲蚀预测_压裂参数表" _
                            & "(序号 INTEGER      ,程序名称  TEXT(50) ,改造起深m FLOAT     ,改造止深m FLOAT    ,施工序号 INTEGER    ," _
                            & " 步骤名称 TEXT(50) ,液体名称 TEXT(50)  ,套注液量_方 FLOAT   ,套注砂量_方 FLOAT  ,套注流量_方pm FLOAT ," _
                            & " 管注液量_方 FLOAT ,管注砂量_方 FLOAT  ,管注流量_方pm FLOAT ,注入方式 TEXT(50)  ,作业名称 TEXT(50)   ," _
                            & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 92
                foundRows = dbSchema.Select("TABLE_NAME='冲蚀预测分析结果表'")
                If foundRows.Length = 0 Then '冲蚀预测-压裂参数表
                    SQL_command = "CREATE TABLE 冲蚀预测分析结果表" _
                            & "(序号 INTEGER         ,冲蚀段起深m  FLOAT  ,冲蚀段止深m  FLOAT    ,套管层数 INTEGER    ,套管段数 INTEGER     ," _
                            & " 套管外径mm FLOAT     ,套管初始壁厚mm FLOAT,套管钢级 TEXT(50)     ,套管冲蚀深度mm FLOAT,管柱作业名称 TEXT(50)," _
                            & " 管柱元件序号 INTEGER ,管柱元件名称 TEXT(50),管柱元件性质 TEXT(30),管柱元件外径mm FLOAT,管柱元件内径mm FLOAT ," _
                            & " 管柱元件钢级 TEXT(30),管柱元件外硬度 FLOAT ,管柱元件内硬度 FLOAT ,管外冲蚀深度mm FLOAT,管内冲蚀深度mm FLOAT ," _
                            & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
            Case 501
                foundRows = dbSchema.Select("TABLE_NAME='套管经历流体流动参数表'")
                If foundRows.Length = 0 Then '套管经历流体流动参数表
                    SQL_command = "CREATE TABLE 套管经历流体流动参数表" _
                            & "(序号 INTEGER         ,产前温度℃  FLOAT  ,产前泵压MPa  FLOAT    ,产前密度kgpcm3 FLOAT    ,产前流量m3pmin FLOAT     ," _
                            & " 产前粘度mPas FLOAT   ,产后温度℃  FLOAT  ,产后泵压MPa  FLOAT    ,产后密度kgpcm3 FLOAT    ,产后流量m3pmin FLOAT     ," _
                            & " 产后粘度mPas FLOAT   ,工作时间day FLOAT  ,流体流向 TEXT(20)     ,井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2016年9月3日 基础数据表新增“封隔器信封曲线参数表”，用以保存信封曲线参数
                'Call database_creat(10001, 1)   '建立数据表  封隔器信封曲线参数表
                '*********************************************************************************************************************************
            Case 10001
                foundRows = dbSchema.Select("TABLE_NAME='封隔器信封曲线参数表'")
                If foundRows.Length = 0 Then '封隔器-信封曲线参数表
                    SQL_command = "CREATE TABLE 封隔器信封曲线参数表" & "(序号  INTEGER ,轴力kN  FLOAT,压差MPa  FLOAT,型号  TEXT(100),最大外径mm FLOAT,最小通径mm FLOAT,长度m FLOAT,坐封方式 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2017年2月12日 基础数据表新增“锚定工具表”，用以保存锚定工具参数
                'Call database_creat(10002, 1)   '建立数据表  锚定工具表
                '*********************************************************************************************************************************
            Case 10002
                foundRows = dbSchema.Select("TABLE_NAME='锚定工具表'")
                If foundRows.Length = 0 Then '没有锚定工具表
                    SQL_command = "CREATE TABLE 锚定工具表 " _
                            & "(型号     TEXT(100),工具名称   TEXT(50),最大外径mm     FLOAT,最小通径mm    FLOAT,总体长度m     FLOAT," _
                            & "重量kg        FLOAT,工作压力MPa   FLOAT,温度范围下℃  FLOAT,温度范围上℃   FLOAT,坐卡方式   TEXT(50)," _
                            & "小坐卡压力MPa FLOAT,大坐卡压力MPa FLOAT,上端扣型  TEXT(100),下端扣型   TEXT(100),生产厂家  TEXT(100)," _
                            & "最大锚定力kN  FLOAT,最大解锚力kN  FLOAT,抗内压强度MPa FLOAT,抗外压强度MPa  FLOAT,抗拉强度kN    FLOAT," _
                            & "备注      TEXT(200))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2017年10月15日 基础数据表新增“开关元件载荷性能图参数表”，用以保存开关元件载荷性能图参数
                'Call database_creat(10003, 1)   '建立数据表  开关元件载荷性能图参数表
                '*********************************************************************************************************************************
            Case 10003
                foundRows = dbSchema.Select("TABLE_NAME='开关元件载荷性能图参数表'")
                If foundRows.Length = 0 Then '开关元件载荷性能图参数表
                    SQL_command = "CREATE TABLE 开关元件载荷性能图参数表" _
                            & "(序号  INTEGER ,轴力kN  FLOAT,压差MPa  FLOAT,型号  TEXT(100),最大外径mm FLOAT,最小通径mm FLOAT,长度m FLOAT,开关类型 TEXT(100))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2017年10月24日 基础数据表新增“锚定工具载荷性能图参数表”，用以保存锚定工具载荷性能图参数
                'Call database_creat(10004, 1)   '建立数据表  锚定工具载荷性能图参数表
                '*********************************************************************************************************************************
            Case 10004
                foundRows = dbSchema.Select("TABLE_NAME='锚定工具载荷性能图参数表'")
                If foundRows.Length = 0 Then '锚定工具载荷性能图参数表
                    SQL_command = "CREATE TABLE 锚定工具载荷性能图参数表" _
                            & "(序号  INTEGER ,轴力kN  FLOAT,压差MPa  FLOAT,型号  TEXT(100),最大外径mm FLOAT,最小通径mm FLOAT,长度m FLOAT,坐卡方式 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2025年7月20日 基础数据表新增“枪弹数据表”，用以保存射孔枪弹参数    李润洲
                'Call database_creat(10010, 1)   '建立数据表  枪弹数据表
                '*********************************************************************************************************************************
            Case 10010
                foundRows = dbSchema.Select("TABLE_NAME='枪弹数据表'")
                If foundRows.Length = 0 Then '枪弹数据表
                    SQL_command = "CREATE TABLE 枪弹数据表" _
                            & "(射孔器分类 TEXT(50)  ,[射孔器名称] TEXT(50),[外径mm] FLOAT       ,[壁厚mm] FLOAT     ,[耐压MPa] FLOAT," _
                            & "[射孔密度孔╱m] FLOAT ,[射孔相位°] FLOAT   ,[射孔弹名称] TEXT(50),[炸药名称] TEXT(50),[装药量g] FLOAT," _
                            & "[装药密度g╱cm3] FLOAT,[耐温°C] FLOAT      ,[套管外径mm] FLOAT   ,[平均孔径mm] FLOAT ,[平均穿深mm] FLOAT," _
                            & "[流动效率] FLOAT      ,[钢靶孔径mm] FLOAT   ,[钢靶穿深mm] FLOAT   ,[备注] TEXT(250))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2025年7月20日 基础数据表新增“炸药表”，用以保存射孔炸药参数    李润洲
                'Call database_creat(10011, 1)   '建立数据表  锚定工具载荷性能图参数表
                '*********************************************************************************************************************************
            Case 10011
                foundRows = dbSchema.Select("TABLE_NAME='炸药表'")
                If foundRows.Length = 0 Then '炸药表
                    SQL_command = "CREATE TABLE 炸药表" _
                            & "(炸药名 TEXT(50),分子式 TEXT(50)    ,C数   FLOAT          ,H数 FLOAT         ,N数  FLOAT," _
                            & "O数 FLOAT       ,相对分子质量 FLOAT ,[生成热kJ╱mol] FLOAT,[爆热KJ_mol] FLOAT," _
                            & "[爆温_K] FLOAT  ,[备注]   TEXT(250))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2025年7月20日 基础数据表新增“岩石数据表”，用以保存目的层岩石参数    李润洲
                'Call database_creat(10012, 1)   '建立数据表  岩石数据表
                '*********************************************************************************************************************************
            Case 10012
                foundRows = dbSchema.Select("TABLE_NAME='岩石数据表'")
                If foundRows.Length = 0 Then '岩石数据表
                    SQL_command = "CREATE TABLE 岩石数据表" _
                            & "(岩石岩性 TEXT(50)   ,岩石强度Mpa   FLOAT  ,弹性模量Mpa FLOAT    ,变形系数  FLOAT," _
                            & "压力敏感系数 FLOAT   ,泊松比        FLOAT  ,[备注]   TEXT(255))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2025年7月24日 基础数据表新增“筛管”，用以保存筛管工具的参数    李润洲
                'Call database_creat(10013, 1)   '建立数据表  筛管
                '*********************************************************************************************************************************
            Case 10013
                foundRows = dbSchema.Select("TABLE_NAME='筛管'")
                If foundRows.Length = 0 Then '筛管
                    SQL_command = "CREATE TABLE 筛管" _
                            & "(规格名称 TEXT(50) ,外径mm FLOAT          ,壁厚mm FLOAT          ,单根长度m    FLOAT       ,单根质量kg FLOAT          ," _
                            & "主体材料 TEXT(50)  ,接头抗拉伸强度kN FLOAT,主材屈服强度MPa FLOAT ,抗内压强度MPa  FLOAT     ,抗外挤强度MPa FLOAT       ," _
                            & "耐温℃ FLOAT       ,耐压差MPa FLOAT       ,[孔隙率%]  FLOAT      ,有效流通面积cm2╱m FLOAT ,拦截的最小颗粒直径mm FLOAT," _
                            & "生产厂家  TEXT(50) ,备注    TEXT(200))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '2025年8月2日 基础数据表增加“射孔施工方式表”    李润洲
                'Call database_creat(10014, 1)   '建立数据表  射孔施工方式表
                '*********************************************************************************************************************************
            Case 10014
                foundRows = dbSchema.Select("TABLE_NAME='射孔施工方式表'")
                If foundRows.Length = 0 Then '射孔施工方式表
                    SQL_command = "CREATE TABLE 射孔施工方式表 (施工方式 TEXT(100))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2017年2月14日增加
                'Call database_creat(20001, 2)  '建立数据表，井眼轨迹插值数据表
                '*********************************************************************************************************************************
            Case 20001
                foundRows = dbSchema.Select("TABLE_NAME='井眼轨迹插值参数表'")
                If foundRows.Length = 0 Then '没有井眼轨迹插值参数表
                    SQL_command = "CREATE TABLE 井眼轨迹插值参数表 " _
                        & " (序号 INTEGER, " & " [井  深(m)] FLOAT, " & " [井斜角(°)] FLOAT, " & " [方位角(°)] FLOAT, " & " [铅垂深度（m）] FLOAT, " & " [水平面内北向长度（m）] FLOAT, " _
                        & " [水平面内东向长度（m）] FLOAT, " & " [比例因子F] FLOAT," & " 井号 TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年6月15日增加
                'Call database_creat(20010, 2)  '建立数据表，目的层参数表
                '李润洲2025年10月23日，删除"人工井底m"字段
                '*********************************************************************************************************************************
            Case 20010
                foundRows = dbSchema.Select("TABLE_NAME='目的层参数表'")
                'SQL_command = "CREATE TABLE 目的层参数表 " _
                '        & " (目的层名称  TEXT(50) ,起始深度m     FLOAT ,终止深度m    FLOAT ," _
                '        & "地层温度℃       FLOAT ,地层压力MPa   FLOAT ,人工井底m    FLOAT ,目的层中部垂深m  FLOAT,目的层中部井深m  FLOAT    ,地层压力系数 FLOAT,地破压力MPa FLOAT," _
                '        & "岩石岩性      TEXT(30) ,岩石强度MPa   FLOAT ,弹性模量MPa  FLOAT ,变形系数         FLOAT,[孔隙度%]        FLOAT    ,渗透率um2    FLOAT,泊松比      FLOAT,压力敏感系数   FLOAT," _
                '        & "地层系        TEXT(50) ,地层组     TEXT(50) ,地层段    TEXT(50) ,最大地应力MPa    FLOAT,最大地应力方向° FLOAT    ,井号   TEXT(50)   ,备注  TEXT(255))"
                If foundRows.Length = 0 Then '没有射孔段参数表
                    SQL_command = "CREATE TABLE 目的层参数表 " _
                        & " (目的层名称  TEXT(50) ,起始深度m     FLOAT ,终止深度m    FLOAT    ," _
                        & "地层温度℃       FLOAT ,地层压力MPa   FLOAT ,目的层中部垂深m  FLOAT,目的层中部井深m  FLOAT    ,地层压力系数 FLOAT,地破压力MPa FLOAT," _
                        & "岩石岩性      TEXT(30) ,岩石强度MPa   FLOAT ,弹性模量MPa  FLOAT    ,变形系数         FLOAT,[孔隙度%]        FLOAT    ,渗透率um2    FLOAT,泊松比      FLOAT,压力敏感系数   FLOAT," _
                        & "地层系        TEXT(50) ,地层组     TEXT(50) ,地层段    TEXT(50)    ,最大地应力MPa    FLOAT,最大地应力方向° FLOAT    ,井号   TEXT(50)   ,备注  TEXT(255))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年6月15日增加
                'Call database_creat(20011, 2)  '建立数据表，射孔工况参数表
                '计算总装药量方法：(1)用装弹数*每弹药量；(2)夹层厚度*孔密
                '*********************************************************************************************************************************
            Case 20011
                foundRows = dbSchema.Select("TABLE_NAME='射孔工况参数表'")
                If foundRows.Length = 0 Then '没有射孔工况参数表
                    
                    SQL_command = "CREATE TABLE 射孔工况参数表 " _
                        & " (工况序号  INTEGER     ,工况名称   TEXT(50)     ,井口温度℃ FLOAT          ,射孔段起始深度m FLOAT  ,射孔夹层总厚度m FLOAT  ," _
                        & "射孔段处井底温度℃ FLOAT,井底初始压力MPa  FLOAT  ,井口加压MPa FLOAT         ,射孔液密度g╱cm3 FLOAT ,[射孔液粘度mPa·s]  FLOAT ,射孔施工方式  TEXT(100)," _
                        & "射孔枪下井时间h  FLOAT  ,装弹数  INTEGER         ,计算总装药量方法 TEXT(10) ,动载系数  FLOAT        ,筛管距射孔顶端距离m FLOAT ," _
                        & "人工井底m   FLOAT       ,作业名称  TEXT(50)      ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年6月18日增加
                'Call database_creat(20012, 2)  '建立数据表，管柱-射孔枪弹表
                '*********************************************************************************************************************************
            Case 20012
                foundRows = dbSchema.Select("TABLE_NAME='管柱_射孔枪弹表'")
                If foundRows.Length = 0 Then '没有管柱-射孔枪弹表
                    SQL_command = "CREATE TABLE 管柱_射孔枪弹表 " _
                        & " (作业名称 TEXT(50)       ,元件序号  INTEGER     ,[元件名称] TEXT(50)  ,[射孔器分类] TEXT(50) ,[壁厚mm]  FLOAT   ,[耐压MPa]  FLOAT  ," _
                        & "[射孔密度孔╱m]  FLOAT    ,[射孔相位°] FLOAT    ,[射孔弹名称] TEXT(50),[炸药名称] TEXT(50)   ,[装药量g] FLOAT   ," _
                        & "[装药密度g╱cm3] FLOAT    ,[耐温°C]  FLOAT      ,[套管外径mm] FLOAT   ,[平均孔径mm] FLOAT    ,[平均穿深mm] FLOAT," _
                        & "[流动效率] FLOAT          ,[钢靶孔径mm] FLOAT    ,[钢靶穿深mm] FLOAT   ,[备注] TEXT(250)      ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年7月5日增加
                'Call database_creat(20013, 2)  '建立数据表，工况_射孔夹层表
                '*********************************************************************************************************************************
            Case 20013
                foundRows = dbSchema.Select("TABLE_NAME='工况_射孔夹层表'")
                If foundRows.Length = 0 Then '没有管柱-射孔枪弹表
                    SQL_command = "CREATE TABLE 工况_射孔夹层表 " _
                        & " (射孔枪序号 INTEGER ,起始深度m FLOAT    ,终止深度m FLOAT    ,工况序号  INTEGER ,射孔枪名称 TEXT(50)," _
                        & "射孔枪状态  TEXT(10) ,工况名称 TEXT(50)  ,作业名称 TEXT(50)  ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年7月14日增加
                'Call database_creat(20014, 2)  '建立数据表，射孔段爆轰计算参数
                '*********************************************************************************************************************************
            Case 20014
                foundRows = dbSchema.Select("TABLE_NAME='射孔段爆轰计算参数'")
                If foundRows.Length = 0 Then '没有射孔段爆轰计算参数表
                    SQL_command = "CREATE TABLE 射孔段爆轰计算参数 " _
                        & " (工况序号 INTEGER        ,工况名称 TEXT(50)            ,射孔段井眼直径mm FLOAT  ,是否有坐封的封隔器 YESNO,射孔枪序号 INTEGER," _
                        & "筛管距射孔顶端距离m FLOAT ,封隔器距射孔顶端距离m FLOAT  ,口袋长度m FLOAT         ,射孔枪名称 TEXT(50)     ,[射孔密度孔╱m]  FLOAT    ,射孔弹名称 TEXT(50)  ," _
                        & "炸药名称 TEXT(50)         ,[装药量g]       FLOAT        ,[装药密度g╱cm3]  FLOAT ,射孔段长度m FLOAT       ,射孔段处井底温度℃ FLOAT  ,井底初始压力MPa  FLOAT," _
                        & "井口加压MPa FLOAT         ,爆热kJ╱mol   FLOAT          ,爆温K       FLOAT       ,爆容m3  FLOAT           ,爆速km╱s  FLOAT          ,爆压MPa  FLOAT           ," _
                        & "井底附加峰值压力MPa FLOAT ,封隔器处附加峰值压力MPa FLOAT,加速度峰值m╱s2 FLOAT   ,作业名称   TEXT(50)     ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年7月14日增加
                'Call database_creat(20015, 2)  '建立数据表，射孔段封隔器计算参数
                '*********************************************************************************************************************************
            Case 20015
                foundRows = dbSchema.Select("TABLE_NAME='射孔段封隔器计算参数'")
                If foundRows.Length = 0 Then '没有射孔段封隔器计算参数
                    SQL_command = "CREATE TABLE 射孔段封隔器计算参数 " _
                        & " (工况序号 INTEGER                ,工况名称  TEXT(50)                    ,钻油管序号  INTEGER              ,钻油管规格  TEXT(50)     ,元件性质 TEXT(30)," _
                        & "封隔器序号 INTEGER                ,封隔器名称 TEXT(50)                   ,射孔枪序号 INTEGER               ,射孔枪名称 TEXT(50)," _
                        & "射孔段垂深m FLOAT                 ,射孔段内压MPa  FLOAT                  ,射孔后负压MPa FLOAT              ,封隔器处下拽力kN    FLOAT," _
                        & "管柱下拽应力MPa FLOAT             ,管柱下拽安全系数 FLOAT                ,管柱上顶应力MPa  FLOAT           ,管柱上顶安全系数 FLOAT   ," _
                        & "封隔器中心管应力MPa FLOAT         ,封隔器中心管安全系数 FLOAT            ,温度校正管柱屈服强度MPa FLOAT    ,温度校正封隔器中心管屈服强度MPa FLOAT," _
                        & "地层校正管柱下拽应力MPa FLOAT     ,地层校正管柱上顶应力MPa  FLOAT        ,地层校正封隔器中心管应力MPa FLOAT,温度地层校正管柱下拽安全系数 FLOAT," _
                        & "温度地层校正管柱上顶安全系数 FLOAT,温度地层校正封隔器中心管安全系数 FLOAT,作业名称   TEXT(50)              ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年7月16日增加
                'Call database_creat(20017, 2)  '建立数据表，射孔油管应力计算参数
                '*********************************************************************************************************************************
            Case 20016
                foundRows = dbSchema.Select("TABLE_NAME='射孔油管应力计算参数'")
                If foundRows.Length = 0 Then '没有射孔油管应力计算参数
                    SQL_command = "CREATE TABLE 射孔油管应力计算参数 " _
                        & " (工况序号 INTEGER       ,工况名称 TEXT(50)     ,计算节点序号  INTEGER ,起始深度m  FLOAT         ,终止深度m  FLOAT    ,元件序号 INTEGER    ,元件规格 TEXT(50)   ," _
                        & "元件外径mm FLOAT         ,元件内径mm FLOAT      ,元件屈服强度MPa FLOAT ,活塞力kN FLOAT           ,管柱悬重kg FLOAT    ,管柱浮容重kN  FLOAT ,下甩力kN   FLOAT," _
                        & "下甩时的总载荷kN FLOAT   ,元件上部应力MPa FLOAT ,元件应力安全系数 FLOAT,元件位置附加压力MPa FLOAT,节点状态 TEXT(10)   ,作业名称 TEXT(50)   ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If
                '*********************************************************************************************************************************
                '  李润洲2025年7月25日增加
                'Call database_creat(20017, 2)  '建立数据表，管柱_筛管表
                '*********************************************************************************************************************************
            Case 20017
                foundRows = dbSchema.Select("TABLE_NAME='管柱_筛管表'")
                If foundRows.Length = 0 Then '没有管柱_筛管表
                    SQL_command = "CREATE TABLE 管柱_筛管表 " _
                        & " (作业名称 TEXT(50) ,元件序号  INTEGER       ,元件名称 TEXT(50)     ,壁厚mm FLOAT             ,单根长度m    FLOAT        ,单根质量kg FLOAT ," _
                        & "主体材料 TEXT(50)   ,接头抗拉伸强度kN FLOAT  ,主材屈服强度MPa FLOAT ,抗内压强度MPa  FLOAT     ,抗外挤强度MPa FLOAT       ," _
                        & "耐温℃ FLOAT        ,耐压差MPa FLOAT         ,[孔隙率%]  FLOAT      ,有效流通面积cm2╱m FLOAT ,拦截的最小颗粒直径mm FLOAT," _
                        & "生产厂家  TEXT(50)  ,备注    TEXT(200)       ,井号   TEXT(50))"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    EXECOleDbCommand.ExecuteNonQuery()
                End If

            Case Else ' 其他数值。
                Debug.Print("没有这种情况")
        End Select
        dbSchema.Dispose()
        cn.Close()
        cn.Dispose()
    End Sub
    '***************************************************************************************************************************
    '                                     判断所选择的accse数据库文件是否是程序要求格式的数据库文件
    '***************************************************************************************************************************
    Function check_file() As Boolean
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '全局变量连接字符串在此赋值
        use_AdoConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & use_dbname & ";Persist Security Info=False"
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        '初始值认为选择的文件合适
        check_file = True
        foundRows = dbSchema.Select("TABLE_NAME='油气井表'")
        If foundRows.Length = 0 Then '没有油气井表
            check_file = check_file And False
        Else
            check_file = check_file And True
        End If
        foundRows = dbSchema.Select("TABLE_NAME='套管数据表'")
        If foundRows.Length = 0 Then '没有套管数据表
            check_file = check_file And False
        Else
            check_file = check_file And True
        End If
        foundRows = dbSchema.Select("TABLE_NAME='管柱数据表'")
        If foundRows.Length = 0 Then '没有管柱数据表
            check_file = check_file And False
        Else
            check_file = check_file And True
        End If
        foundRows = dbSchema.Select("TABLE_NAME='管柱_封隔定位元件'")
        If foundRows.Length = 0 Then '没有管柱_封隔定位元件表
            check_file = check_file And False
        Else
            check_file = check_file And True
        End If
        foundRows = dbSchema.Select("TABLE_NAME='管柱_开关元件'")
        If foundRows.Length = 0 Then '没有管柱_开关元件表
            check_file = check_file And False
        Else
            check_file = check_file And True
        End If
        dbSchema.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function

    '***************************************************************************************************************************
    '                                     判断油气井表中是否有数据
    '20191214升级
    '***************************************************************************************************************************
    Function yqj_ok() As Boolean
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        yqj_ok = False
        SQL_command = "select * from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            yqj_ok = True
        Else
            yqj_ok = False
        End If
        RECreader.Close()
        EXECOleDbCommand.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function
    '***************************************************************************************************************************
    '                                     从参数表中读取参数，返回为字符串
    '20200109升级
    '***************************************************************************************************************************
    Function get_canshu(ByVal cs_name As String) As String
        Dim SQL_command As String
        get_canshu = ""
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 计算参数表 where 井号='" & well_name & "' and 参数名称='" & cs_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.Read Then
            get_canshu = RECreader.Item("参数值").ToString()
        End If
        RECreader.Close()
        EXECOleDbCommand.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function

    '***************************************************************************************************************************
    '                                     往计算参数表中写参数
    '20190914升级
    '***************************************************************************************************************************
    Sub write_chanshu(ByVal cs_name As String, ByVal cs_value As String)
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select * from 计算参数表 where 井号='" & well_name & "' and 参数名称='" & cs_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If RECreader.HasRows Then
            SQL_command = "update 计算参数表 set 参数值='" & cs_value & "' where 参数名称='" & cs_name & "'and 井号='" & well_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
        Else
            SQL_command = "insert into 计算参数表(参数名称,参数值,井号) values ('" & cs_name & "','" & cs_value & "','" & well_name & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
        End If
        RECreader.Close()
        EXECOleDbCommand.Dispose()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Sub
    '***************************************************************************************************************************
    '                                     管柱力学分析数据合法性、数据完整性检查函数
    ' 程序升级记事：
    '                                                                                           秦彦斌 2019年12月15日最后整理
    '   (1) 2019年12月15日升级时从M_gzlxfx_data_ok模块移过来，去除单独M_gzlxfx_data_ok模块。
    '   (2) 用GetOleDbSchemaTable方法查找库中的表。
    '                                                                                
    '参数说明：
    '编程思路：
    '   检查管柱力学分析所必须的数据是否完整
    '函数返回：逻辑型结果，数据合法，返回真，否则，返回假
    '***************************************************************************************************************************
    Public Function gzlxfx_data_ok() As Boolean
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim rowfilter As String
        gzlxfx_data_ok = True
        If well_name = "" Then
            msg_prompt = "没有指定井号，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlxfx_data_ok = False
            Exit Function
        End If

        If zuoye_name = "" Then
            msg_prompt = "没有指定管柱作业名称，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlxfx_data_ok = False
            Exit Function
        End If

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "井身结构数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlxfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        If Not TSM_ver_switch = 1 Then
            dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            rowfilter = "TABLE_NAME='测井数据表'"
            foundRows = dbSchema.Select(rowfilter)
            If foundRows.Length = 0 Then
                msg_prompt = "井斜数据数据没有建立，不能进行运算！"
                msg_buttons = 0 + 16
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                gzlxfx_data_ok = False
                RECreader.Close()
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            Else
                SQL_command = "select top 1 * from 测井数据表"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                RECreader = EXECOleDbCommand.ExecuteReader()
                If Not RECreader.HasRows Then
                    msg_prompt = "测井数据没有输入，不能进行运算！"
                    msg_buttons = 0 + 16
                    msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                    gzlxfx_data_ok = False
                    RECreader.Close()
                    cn_userdb.Close()
                    cn_userdb.Dispose()
                    Exit Function
                End If
            End If
            dbSchema.Clear()
            dbSchema.Dispose()
        End If
        SQL_command = "select top 1 * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "管柱数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlxfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        SQL_command = "select top 1 * from 工况参数表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "工况数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlxfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function
    '***************************************************************************************************************************
    '                                     摩阻摩矩分析数据合法性、数据完整性检查函数
    ' 程序升级记事：
    '                                                                                           秦彦斌 2020年2月10日最后整理
    '   (1) 2019年12月15日升级时从M_mzmjfx_data_ok模块移过来，去除单独M_mzmjfx_data_ok模块。
    '参数说明：
    '编程思路：
    '   检查摩阻摩矩分析所必须的数据是否完整
    '函数返回：逻辑型结果，数据合法，返回真，否则，返回假
    '***************************************************************************************************************************
    Public Function mzmjfx_data_ok() As Boolean
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim rowfilter As String
        mzmjfx_data_ok = True

        If well_name = "" Then
            msg_prompt = "没有指定井号，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            mzmjfx_data_ok = False
            Exit Function
        End If

        If zuoye_name = "" Then
            msg_prompt = "没有指定管柱作业名称，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            mzmjfx_data_ok = False
            Exit Function
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "井身结构数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            mzmjfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        rowfilter = "TABLE_NAME='测井数据表'"
        foundRows = dbSchema.Select(rowfilter)
        If foundRows.Length = 0 Then
            msg_prompt = "井斜数据数据没有建立，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            mzmjfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        Else
            SQL_command = "select top 1 * from 测井数据表"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            RECreader = EXECOleDbCommand.ExecuteReader()
            If Not RECreader.HasRows Then
                msg_prompt = "测井数据没有输入，不能进行运算！"
                msg_buttons = 0 + 16
                msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
                mzmjfx_data_ok = False
                RECreader.Close()
                cn_userdb.Close()
                cn_userdb.Dispose()
                Exit Function
            End If
        End If
        dbSchema.Clear()
        dbSchema.Dispose()
        SQL_command = "select top 1 * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "管柱数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            mzmjfx_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function
    '***************************************************************************************************************************
    '                                     管柱拉伸强度安全分析数据合法性、数据完整性检查函数
    '                                                                                           秦彦斌 2021年11月12日最后整理
    '程序功能：
    '    检查管柱拉伸强度安全分析所必须的数据是否完整：
    '    (1)指定井号，变量well_name有赋值
    '    (2)指定管柱作业名称，变量zuoye_name有赋值
    '    (3)有井身结构数据
    '    (4)有管柱数据
    '
    '程序升级记事：
    '    (1) 2021年11月12日因实现管柱拉伸强度安全分析增加此函数
    '
    '函数返回：逻辑型结果，数据合法，返回真，否则，返回假
    '***************************************************************************************************************************
    Public Function gzlsaqpj_data_ok() As Boolean
        gzlsaqpj_data_ok = True
        If well_name = "" Then
            msg_prompt = "没有指定井号，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlsaqpj_data_ok = False
            Exit Function
        End If

        If zuoye_name = "" Then
            msg_prompt = "没有指定管柱作业名称，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlsaqpj_data_ok = False
            Exit Function
        End If
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "井身结构数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlsaqpj_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        SQL_command = "select top 1 * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "管柱数据没有输入，不能进行运算！"
            msg_buttons = 0 + 16
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            gzlsaqpj_data_ok = False
            RECreader.Close()
            cn_userdb.Close()
            cn_userdb.Dispose()
            Exit Function
        End If
        RECreader.Close()
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Function
    '***************************************************************************************************************************
    '                                     磨损套管剩余强度分析数据合法性、数据完整性检查函数
    ' 程序升级记事：
    '                                                                                           秦彦斌 2020年2月23日最后整理
    '   (1) 2019年12月15日升级时从tgwear_data_ok模块移过来，去除单独tgwear_data_ok模块。
    '参数说明：
    '编程思路：
    '   检查磨损套管剩余强度分析所必须的数据是否完整
    '函数返回：逻辑型结果，数据合法，返回真，否则，返回假
    '***************************************************************************************************************************
    Public Function tgwear_data_ok() As Boolean
        Dim SQL_command As String
        Dim AccessTab As String
        Dim DC_ok As Boolean
        Dim WDP_ok As Boolean
        Dim DP_ok As Boolean

        tgwear_data_ok = True
        DC_ok = True
        WDP_ok = True
        DP_ok = True

        On Error GoTo ErrHandler

        '油气井基本参数是否输入判断
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        SQL_command = "select top 1 * from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "油气井表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        '井身结构数据是否输入判断
        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "井身结构数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        '钻井日志数据是否输入判断
        SQL_command = "select top 1 * from 钻井日志表" ' where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "钻井日志表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        '测井数据表是否输入判断
        SQL_command = "select top 1 * from 测井数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "测井数据表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        SQL_command = "select top 1 * from 钻具组合表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "钻具组合表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        SQL_command = "select top 1 * from 钻具组合_加重钻杆表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            WDP_ok = False
        End If
        SQL_command = "select top 1 * from 钻具组合_普通钻杆表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            DP_ok = False
        End If
        SQL_command = "select top 1 * from 钻具组合_钻铤表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            DC_ok = False
        End If
        If (WDP_ok = False) And (DP_ok = False) And (DC_ok = False) Then
            msg_prompt = "钻具组合表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        SQL_command = "select top 1 * from 钻具组合_普通钻杆表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        If Not RECreader.HasRows Then
            msg_prompt = "钻具组合_普通钻杆表数据没有输入，不能进行运算！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
            tgwear_data_ok = False
            Exit Function
        End If
        Exit Function
ErrHandler:
        msg_prompt = "数据读取出错,请检查测井数据和钻井日志数据库表结构！"
        msg_buttons = 0 + 48
        msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        tgwear_data_ok = False
    End Function

    '***************************************************************************************************************************
    '                                       颜色查取函数
    '                                                                              秦彦斌 2020年1月3日
    '   传入颜色序号，1开头，返回Uint32值
    '***************************************************************************************************************************
    Public Function GetUint32color(ByVal xuhao As Integer) As UInteger
        Dim c1 As Integer
        Dim c2 As Integer
        Dim c3 As Integer
        Dim foundRows() As DataRow
        Dim rowfilter As String
        Dim color_in_tab As Boolean
        Dim need_crt_color As Boolean
        Dim HEXcolor_name As String
        need_crt_color = False
        cn_basedb = New System.Data.OleDb.OleDbConnection(AdoConString)
        cn_basedb.Open()
        SQL_command = "select * from 绘图颜色表 order by 序号"
        ad.SelectCommand = New OleDbCommand(SQL_command, cn_basedb)
        color_Table.Clear()
        ad.Fill(color_Table)
        ad.Dispose()
        If xuhao > 0 And xuhao < color_Table.Rows.Count Then
            rowfilter = "序号=" & xuhao.ToString
            foundRows = color_Table.Select(rowfilter)
            If foundRows.Length > 0 Then
                GetUint32color = System.Convert.ToInt32(foundRows(0).Item("HEXBGR色名").ToString, 16)
            Else
                need_crt_color = True
            End If
        Else
            need_crt_color = True
        End If
        If need_crt_color = True Then
            color_in_tab = False
            Do
                c1 = Val(Int((255 - 0 + 1) * Rnd() + 0))
                c2 = Val(255 - Int((255 - 0 + 1) * Rnd() + 0))
                c3 = Val(Int((255 - 0 + 1) * Rnd() + 0))
                HEXcolor_name = Mid(Color.FromArgb(255, c1, c2, c3).Name, 2, 6)
                rowfilter = "HEXBGR色名='" & HEXcolor_name & "'"
                foundRows = color_Table.Select(rowfilter)
                If foundRows.Length > 0 Then
                    color_in_tab = True
                Else
                    color_in_tab = False
                End If
            Loop While c1 > 128 And c2 > 128 And c3 > 128 And color_in_tab = True
            GetUint32color = System.Convert.ToInt32(HEXcolor_name, 16)
        End If
        cn_basedb.Close()
    End Function
    '*********************************************************************************************************************************************
    ' 删除指定作业的所有数据
    ' 返回删除的表的个数
    ' 20251112秦彦斌从ipt_mdceng.vb中移至这里
    '*********************************************************************************************************************************************
    Public Function delete_zy(ByVal zyname As String) As Integer
        Dim table_name As String
        Dim col_name As String
        Dim i As Short = 0
        Dim hasZYField As Boolean
        Dim hasJHField As Boolean
        Dim SQL_command As String
        '***********************************************************************************
        ' 遍历库中所有表，删除 井号=well_name、作业名称=zyname 的所有记录
        '***********************************************************************************
        i = 0
        Try
            Using cn_userdb As New System.Data.OleDb.OleDbConnection(use_AdoConString)
                cn_userdb.Open()
                Using dbSchema As DataTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                    If dbSchema.Rows.Count > 0 Then
                        For Each dbschema_row As DataRow In dbSchema.Rows
                            hasZYField = False
                            hasJHField = False
                            table_name = dbschema_row.Item("TABLE_NAME").ToString
                            Using columnTable As DataTable = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, table_name, Nothing})
                                For Each columnTab_row As DataRow In columnTable.Rows
                                    col_name = columnTab_row.Item("COLUMN_NAME").ToString
                                    If col_name = "作业名称" Then
                                        hasZYField = True
                                    End If
                                    If col_name = "井号" Then
                                        hasJHField = True
                                    End If
                                Next
                                If hasZYField And hasJHField Then
                                    '如果表有“作业名称”和“井号”字段，且存有“井号=well_name、作业名称=zyname”的记录，删除记录
                                    SQL_command = "select * from " & table_name & " where 井号='" & well_name & "' And 作业名称='" & zyname & "'"
                                    Using EXECOleDbCommand As New OleDbCommand(SQL_command, cn_userdb)
                                        Using RECreader As OleDbDataReader = EXECOleDbCommand.ExecuteReader()
                                            If RECreader.Read Then
                                                SQL_command = "delete * from " & table_name & " where 井号='" & well_name & "' And 作业名称='" & zyname & "'"
                                                Using EXECOleDbCommand2 As New OleDbCommand(SQL_command, cn_userdb)
                                                    EXECOleDbCommand2.ExecuteNonQuery()
                                                End Using
                                                i = i + 1
                                            End If
                                        End Using
                                    End Using
                                End If 'hasZYField And hasJHField 判断结束
                            End Using
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            msg_prompt = "删除作业‘" & zyname & "’时出错，请检查数据表内容！"
            msg_buttons = 0 + 48
            msg_return = MsgBox(msg_prompt, msg_buttons, sofe_name)
        Finally
            delete_zy = i
        End Try
    End Function
End Module