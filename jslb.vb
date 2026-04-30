Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_jslb
    '***************************************************************************************************************************
    '                            给定节点间距、管柱长度、库表名称生成分段计算列表子程序
    '                                                                                         2021年11月30日秦彦斌最后修订整理
    '输入参数：
    '  max_dlt_L               划分节点时节点的最大间距            类型：Double
    '  guanzhu_length          管柱下入长度 <=管柱总长             类型：Double
    '  table_name              节点情况表名称                      类型：string
    '全局变量使用：
    '   well_name:             当前井号
    '   zuoye_name:            当前管柱作业名称
    '编程思路：
    '  1、删除临时分段情况表
    '  2、由于划分节点计算是所有计算的第一个，为了防止数据重复，计算前先删除该作业下的数据。
    '  3、建立临时分段情况表
    '  4、把管柱各段加入临时分段情况表
    '  5、加入井口数据,井口下深为0
    '  6、以段点下深排序并追加到节点情况表,按max_dlt_L长划分节点,加入过程中，赋必要的初值,顺便填写节点计算参数表
    '  7、删除临时数据表
    '
    '使用注意事项：
    '   使用前需用函数gzlxfx_data_ok()判断数据是否完整，否则运行程序会出错。
    '
    '运行结果：在所给定的节点情况表中填入合适的数据
    '修订情况：
    '    20171203,修改程序，由原来的节点间距一给定，改为节点间距、管柱长度，库表名称三给定
    '    20200120,程序迁移到VS2008完成，数据读取用ADO.NET提供的方法
    '    20211130,套管节点对管柱力学分析没影响，不用了；不可能有同一深度的重合点，故不需判定。管柱元件两头都设节点，原来除井口外
    '             仅有下端。
    '    20220911,
    '***************************************************************************************************************************
    Sub jslb(ByVal max_dlt_L As Double, ByVal guanzhu_length As Double, ByVal table_name As String) '生成分段计算列表
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim ad As New System.Data.OleDb.OleDbDataAdapter
        Dim dbSchema As DataTable
        Dim SQL_command As String
        Dim foundRows() As DataRow
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim depth As Double
        Dim dlt_L As Double
        'Dim RECreader2 As OleDbDataReader
        'Dim temp_str As String
        'Dim temp_str2 As String
        'Dim zhijing As Double
        'Dim tg_depth As Double
        'Dim cengshu As Double
        Dim i As Short
        Dim j As Short
        Dim last_depth As Double
        Dim last_csh As Double
        Dim last_jxj As Double
        Dim last_fwj As Double
        Dim devid_num As Short
        Dim last_ID As String
        Dim last_xingzhi As String
        Dim last_leixing As String
        Dim ret_jxj As Double '节点井斜角
        Dim ret_fwj As Double '节点方位角
        Dim ret_csh As Double '垂深临时变量

        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        '*************************************************************************************************************************************
        '1、删除临时分段情况表
        '*************************************************************************************************************************************
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = dbSchema.Select("TABLE_NAME='lsb_fdqk'")
        If foundRows.Length <> 0 Then
            SQL_command = "DROP TABLE lsb_fdqk"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        dbSchema.Dispose()
        '*************************************************************************************************************************************
        '2、由于划分节点计算是所有计算的第一个，为了防止数据重复，计算前先删除该作业下的数据。
        '*************************************************************************************************************************************
        SQL_command = "delete * from  " & table_name & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        '*************************************************************************************************************************************
        '3、建立临时分段情况表
        '*************************************************************************************************************************************
        SQL_command = "select * into lsb_fdqk from  " & table_name & "  where 0>1"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        EXECOleDbCommand.ExecuteNonQuery()
        EXECOleDbCommand.Dispose()
        '***************************************************************************************************************************
        '4、把管柱各段加入临时分段情况表
        '***************************************************************************************************************************
        depth = guanzhu_length
        SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号 DESC"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
                    & "'" & zuoye_name & "'," & "'管柱'," & "'" & RECreader.Item("元件序号").ToString & "'," & "'" & RECreader.Item("元件性质").ToString & "'," & CStr(depth) & "," _
                    & "'" & well_name & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
            depth = depth - Val(RECreader.Item("元件长度m").ToString)
            If depth <= 0 Then
                '把最后一个元件加入，退出。20220911
                depth = 0.0#
                SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
                        & "'" & zuoye_name & "'," & "'管柱'," & "'" & RECreader.Item("元件序号").ToString & "'," & "'" & RECreader.Item("元件性质").ToString & "'," & CStr(depth) & "," _
                        & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
                Exit While
            End If
        End While
        RECreader.Close()

        ''***************************************************************************************************************************
        ''3、把管柱最下端以上内侧套管各段加入临时分段情况表
        ''  由于井口下深为零，故悬挂深度为0的套管不用加入，但后面加入的套管内径要比井口的内径要小。
        ''***************************************************************************************************************************
        ''列出所有悬挂深度
        'SQL_command = "select distinct 悬挂深度m from 套管数据表 where 井号='" & well_name & "' and 悬挂深度m<=" & CStr(guanzhu_length) & " order by 悬挂深度m"
        'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        'RECreader = EXECOleDbCommand.ExecuteReader()
        'If RECreader.HasRows Then
        '    '取第一个悬挂深度套管参数
        '    RECreader.Read()
        '    SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "' and  悬挂深度m=" & RECreader.Item("悬挂深度m").ToString & " order by 套管外径mm"
        '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        '    RECreader2 = EXECOleDbCommand.ExecuteReader()
        '    RECreader2.Read()
        '    zhijing = Val(RECreader2.Item("套管外径mm").ToString)
        '    cengshu = Val(RECreader2.Item("套管层数").ToString)
        '    RECreader2.Close()
        '    '悬挂深度循环
        '    While RECreader.Read
        '        tg_depth = Val(RECreader.Item("悬挂深度m").ToString)
        '        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "' and  悬挂深度m=" & RECreader.Item("悬挂深度m").ToString & " order by 套管外径mm"
        '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        '        RECreader2 = EXECOleDbCommand.ExecuteReader()
        '        RECreader2.Read()
        '        '判断是否与上一个在一个层内，若是同一层，直接加入，若不在同一层，需判断直径看是否在内侧
        '        If Val(RECreader2.Item("套管层数").ToString) = cengshu Then
        '            zhijing = Val(RECreader2.Item("套管外径mm").ToString)
        '            cengshu = Val(RECreader2.Item("套管层数").ToString)
        '            temp_str = RECreader2.Item("套管层数").ToString & "-" & RECreader2.Item("套管段数").ToString
        '            temp_str2 = RECreader2.Item("套管类型").ToString
        '            SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
        '                & "'" & zuoye_name & "'," & "'套管'," & "'" & temp_str & "'," & "'" & temp_str2 & "'," & CStr(tg_depth) & "," & "'" & well_name & "')"
        '            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        '            EXECOleDbCommand.ExecuteNonQuery()
        '        Else
        '            '与上一个直径比较，若小于上一个，表示比上一个靠内，加入
        '            If Val(RECreader2.Item("套管外径mm").ToString) <= zhijing Then
        '                zhijing = Val(RECreader2.Item("套管外径mm").ToString)
        '                cengshu = Val(RECreader2.Item("套管层数").ToString)
        '                temp_str = RECreader2.Item("套管层数").ToString & "-" & RECreader2.Item("套管段数").ToString
        '                temp_str2 = RECreader2.Item("套管类型").ToString
        '                SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
        '                    & "'" & zuoye_name & "'," & "'套管'," & "'" & temp_str & "'," & "'" & temp_str2 & "'," & CStr(tg_depth) & "," & "'" & well_name & "')"
        '                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        '                EXECOleDbCommand.ExecuteNonQuery()
        '            End If
        '        End If
        '        RECreader2.Close()
        '    End While
        'End If

        '***************************************************************************************************************************
        '5、加入井口数据,井口下深为0
        '***************************************************************************************************************************
        'last_depth = 0.0#
        'dlt_L = 0.0#
        'last_ID = "1"
        'last_xingzhi = "井口"
        'last_leixing = "井口"
        'j = 1
        'ret_jxj = 0.0#
        'ret_fwj = 0.0#
        'ret_csh = last_depth
        'last_csh = last_depth
        'SQL_command = "insert into   " & table_name _
        '    & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
        '    & "'" & zuoye_name & "'," & "'" & "井口" & "'," & "'" & last_ID & "'," & "'" & "井口" & "'," & CStr(last_depth) & "," _
        '    & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," & CStr(ret_fwj) & "," & "'" & well_name & "')"
        'EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        'EXECOleDbCommand.ExecuteNonQuery()
        'EXECOleDbCommand.Dispose()
        '*************************************************************************************************************************************
        ' 6、以段点下深排序并追加到节点情况表,按max_dlt_L长划分节点,加入过程中，赋必要的初值,顺便填写节点计算参数表
        '*************************************************************************************************************************************
        dlt_L = 0.0#
        j = 0
        SQL_command = "select * from lsb_fdqk order by 节点下深m"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        While RECreader.Read
            '井口
            If j = 0 Then
                last_depth = 0.0#
                ret_jxj = 0.0#
                ret_fwj = 0.0#
                ret_csh = last_depth
                last_csh = last_depth
                j = j + 1
                SQL_command = "insert into  " & table_name _
                    & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
                    & "'" & zuoye_name & "'," & "'" & RECreader.Item("节点类型").ToString & "'," & "'" & RECreader.Item("节点ID").ToString & "'," _
                    & "'" & RECreader.Item("节点性质").ToString & "'," & CStr(last_depth) & "," & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," _
                    & CStr(ret_fwj) & "," & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            End If
            '不是第一个，且不是井口，写分段的上面点，20220911
            If j <> 1 And last_depth <> 0 Then
                ret_jxj = last_jxj
                ret_fwj = last_fwj
                ret_csh = last_csh
                j = j + 1
                SQL_command = "insert into  " & table_name _
                    & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
                    & "'" & zuoye_name & "'," & "'" & RECreader.Item("节点类型").ToString & "'," & "'" & RECreader.Item("节点ID").ToString & "'," _
                    & "'" & RECreader.Item("节点性质").ToString & "'," & CStr(last_depth) & "," & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," _
                    & CStr(ret_fwj) & "," & "'" & well_name & "')"
                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                EXECOleDbCommand.ExecuteNonQuery()
                EXECOleDbCommand.Dispose()
            End If
            last_ID = RECreader.Item("节点ID").ToString
            last_xingzhi = RECreader.Item("节点性质").ToString
            last_leixing = RECreader.Item("节点类型").ToString

            '（当前节点下深-上一节点下深）> 节点间距，则分段
            If Val(RECreader.Item("节点下深m").ToString) - last_depth > max_dlt_L Then
                devid_num = Int((Val(RECreader.Item("节点下深m").ToString) - last_depth) / max_dlt_L + 0.5)
                dlt_L = (Val(RECreader.Item("节点下深m").ToString) - last_depth) / devid_num
                For i = 1 To devid_num - 1
                    last_depth = last_depth + dlt_L
                    j = j + 1
                    ret_jxj = 0.0#
                    ret_fwj = 0.0#
                    ret_csh = last_depth
                    If Not TSM_ver_switch = 1 Then
                        Call cal_jx_fw_cs(last_depth, ret_jxj, ret_fwj, ret_csh)
                    End If
                    SQL_command = "insert into   " & table_name _
                        & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
                        & "'" & zuoye_name & "'," & "'" & "管柱" & "'," & "'" & last_ID & "'," & "'" & "计算点" & "'," & CStr(last_depth) & "," _
                        & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," & CStr(ret_fwj) & "," & "'" & well_name & "')"
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
                    EXECOleDbCommand.ExecuteNonQuery()
                    EXECOleDbCommand.Dispose()
                Next i
            End If

            '写分段的下面点
            ret_jxj = 0.0#
            ret_fwj = 0.0#
            ret_csh = Val(RECreader.Item("节点下深m").ToString)
            If Not TSM_ver_switch = 1 Then
                Call cal_jx_fw_cs(Val(RECreader.Item("节点下深m").ToString), ret_jxj, ret_fwj, ret_csh)
            End If
            j = j + 1
            SQL_command = "insert into  " & table_name _
                & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
                & "'" & zuoye_name & "'," & "'" & RECreader.Item("节点类型").ToString & "'," & "'" & RECreader.Item("节点ID").ToString & "'," _
                & "'" & RECreader.Item("节点性质").ToString & "'," & RECreader.Item("节点下深m").ToString & "," & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," _
                & CStr(ret_fwj) & "," & "'" & well_name & "')"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()

            last_depth = Val(RECreader.Item("节点下深m").ToString)
            last_jxj = ret_jxj
            last_fwj = ret_fwj
            last_csh = ret_csh
            last_ID = RECreader.Item("节点ID").ToString
            last_xingzhi = RECreader.Item("节点性质").ToString
            last_leixing = RECreader.Item("节点类型").ToString
        End While
        RECreader.Close()
        '***************************************************************************************************************************
        '  7、删除临时数据表
        '***************************************************************************************************************************
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = dbSchema.Select("TABLE_NAME='lsb_fdqk'")
        If foundRows.Length <> 0 Then
            SQL_command = "DROP TABLE lsb_fdqk"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
            EXECOleDbCommand.ExecuteNonQuery()
            EXECOleDbCommand.Dispose()
        End If
        dbSchema.Dispose()
        'MsgBox "运行生成计算列表子程序"    '调试语句
        cn_userdb.Close()
        cn_userdb.Dispose()
    End Sub

    '                                             原算法备份---20211129
    '***************************************************************************************************************************
    '                            给定节点间距、给定管柱长度，给定库表名称，生成分段计算列表子程序
    '输入参数：
    '  max_dlt_L               划分节点时节点的最大间距            类型：Double
    '  guanzhu_length          管柱下入长度 <=管柱总长             类型：Double
    '  table_name              节点情况表名称                      类型：string
    '全局变量使用：
    '   well_name:             当前井号
    '   zuoye_name:            当前管柱作业名称
    '编程思路：
    '  1、把管柱各段加入计算列表
    '  2、把管柱最下端以上内侧套管各段加入计算列表
    '  3、以段点下深排序并追加到节点情况表
    '  4、建立与节点计算参数表结构相同的临时表，将每一工况下的井口、液面节点加入，将前面的各节点加入，注意填写垂深。
    '  5、对每一工况按节点垂深排序，合并相同垂深的节点，填写节点计算参数表数据，赋必要的初值
    '  6、删除临时数据表
    '
    '使用注意事项：
    '   使用前需用函数gzlxfx_data_ok()判断数据是否完整，否则运行程序会出错。
    '
    '运行结果：在所给定的节点情况表中填入合适的数据
    '修订情况：
    '    20171203,修改程序，由原来的节点间距一给定，改为节点间距、管柱长度，库表名称三给定
    '    20200120,程序迁移到VS2008完成，数据读取用ADO.NET提供的方法
    '***************************************************************************************************************************
    'Sub jslb(ByVal max_dlt_L As Double, ByVal guanzhu_length As Double, ByVal table_name As String) '生成分段计算列表
    '    Dim cn_userdb As System.Data.OleDb.OleDbConnection
    '    Dim ad As New System.Data.OleDb.OleDbDataAdapter
    '    Dim dbSchema As DataTable
    '    Dim SQL_command As String
    '    Dim foundRows() As DataRow
    '    Dim EXECOleDbCommand As OleDbCommand
    '    Dim RECreader As OleDbDataReader
    '    Dim RECreader2 As OleDbDataReader
    '    Dim RECreader3 As OleDbDataReader

    '    Dim temp_str As String
    '    Dim temp_str2 As String
    '    Dim depth As Double
    '    Dim zhijing As Double
    '    Dim tg_depth As Double
    '    Dim dlt_L As Double
    '    Dim cengshu As Double
    '    Dim i As Short
    '    Dim j As Short
    '    Dim last_depth As Double
    '    Dim devid_num As Short
    '    Dim last_ID As String
    '    Dim last_xingzhi As String
    '    Dim last_leixing As String
    '    Dim ret_jxj As Double '节点井斜角
    '    Dim ret_fwj As Double '节点方位角
    '    Dim ret_csh As Double '垂深临时变量
    '    cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
    '    cn_userdb.Open()
    '    dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
    '    foundRows = dbSchema.Select("TABLE_NAME='lsb_fdqk'")
    '    If foundRows.Length <> 0 Then
    '        SQL_command = "DROP TABLE lsb_fdqk"
    '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '        EXECOleDbCommand.ExecuteNonQuery()
    '    End If
    '    dbSchema.Dispose()
    '    SQL_command = "select * into lsb_fdqk from  " & table_name & "  where 0>1"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    EXECOleDbCommand.ExecuteNonQuery()
    '    '***************************************************************************************************************************
    '    '1、把管柱各段加入计算列表
    '    '***************************************************************************************************************************
    '    depth = guanzhu_length
    '    i = 0
    '    SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号 DESC"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    RECreader = EXECOleDbCommand.ExecuteReader()
    '    While RECreader.Read
    '        SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
    '                & "'" & zuoye_name & "'," & "'管柱'," & "'" & RECreader.Item("元件序号").ToString & "'," & "'" & RECreader.Item("元件性质").ToString & "'," & CStr(depth) & "," _
    '                & "'" & well_name & "')"
    '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '        EXECOleDbCommand.ExecuteNonQuery()
    '        depth = depth - Val(RECreader.Item("元件长度m").ToString)
    '        If depth < 0 Then
    '            depth = 0.0#
    '            i = i + 1
    '            If i > 1 Then
    '                Exit While
    '            End If
    '        End If
    '    End While
    '    RECreader.Close()
    '    '***************************************************************************************************************************
    '    '2、加入井口数据,井口下深为0
    '    '***************************************************************************************************************************
    '    SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
    '        & "'" & zuoye_name & "','井口','井口','井口'," & CStr(0) & ",'" & well_name & "')"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    EXECOleDbCommand.ExecuteNonQuery()
    '    '***************************************************************************************************************************
    '    '3、把管柱最下端以上内侧套管各段加入计算列表
    '    '  由于井口下深为零，故悬挂深度为0的套管不用加入，但后面加入的套管内径要比井口的内径要小。
    '    '***************************************************************************************************************************
    '    '列出所有悬挂深度
    '    SQL_command = "select distinct 悬挂深度m from 套管数据表 where 井号='" & well_name & "' and 悬挂深度m<=" & CStr(guanzhu_length) & " order by 悬挂深度m"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    RECreader = EXECOleDbCommand.ExecuteReader()
    '    If RECreader.HasRows Then
    '        '取第一个悬挂深度套管参数
    '        RECreader.Read()
    '        SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "' and  悬挂深度m=" & RECreader.Item("悬挂深度m").ToString & " order by 套管外径mm"
    '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '        RECreader2 = EXECOleDbCommand.ExecuteReader()
    '        RECreader2.Read()
    '        zhijing = Val(RECreader2.Item("套管外径mm").ToString)
    '        cengshu = Val(RECreader2.Item("套管层数").ToString)
    '        RECreader2.Close()
    '        '悬挂深度循环
    '        While RECreader.Read
    '            tg_depth = Val(RECreader.Item("悬挂深度m").ToString)
    '            SQL_command = "select top 1 * from 套管数据表 where 井号='" & well_name & "' and  悬挂深度m=" & RECreader.Item("悬挂深度m").ToString & " order by 套管外径mm"
    '            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '            RECreader2 = EXECOleDbCommand.ExecuteReader()
    '            RECreader2.Read()
    '            '判断是否与上一个在一个层内，若是同一层，直接加入，若不在同一层，需判断直径看是否在内侧
    '            If Val(RECreader2.Item("套管层数").ToString) = cengshu Then
    '                zhijing = Val(RECreader2.Item("套管外径mm").ToString)
    '                cengshu = Val(RECreader2.Item("套管层数").ToString)
    '                temp_str = RECreader2.Item("套管层数").ToString & "-" & RECreader2.Item("套管段数").ToString
    '                temp_str2 = RECreader2.Item("套管类型").ToString
    '                SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
    '                    & "'" & zuoye_name & "'," & "'套管'," & "'" & temp_str & "'," & "'" & temp_str2 & "'," & CStr(tg_depth) & "," & "'" & well_name & "')"
    '                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                EXECOleDbCommand.ExecuteNonQuery()
    '            Else
    '                '与上一个直径比较，若小于上一个，表示比上一个靠内，加入
    '                If Val(RECreader2.Item("套管外径mm").ToString) <= zhijing Then
    '                    zhijing = Val(RECreader2.Item("套管外径mm").ToString)
    '                    cengshu = Val(RECreader2.Item("套管层数").ToString)
    '                    temp_str = RECreader2.Item("套管层数").ToString & "-" & RECreader2.Item("套管段数").ToString
    '                    temp_str2 = RECreader2.Item("套管类型").ToString
    '                    SQL_command = "insert into  lsb_fdqk (作业名称,节点类型,节点ID,节点性质,节点下深m,井号) values (" _
    '                        & "'" & zuoye_name & "'," & "'套管'," & "'" & temp_str & "'," & "'" & temp_str2 & "'," & CStr(tg_depth) & "," & "'" & well_name & "')"
    '                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                    EXECOleDbCommand.ExecuteNonQuery()
    '                End If
    '            End If
    '            RECreader2.Close()
    '        End While
    '    End If
    '    '*************************************************************************************************************************************
    '    ' 4、以段点下深排序并追加到节点情况表,按max_dlt_L长划分节点,加入过程中，将重复下深的结点合并，赋必要的初值,顺便填写节点计算参数表
    '    '*************************************************************************************************************************************
    '    '   1、由于划分节点计算是所有计算的第一个，为了防止数据重复，计算前先删除该工况下的数据。
    '    SQL_command = "delete * from  " & table_name & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "'"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    EXECOleDbCommand.ExecuteNonQuery()
    '    SQL_command = "select * from lsb_fdqk order by 节点下深m"
    '    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '    RECreader = EXECOleDbCommand.ExecuteReader()
    '    last_depth = 0.0#
    '    dlt_L = 0.0#
    '    last_ID = "井口"
    '    last_xingzhi = "井口"
    '    last_leixing = "井口"
    '    j = 0
    '    While RECreader.Read
    '        '（当前节点下深-上一节点下深）> 节点间距，则分段
    '        If Val(RECreader.Item("节点下深m").ToString) - last_depth > max_dlt_L Then
    '            devid_num = Int((Val(RECreader.Item("节点下深m").ToString) - last_depth) / max_dlt_L + 0.5)
    '            dlt_L = (Val(RECreader.Item("节点下深m").ToString) - last_depth) / devid_num
    '            For i = 1 To devid_num - 1
    '                last_depth = last_depth + dlt_L
    '                SQL_command = "select * from  " & table_name & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点下深m=" & CStr(last_depth)
    '                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                RECreader3 = EXECOleDbCommand.ExecuteReader()
    '                If RECreader3.Read Then
    '                    temp_str = RECreader3.Item("节点类型").ToString
    '                    temp_str2 = "计算点"
    '                    If Not InStr(temp_str, temp_str2) Then
    '                        temp_str = temp_str & "," & temp_str2
    '                        SQL_command = "update  " & table_name & "  set " _
    '                            & "节点类型='" & temp_str & "'" & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点下深m=" & CStr(last_depth)
    '                        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                        EXECOleDbCommand.ExecuteNonQuery()
    '                    End If
    '                Else
    '                    j = j + 1
    '                    ret_jxj = 0.0#
    '                    ret_fwj = 0.0#
    '                    ret_csh = last_depth
    '                    If Not TSM_ver_switch = 1 Then
    '                        Call cal_jx_fw_cs(last_depth, ret_jxj, ret_fwj, ret_csh)
    '                    End If
    '                    SQL_command = "insert into   " & table_name _
    '                        & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
    '                        & "'" & zuoye_name & "'," & "'" & "计算点" & "'," & "'" & last_ID & "'," & "'" & "计算点" & "'," & CStr(last_depth) & "," _
    '                        & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," & CStr(ret_fwj) & "," & "'" & well_name & "')"
    '                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                    EXECOleDbCommand.ExecuteNonQuery()
    '                End If
    '                RECreader3.Close()
    '            Next i
    '        End If

    '        SQL_command = "select * from  " & table_name _
    '            & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点下深m=" & RECreader.Item("节点下深m").ToString
    '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '        RECreader3 = EXECOleDbCommand.ExecuteReader()
    '        If RECreader3.Read Then
    '            temp_str = RECreader3.Item("节点类型").ToString
    '            temp_str2 = RECreader.Item("节点类型").ToString
    '            If Not InStr(temp_str, temp_str2) Then
    '                temp_str = temp_str & "," & temp_str2
    '                SQL_command = "update  " & table_name & "  set " _
    '                & "节点类型='" & temp_str & "'" & "  where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' and 节点下深m=" & CStr(last_depth)
    '                EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '                EXECOleDbCommand.ExecuteNonQuery()
    '            End If
    '        Else
    '            ret_jxj = 0.0#
    '            ret_fwj = 0.0#
    '            ret_csh = Val(RECreader.Item("节点下深m").ToString)
    '            If Not TSM_ver_switch = 1 Then
    '                Call cal_jx_fw_cs(Val(RECreader.Item("节点下深m").ToString), ret_jxj, ret_fwj, ret_csh)
    '            End If
    '            j = j + 1
    '            SQL_command = "insert into  " & table_name _
    '                & " (作业名称,节点类型,节点ID,节点性质,节点下深m,节点编号,节点垂深m,井斜角rad,方位角rad,井号) values (" _
    '                & "'" & zuoye_name & "'," & "'" & RECreader.Item("节点类型").ToString & "'," & "'" & RECreader.Item("节点ID").ToString & "'," _
    '                & "'" & RECreader.Item("节点性质").ToString & "'," & RECreader.Item("节点下深m").ToString & "," & CStr(j) & "," & CStr(ret_csh) & "," & CStr(ret_jxj) & "," _
    '                & CStr(ret_fwj) & "," & "'" & well_name & "')"
    '            EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '            EXECOleDbCommand.ExecuteNonQuery()
    '        End If
    '        RECreader3.Close()
    '        last_depth = Val(RECreader.Item("节点下深m").ToString)
    '        last_ID = RECreader.Item("节点ID").ToString
    '        last_xingzhi = RECreader.Item("节点性质").ToString
    '        last_leixing = RECreader.Item("节点类型").ToString
    '    End While
    '    RECreader.Close()
    '    '***************************************************************************************************************************
    '    '  6、删除临时数据表
    '    '***************************************************************************************************************************
    '    dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
    '    foundRows = dbSchema.Select("TABLE_NAME='lsb_fdqk'")
    '    If foundRows.Length <> 0 Then
    '        SQL_command = "DROP TABLE lsb_fdqk"
    '        EXECOleDbCommand = New OleDbCommand(SQL_command, cn_userdb)
    '        EXECOleDbCommand.ExecuteNonQuery()
    '    End If
    '    dbSchema.Dispose()
    '    'MsgBox "运行生成计算列表子程序"    '调试语句
    '    cn_userdb.Close()
    '    cn_userdb.Dispose()
    'End Sub
End Module
