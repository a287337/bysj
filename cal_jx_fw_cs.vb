Option Strict Off
Option Explicit On
Imports System.Data.OleDb
Module M_cal_jx_fw_cs
    '************************************************************************************************************************************************
    '                     用线性插值法计算当前井任意深度in_depth处的井斜角、方位角、垂深子程序
    ' 程序升级记事：
    '                                                                                                           秦彦斌 2020年1月26日最后整理更新
    '   20190717-从VB6升级到VS2008完成，但数据库技术用ADODB。
    '   2020年1月26日，数据库用ADO.NET操作迁移完成。由于此子程序需要频繁调用，改原来频繁读取井斜数据表的方法，利用ADO.NET提供的DataTable，设置全局
    '变量Testwell_Table，一次填充，多次使用，避免频繁读取数据库。
    '
    '全局变量使用：
    '   well_name         井号
    '   Testwell_Table    测斜数据表，考虑到此程序要频繁调用，故测斜数据由调用它的代码打开并读取到Table中，避免频繁读取数据库,若没有打开，则本程序
    '                     打开。
    '
    '输入参数：
    '   in_depth          任意深度（这里应该是测深）
    '
    '编程思路：
    '   如果要求点的测深小于1m，直接返回0。否则，遍历井斜数据表
    '       获取节点测深、井斜角参数
    '       如果要求点的测深等于节点测深，返回井斜角
    '       如果要求点的测深介于当前节点与前一节点之间，用内插法求得要求点的井斜角
    '   遍历完成
    '   如果要求点的测深大于最深节点的测深，认为此点的井斜角和方位角就等于最深节点的井斜角和方位角，即认为井保持井斜、方位不变继续“直”进而成。
    '
    '算法依据:
    '
    '
    '本子程度使用注意事项：
    '   任意深度in_depth不得大于井的钻深，但程序中未做判断。
    '       原因1：是从哪里取得钻深：（1）井的信息，（2）井深结构数据
    '       原因2：代码放在什么位置合适。此为频繁使用的函数，放这里不太妥当
    '
    '子程序以传地址的方式计算出给定井任意深度in_depth处的：
    '      井斜角 单位：弧度 实型数
    '      方位角 单位：弧度 实型数
    '      垂深   单位：米 实型数
    '************************************************************************************************************************************************
    Sub cal_jx_fw_cs(ByVal in_depth As Double, ByRef ret_jxj As Double, ByRef ret_fwj As Double, ByRef ret_csh As Double)
        Dim Testwell_row As DataRow
        Dim depth() As Double '测深
        Dim jxj() As Double '井斜角
        Dim fwj() As Double '方位角
        Dim fwj_i As Double
        Dim fwj_i_1 As Double
        Dim m_jxj As Double '井斜角
        Dim i As Short
        Dim arr_length As Short '数组大小
        Dim m_fwj As Double
        Dim tol_depth As Double
        Dim dlt_depth As Double
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim ad As New System.Data.OleDb.OleDbDataAdapter
        Dim SQL_command As String

        '首先读出测井记录数据
        arr_length = Testwell_Table.Rows.Count
        If arr_length = 0 Then
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select  * from 测井数据表 where 井号='" & well_name & "' order by [井  深(m)]"
            ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
            Testwell_Table.Clear()
            ad.Fill(Testwell_Table)
            ad.Dispose()
            cn_userdb.Close()
        End If

        '重新定义数组，方便后续循环计算
        ReDim depth(arr_length)
        ReDim jxj(arr_length)
        ReDim fwj(arr_length)
        '返回值赋初值
        ret_jxj = 0.0#
        ret_fwj = 0.0#
        ret_csh = 0.0#
        depth(0) = 0.0#
        jxj(0) = 0.0#
        fwj(0) = 0.0#
        i = 0
        tol_depth = 0.0#
        '边循环边计算
        For Each Testwell_row In Testwell_Table.Select
            i = i + 1
            '获取节点测深、井斜角、方位角参数
            depth(i) = Val(Testwell_row.Item("井  深(m)").ToString)
            jxj(i) = Val(Testwell_row.Item("井斜角(°)").ToString) * 3.1415926535 / 180.0#
            fwj(i) = Val(Testwell_row.Item("方位角(°)").ToString) * 3.1415926535 / 180.0#
            '计算节点垂深
            dlt_depth = 0.5 * (depth(i) - depth(i - 1)) * (System.Math.Cos(jxj(i)) + System.Math.Cos(jxj(i - 1)))
            tol_depth = tol_depth + dlt_depth
            '如果等于节点测深，节点的井斜角、方位角、垂深即为所求，退出子程序
            If in_depth = depth(i) Then
                ret_jxj = jxj(i)
                ret_fwj = fwj(i)
                ret_csh = tol_depth
                Exit Sub
            Else
                If in_depth > depth(i - 1) And in_depth < depth(i) Then
                    '如果要求点的测深介于当前节点与前一节点之间，用内插法求得要求点的井斜角、方位角，然后计算此点垂深
                    tol_depth = tol_depth - dlt_depth
                    m_jxj = jxj(i) - (jxj(i) - jxj(i - 1)) * (depth(i) - in_depth) / (depth(i) - depth(i - 1))
                    '方位角的计算要找最近的。
                    m_fwj = fwj(i) - (fwj(i) - fwj(i - 1)) * (depth(i) - in_depth) / (depth(i) - depth(i - 1))
                    If System.Math.Abs(fwj(i) - fwj(i - 1)) > 3.1415926535 Then
                        fwj_i = fwj(i)
                        fwj_i_1 = fwj(i - 1)
                        If fwj_i < fwj_i_1 Then
                            fwj_i = fwj_i + 2 * 3.1415926535
                        Else
                            fwj_i_1 = fwj_i_1 + 2 * 3.1415926535
                        End If
                        m_fwj = fwj_i - (fwj_i - fwj_i_1) * (depth(i) - in_depth) / (depth(i) - depth(i - 1))
                        If m_fwj > (2 * 3.1415926535) Then
                            m_fwj = m_fwj - (2 * 3.1415926535)
                        End If
                    End If
                    dlt_depth = 0.5 * (in_depth - depth(i - 1)) * (System.Math.Cos(m_jxj) + System.Math.Cos(jxj(i - 1)))
                    tol_depth = tol_depth + dlt_depth
                    ret_jxj = m_jxj
                    ret_fwj = m_fwj
                    ret_csh = tol_depth
                    Exit Sub
                End If
            End If
        Next
        '如果要求点的测深大于最深节点的测深，认为此点的井斜角和方位角就等于最深节点的井斜角和方位角，即认为井保持井斜、方位不变继续“直”进而成
        If in_depth > depth(i) Then
            '用外插法求得要求点的井斜角、方位角，然后计算此点垂深，此方法不可取，因这样认为井是按照最后两点的弯曲趋势继续弯曲。
            'm_jxj = jxj(i) + (jxj(i) - jxj(i - 1)) * (in_depth - depth(i)) / (depth(i) - depth(i - 1))
            'm_fwj = fwj(i) + (fwj(i) - fwj(i - 1)) * (in_depth - depth(i)) / (depth(i) - depth(i - 1))
            m_jxj = jxj(i)
            m_fwj = fwj(i)
            dlt_depth = 0.5 * (in_depth - depth(i)) * (System.Math.Cos(m_jxj) + System.Math.Cos(jxj(i)))
            tol_depth = tol_depth + dlt_depth
            ret_jxj = m_jxj
            ret_fwj = m_fwj
            ret_csh = tol_depth
            Exit Sub
        End If
    End Sub


    '************************************************************************************************************************************************
    '                     用三次样条函数拟合井斜角和方位角，计算当前井任意深度in_depth处的井斜角、方位角、垂深子程序
    ' 程序升级记事：
    '                                                                                                           秦彦斌 2020年9月20日最后整理更新
    '   2020年9月18日开始往VS2008迁移，2020年9月20日仅代码迁移完成，未进行程序测试，若使用须测试正确性。
    '   由于此子程序需要频繁调用，改原来频繁读取井斜数据表的方法，利用ADO.NET提供的DataTable，设置全局变量Testwell_Table和Ythshcsh_Table，一次填充，
    '多次使用，避免频繁读取数据库。
    '   
    '
    '
    ' 算法来源：刘巨保 岳欠杯 等编著《石油钻采管柱力学》，石油工业出版社，ISBN:9787502186609，2011.8，P7-13
    '
    '全局变量使用：
    '   well_name        井号
    '
    '输入参数：
    '   in_depth          任意深度（这里应该是测深）
    '
    '编程思路：
    '
    '算法依据:
    '
    '
    '本子程度使用注意事项：
    '
    '子程序以传地址的方式计算出给定井任意深度in_depth处的：
    '      井斜角 单位：弧度 实型数
    '      方位角 单位：弧度 实型数
    '      垂深   单位：米 实型数
    '************************************************************************************************************************************************
    Sub cal_jx_fw_cs2(ByVal in_depth As Double, ByRef ret_jxj As Double, ByRef ret_fwj As Double, ByRef ret_csh As Double)
        Dim cn_userdb As System.Data.OleDb.OleDbConnection
        Dim dbSchema As DataTable
        Dim foundRows() As DataRow
        Dim ad As New System.Data.OleDb.OleDbDataAdapter
        Dim Testwell_row As DataRow
        Dim Ythshcsh_row As DataRow
        Dim SQL_command As String

        Dim m_fwj As Double
        Dim depth() As Double '测深
        Dim jxj() As Double '井斜角
        Dim fwj() As Double '方位角
        Dim D_M() As Double '井斜角计算时的系数M
        Dim x_m() As Double '方位角计算时的系数m
        Dim m_jxj As Double '井斜角
        Dim i As Integer
        Dim N As Integer
        Dim arr_length As Short '数组大小
        Dim tol_depth As Double
        Dim dlt_depth As Double
        Dim lk As Double

        '**********************************************************************************************************************
        '   判断有无“三次井眼样条函数参数表”，若没有，调用cal_xs_caljxfwcs()过程建立之
        '**********************************************************************************************************************
        cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn_userdb.Open()
        dbSchema = cn_userdb.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = dbSchema.Select("三次井眼样条函数参数表'")
        If foundRows.Length = 0 Then
            dbSchema.Dispose()
            cn_userdb.Close()
            Call cal_xs_caljxfwcs()
        Else
            dbSchema.Dispose()
            cn_userdb.Close()
        End If
        '**********************************************************************************************************************
        '  读出“测井数据表”
        '**********************************************************************************************************************
        arr_length = Testwell_Table.Rows.Count
        If arr_length = 0 Then
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select  * from 测井数据表 where 井号='" & well_name & "' order by [井  深(m)]"
            ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
            Testwell_Table.Clear()
            ad.Fill(Testwell_Table)
            ad.Dispose()
            cn_userdb.Close()
        End If
        arr_length = Testwell_Table.Rows.Count
        '**********************************************************************************************************************
        '  读出“三次井眼样条函数参数表”
        '**********************************************************************************************************************
        arr_length = Ythshcsh_Table.Rows.Count
        If arr_length = 0 Then
            cn_userdb = New System.Data.OleDb.OleDbConnection(use_AdoConString)
            cn_userdb.Open()
            SQL_command = "select  * from 三次井眼样条函数参数表 where 井号='" & well_name & "' order by [井  深(m)]"
            ad.SelectCommand = New OleDbCommand(SQL_command, cn_userdb)
            Ythshcsh_Table.Clear()
            ad.Fill(Ythshcsh_Table)
            ad.Dispose()
            cn_userdb.Close()
        End If
        arr_length = Testwell_Table.Rows.Count
        N = arr_length
        '**********************************************************************************************************************
        '  重新定义数组，方便后续循环计算
        '**********************************************************************************************************************
        ReDim depth(N)
        ReDim jxj(N)
        ReDim fwj(N)
        ReDim D_M(N)
        ReDim x_m(N)


        '返回值赋初值
        ret_jxj = 0.0#
        ret_fwj = 0.0#
        ret_csh = 0.0#

        depth(0) = 0.0#
        jxj(0) = 0.0#
        fwj(0) = 0.0#
        i = 0
        tol_depth = 0.0#
        For Each Testwell_row In Testwell_Table.Select
            depth(i) = Val(Testwell_row.Item("井  深(m)").ToString)
            jxj(i) = Val(Testwell_row.Item("井斜角(°)").ToString) * 3.1415926535 / 180.0#
            fwj(i) = Val(Testwell_row.Item("方位角(°)").ToString) * 3.1415926535 / 180.0#
        Next
        For Each Ythshcsh_row In Ythshcsh_Table.Select
            D_M(i) = Val(Ythshcsh_row.Item("D_M").ToString)
            x_m(i) = Val(Ythshcsh_row.Item("x_m").ToString)
        Next
        '边循环边计算
        Do
            '计算点在第一测点上面，则将第一测点数据给该计算点
            If i = 0 And in_depth < depth(i) Then
                ret_jxj = jxj(i)
                ret_fwj = fwj(i)
                ret_csh = in_depth * System.Math.Cos(jxj(i))
                Exit Sub
            End If
            '如果等于节点测深，节点的井斜角、方位角、垂深即为所求，退出子程序
            If in_depth = depth(i) Then
                ret_jxj = jxj(i)
                ret_fwj = fwj(i)
                ret_csh = tol_depth
                Exit Sub
            End If

            '计算节点前各测点垂深累加
            If i = 0 Then
                dlt_depth = depth(i) * System.Math.Cos(jxj(i))
                tol_depth = tol_depth + dlt_depth
            Else
                dlt_depth = 0.5 * (depth(i) - depth(i - 1)) * (System.Math.Cos(jxj(i)) + System.Math.Cos(jxj(i - 1)))
                tol_depth = tol_depth + dlt_depth
                If in_depth > depth(i - 1) And in_depth < depth(i) Then
                    '如果要求点的测深介于当前节点与前一节点之间，用三次井眼样条函数求得要求点的井斜角、方位角，然后计算此点垂深
                    tol_depth = tol_depth - dlt_depth
                    lk = depth(i) - depth(i - 1)
                    m_jxj = D_M(i - 1) * (depth(i) - in_depth) ^ 3 / 6.0# / lk + D_M(i) * (in_depth - depth(i - 1)) ^ 3 / 6.0# / lk + (jxj(i) / lk - D_M(i) * lk / 6) * (in_depth - depth(i - 1)) + (jxj(i - 1) / lk - D_M(i - 1) * lk / 6) * (depth(i) - in_depth)
                    m_fwj = x_m(i - 1) * (depth(i) - in_depth) ^ 3 / 6.0# / lk + x_m(i) * (in_depth - depth(i - 1)) ^ 3 / 6.0# / lk + (fwj(i) / lk - x_m(i) * lk / 6) * (in_depth - depth(i - 1)) + (fwj(i - 1) / lk - x_m(i - 1) * lk / 6) * (depth(i) - in_depth)
                    dlt_depth = 0.5 * (in_depth - depth(i - 1)) * (System.Math.Cos(m_jxj) + System.Math.Cos(jxj(i - 1)))
                    tol_depth = tol_depth + dlt_depth

                    ret_jxj = m_jxj
                    ret_fwj = m_fwj
                    ret_csh = tol_depth
                    Exit Sub
                End If
            End If
            i = i + 1
        Loop Until i >= Testwell_Table.Rows.Count Or i >= Ythshcsh_Table.Rows.Count

        '如果要求点的测深大于最深节点的测深，认为此点的井斜角和方位角就等于最深节点的井斜角和方位角，即认为井保持井斜、方位不变继续“直”进而成
        If in_depth > depth(i - 1) Then
            '用外插法求得要求点的井斜角、方位角，然后计算此点垂深，此方法不可取，因这样认为井是按照最后两点的弯曲趋势继续弯曲。
            'm_jxj = jxj(i) + (jxj(i) - jxj(i - 1)) * (in_depth - depth(i)) / (depth(i) - depth(i - 1))
            'm_fwj = fwj(i) + (fwj(i) - fwj(i - 1)) * (in_depth - depth(i)) / (depth(i) - depth(i - 1))
            m_jxj = jxj(i)
            m_fwj = fwj(i)
            dlt_depth = (in_depth - depth(i)) * System.Math.Cos(m_jxj)
            tol_depth = tol_depth + dlt_depth
            ret_jxj = m_jxj

            ret_fwj = m_fwj
            ret_csh = tol_depth
            Exit Sub
        End If
    End Sub

	'***************************************************************************************************************************
	'                     用追赶法计算三次样条函数拟合井斜角和方位角时拟合系数计算子程序
	' 算法来源：刘巨保 岳欠杯 等编著《石油钻采管柱力学》，石油工业出版社，ISBN:9787502186609，2011.8，P7-13
    '                                                                                                20190913秦彦斌最后修改
	'全局变量使用：
	'   well_name        井号
	'
	'输入参数：无，但要求“测井数据表”中有数据
	'
	'编程思路：
	'   计算前先重新生成"三次井眼样条函数参数表"
	'   打开测井记录数据
	'   重新定义数组，方便后续循环计算
	'   读取测井记录数据
	'   计算矩阵参数
	'   追赶法解三对角方程组,求M，即D_M
	'   追赶法解三对角方程组,求m，即x_m
	'   写数据到“三次井眼样条函数参数表”中
	'***************************************************************************************************************************
	Sub cal_xs_caljxfwcs()
        Dim cn As System.Data.OleDb.OleDbConnection
        Dim dr As New System.Data.OleDb.OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim tb As New DataTable
        Dim schemaTable As New DataTable
        Dim row As DataRow
        Dim foundRows() As DataRow

		Dim SQL_command As String

		Dim xuhao() As Short '序号
		Dim depth() As Double '测深
		Dim jxj() As Double '井斜角
		Dim fwj() As Double '方位角
		
		Dim D_M() As Double '井斜角计算时的系数M
		Dim x_m() As Double '方位角计算时的系数m
		
		Dim lmd() As Double '追赶法计算M、m时的系数
		Dim mu() As Double '追赶法计算M、m时的系数
		Dim D_D() As Double '井斜角计算时的系数D
		Dim x_d() As Double '方位角计算时的系数d
		
		Dim u() As Double '追赶法计算中间变量u
		Dim l() As Double '追赶法计算中间变量l
		Dim y1() As Double '追赶法计算中间变量y
		Dim y2() As Double '追赶法计算中间变量y
		
        Dim lk As Double
		Dim lk_add_1 As Double
        Dim i As Integer
        Dim N As Integer
        Dim arr_length As Integer '数组大小

        '**********************************************************************************************************************
		'计算前先重新生成"三次井眼样条函数参数表"
		'**********************************************************************************************************************
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        foundRows = schemaTable.Select("TABLE_NAME='三次井眼样条函数参数表'")
        If foundRows.Length > 0 Then
            SQL_command = "DROP TABLE 三次井眼样条函数参数表"
            cmd = New OleDbCommand(SQL_command, cn)
            cmd.ExecuteNonQuery()
        End If
        schemaTable.Dispose()
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        Call database_creat(21, 2) '三次井眼样条函数参数表
        '**********************************************************************************************************************
        '打开测井记录数据
        '**********************************************************************************************************************
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select  * from 测井数据表 where 井号='" & well_name & "' order by [井  深(m)]"
        dr.SelectCommand = New OleDbCommand(SQL_command, cn)
        dr.Fill(tb)
        cn.Close()
        arr_length = tb.Rows.Count
        N = arr_length
        If arr_length = 0 Then Exit Sub
        '**********************************************************************************************************************
        '重新定义数组，方便后续循环计算
        '**********************************************************************************************************************
        ReDim xuhao(N)
        ReDim depth(N)
        ReDim jxj(N)
        ReDim fwj(N)

        ReDim D_M(N)
        ReDim x_m(N)

        ReDim lmd(N)
        ReDim mu(N)
        ReDim D_D(N)
        ReDim x_d(N)

        ReDim u(N)
        ReDim l(N)
        ReDim y1(N)
        ReDim y2(N)
        '**********************************************************************************************************************
        '读取测井记录数据
        '**********************************************************************************************************************
        depth(0) = 0.0#
        jxj(0) = 0.0#
        fwj(0) = 0.0#
        i = 0
        For Each row In tb.Select
            xuhao(i) = Val(row.Item(0).ToString)
            depth(i) = Val(row.Item(1).ToString)
            jxj(i) = Val(row.Item(2).ToString) * 3.1415926535 / 180.0#
            fwj(i) = Val(row.Item(3).ToString) * 3.1415926535 / 180.0#
            i = i + 1
        Next
        tb.Dispose()
        dr.Dispose()

        '**********************************************************************************************************************
        '计算矩阵参数
        '**********************************************************************************************************************
        lk = depth(1) - depth(0)
        D_D(0) = 6.0# * ((jxj(1) - jxj(0)) / lk - 0) / lk
        x_d(0) = 6.0# * ((fwj(1) - fwj(0)) / lk - 0) / lk
        lmd(0) = 1
        For i = 1 To N - 1 Step 1
            lk = depth(i) - depth(i - 1)
            lk_add_1 = depth(i + 1) - depth(i)
            D_D(i) = 6.0# * ((jxj(i + 1) - jxj(i)) / lk_add_1 - (jxj(i) - jxj(i - 1)) / lk) / (lk_add_1 + lk)
            x_d(i) = 6.0# * ((fwj(i + 1) - fwj(i)) / lk_add_1 - (fwj(i) - fwj(i - 1)) / lk) / (lk_add_1 + lk)
            lmd(i) = lk_add_1 / (lk_add_1 + lk)
            mu(i) = 1 - lmd(i)
        Next i
        lk = depth(N) - depth(N - 1)
        D_D(N) = 6.0# * (0 - (jxj(N) - jxj(N - 1)) / lk) / lk
        x_d(N) = 6.0# * (0 - (fwj(N) - fwj(N - 1)) / lk) / lk
        mu(N) = 0

        '**********************************************************************************************************************
        '追赶法解三对角方程组,求M，即D_M
        '追赶法解三对角方程组,求m，即x_m
        '**********************************************************************************************************************
        u(0) = 2
        y1(0) = D_D(0)
        y2(0) = x_d(0)
        For i = 1 To N Step 1
            l(i) = mu(i) / u(i - 1)
            u(i) = 2 - l(i) * lmd(i - 1)
            y1(i) = D_D(i) - l(i) * y1(i - 1)
            y2(i) = x_d(i) - l(i) * y2(i - 1)
        Next i
        D_M(N) = y1(N) / u(N)
        x_m(N) = y2(N) / u(N)
        For i = N - 1 To 0 Step -1
            D_M(i) = (y1(i) - lmd(i) * D_M(i + 1)) / u(i)
            x_m(i) = (y2(i) - lmd(i) * x_m(i + 1)) / u(i)
        Next i
        '**********************************************************************************************************************
        '写数据到“三次井眼样条函数参数表”中
        '**********************************************************************************************************************
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        For i = 0 To N - 1 Step 1
            SQL_command = "insert into 三次井眼样条函数参数表([序   号],[井  深(m)],[D_M],[x_m],[井号]) values (" & CStr(xuhao(i)) & "," & CStr(depth(i)) & "," & CStr(D_M(i)) & "," & CStr(x_m(i)) & ",'" & well_name & "')"
            cmd = New OleDbCommand(SQL_command, cn)
            cmd.ExecuteNonQuery()
        Next i
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
    End Sub
End Module