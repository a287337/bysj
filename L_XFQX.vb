Option Strict Off
Option Explicit On
Module L_XFQX
    '*********************************************************************************************************************************************
    '绘制信封曲线。输入绘制X-Y数组，及元素个数，完成曲线初始化及绘制。若元素个数为0，则只初始化。
    '输入参数：iPlotX1：绘制信封曲线的iplotx对象,注意：初始iPlotX1中不设cannel；y_ZL()：Y轴轴力数组（纵坐标）；x_YC()：X轴压差数组（横坐标）
    '          counts：X、Y轴数据数组的个数。数组下标从0开始。要求y_ZL()和x_YC()数组对应配对，元素个数相同
    '          xTitle：X轴标题；yTitle：Y轴标题
    '输出参数：无，直接在指定对象上绘制图形
    '*********************************************************************************************************************************************
    Public Sub draw_xfqx(ByRef iPlotX1 As AxiPlotLibrary.AxiPlotX, ByVal y_ZL() As Double, ByVal x_YC() As Double, ByVal counts As Short, ByVal xTitle As String, ByVal yTitle As String)
        Dim minZL As Double, maxZL As Double, maxYC As Double
        Dim minYC As Single
        Dim i As Short
        Dim Channel_count As Integer
        Dim XAxis_count As Integer
        Dim YAxis_count As Integer

        '初始化图形
        Channel_count = 0
        XAxis_count = 0
        YAxis_count = 0
        iPlotX1.ClearAllData()
        iPlotX1.get_XAxis(XAxis_count).Title = xTitle '"压差(MPa)"
        iPlotX1.get_YAxis(YAxis_count).Title = yTitle '"轴力(kN)"
        iPlotX1.get_ToolBar(0).ShowEditButton = False
        iPlotX1.get_Channel(Channel_count).Visible = True
        iPlotX1.get_Channel(Channel_count).VisibleInLegend = False
        iPlotX1.get_Channel(Channel_count).TraceLineWidth = 2
        iPlotX1.get_XAxis(XAxis_count).Min = -100
        iPlotX1.get_XAxis(XAxis_count).Span = 200
        iPlotX1.get_XAxis(XAxis_count).DesiredStart = iPlotX1.get_XAxis(XAxis_count).Min
        iPlotX1.get_XAxis(XAxis_count).DesiredIncrement = System.Math.Round((iPlotX1.get_XAxis(XAxis_count).Span) / 5, 0)
        iPlotX1.get_YAxis(YAxis_count).Min = -100
        iPlotX1.get_YAxis(YAxis_count).Span = 200
        iPlotX1.get_YAxis(YAxis_count).DesiredStart = iPlotX1.get_YAxis(YAxis_count).Min
        iPlotX1.get_YAxis(YAxis_count).DesiredIncrement = System.Math.Round((iPlotX1.get_YAxis(YAxis_count).Span) / 5, 0)
        If counts = 0 Then Exit Sub
        i = 0
        maxZL = y_ZL(i)
        minZL = y_ZL(i)
        maxYC = x_YC(i)
        minYC = x_YC(i)
        For i = 0 To counts - 1
            If y_ZL(i) > maxZL Then maxZL = y_ZL(i)
            If y_ZL(i) < minZL Then minZL = y_ZL(i)
            If x_YC(i) > maxYC Then maxYC = x_YC(i)
            If x_YC(i) < minYC Then minYC = x_YC(i)
        Next i
        '计算X、Y轴最大、最小刻度值=坐标点最大最小值向外延伸8%
        iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(System.Math.Round(minYC, 2) - System.Math.Sign(minYC) * System.Math.Round(minYC * 0.08, 2), 2)
        iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(maxYC, 2) + System.Math.Sign(maxYC) * System.Math.Round(maxYC * 0.08, 2)
        If iPlotX1.get_XAxis(XAxis_count).Span > 0 And iPlotX1.get_XAxis(XAxis_count).Min < 0 Then
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span - iPlotX1.get_XAxis(XAxis_count).Min, 2)
        Else
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span - System.Math.Sign(iPlotX1.get_XAxis(XAxis_count).Min) * iPlotX1.get_XAxis(XAxis_count).Min, 2)
        End If
        iPlotX1.get_YAxis(YAxis_count).Min = System.Math.Round(minZL, 2) - System.Math.Sign(minZL) * System.Math.Round(minZL * 0.08, 2)
        iPlotX1.get_YAxis(YAxis_count).Span = System.Math.Round(maxZL, 2) + System.Math.Sign(maxZL) * System.Math.Round(maxZL * 0.08, 2)
        If iPlotX1.get_YAxis(YAxis_count).Span > 0 And iPlotX1.get_YAxis(YAxis_count).Min < 0 Then
            iPlotX1.get_YAxis(YAxis_count).Span = iPlotX1.get_YAxis(YAxis_count).Span - iPlotX1.get_YAxis(YAxis_count).Min
        Else
            iPlotX1.get_YAxis(YAxis_count).Span = iPlotX1.get_YAxis(YAxis_count).Span - System.Math.Sign(iPlotX1.get_YAxis(YAxis_count).Min) * iPlotX1.get_YAxis(YAxis_count).Min
        End If
        '将坐标原点和坐标长度转为5的整倍数()
        If (System.Math.Abs(iPlotX1.get_XAxis(XAxis_count).Min) Mod 5 < 2.5) Then
            iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(Int(iPlotX1.get_XAxis(XAxis_count).Min / 10) * 10, 0)
        Else
            iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(Int(iPlotX1.get_XAxis(XAxis_count).Min / 10) * 10 + 5, 0)
        End If
        If (System.Math.Abs(iPlotX1.get_XAxis(XAxis_count).Span) Mod 5 < 2.5) Then
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span / 10, 0) * 10
        Else
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span / 10, 0) * 10 + 5, 0)
        End If
        iPlotX1.get_XAxis(XAxis_count).DesiredStart = iPlotX1.get_XAxis(XAxis_count).Min
        iPlotX1.get_XAxis(XAxis_count).DesiredIncrement = System.Math.Round((iPlotX1.get_XAxis(XAxis_count).Span) / 10, 0)
        iPlotX1.get_YAxis(YAxis_count).DesiredStart = iPlotX1.get_YAxis(YAxis_count).Min
        iPlotX1.get_YAxis(YAxis_count).DesiredIncrement = System.Math.Round((iPlotX1.get_YAxis(YAxis_count).Span) / 50, 0)
        For i = 0 To counts - 1
            iPlotX1.get_Channel(Channel_count).AddXY(x_YC(i), y_ZL(i))
        Next i
        '占据坐标刻度最大、最小值的点
        Channel_count = iPlotX1.AddChannel()
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(iPlotX1.get_XAxis(XAxis_count).Min, System.Math.Round(iPlotX1.get_YAxis(YAxis_count).Min + iPlotX1.get_YAxis(YAxis_count).Span, 0))
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Min + iPlotX1.get_XAxis(XAxis_count).Span, 0), iPlotX1.get_YAxis(YAxis_count).Min)
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Color = &HFFFFFF
        'iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Color = GetUint32color(iPlotX1.ChannelCount)
    End Sub
    '*********************************************************************************************************************************************
    '信封曲线上点一个点
    '*********************************************************************************************************************************************
    Public Sub drawDot(ByRef iPlotX1 As AxiPlotLibrary.AxiPlotX, ByVal x As Single, ByVal y As Single, ByVal dotTitle As String)
        '任意显示的点
        Dim dlt_x As Single
        Dim dlt_y As Single
        iPlotX1.AddChannel()
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Color = GetUint32color(iPlotX1.ChannelCount)
        'If System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Color)).Equals(System.Drawing.Color.White) Then
        '    iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Color = System.Convert.ToUInt32(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
        'End If
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).TraceLineStyle = iPlotLibrary.TxiPlotLineStyle.iplsSolid
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).TraceLineWidth = 8
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).Visible = True
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).VisibleInLegend = True
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).TitleText = dotTitle
        dlt_x = iPlotX1.get_XAxis(0).Span / 300000
        dlt_y = iPlotX1.get_YAxis(0).Span / 100000
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(x - dlt_x, y)
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(x, y + dlt_y)
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(x + dlt_x, y)
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(x, y - dlt_y)
        iPlotX1.get_Channel(iPlotX1.ChannelCount - 1).AddXY(x - dlt_x, y)
    End Sub

    '*********************************************************************************************************************************************
    '                                         在指定的AxiPlotX对象的指定Channel上点一个点
    '输入参数：iPlotX1：绘制信封曲线的iplotx对象
    '          Channel_no：绘图Channel号
    '          y：纵坐标
    '          x：横坐标
    '输出参数：无，直接在指定对象上绘制图形
    '*********************************************************************************************************************************************
    Public Sub drawDot_at_ch(ByRef iPlotX1 As AxiPlotLibrary.AxiPlotX, ByVal Channel_no As Short, ByVal x As Single, ByVal y As Single)
        '任意显示的点
        Dim dlt_x As Single
        Dim dlt_y As Single
        iPlotX1.get_Channel(Channel_no).TraceLineStyle = iPlotLibrary.TxiPlotLineStyle.iplsSolid
        iPlotX1.get_Channel(Channel_no).TraceLineWidth = 8
        iPlotX1.get_Channel(Channel_no).Visible = True
        dlt_x = iPlotX1.get_XAxis(0).Span / 300000
        dlt_y = iPlotX1.get_YAxis(0).Span / 100000
        iPlotX1.get_Channel(Channel_no).AddXY(x - dlt_x, y)
        iPlotX1.get_Channel(Channel_no).AddXY(x, y + dlt_y)
        iPlotX1.get_Channel(Channel_no).AddXY(x + dlt_x, y)
        iPlotX1.get_Channel(Channel_no).AddXY(x, y - dlt_y)
        iPlotX1.get_Channel(Channel_no).AddXY(x - dlt_x, y)
    End Sub

    '*********************************************************************************************************************************************
    '                                         在指定的AxiPlotX对象的指定Channel上绘制曲线
    '
    '输入绘制X-Y数组，及元素个数，完成曲线初始化及绘制。若元素个数为0，则只初始化。
    '输入参数：iPlotX1：绘制信封曲线的iplotx对象
    '          Channel_no：绘图Channel号
    '          y_ZZB()：Y轴数组（纵坐标）
    '          x_HZB()：X轴数组（横坐标）
    '          counts：X、Y轴数据数组的个数。数组下标从0开始。要求y_ZZB()和x_HZB()数组对应配对，元素个数相同

    '输出参数：无，直接在指定对象上绘制图形
    '*********************************************************************************************************************************************
    Public Sub draw_blqx_at_ch(ByRef iPlotX1 As AxiPlotLibrary.AxiPlotX, ByVal Channel_no As Short, ByVal x_HZB() As Double, ByVal y_ZZB() As Double, ByVal counts As Short)
        Dim i As Short
        Dim XAxis_count As Integer
        Dim YAxis_count As Integer
        i = 0
        XAxis_count = 0
        YAxis_count = 0
        For i = 0 To counts - 1
            iPlotX1.get_Channel(Channel_no).AddXY(x_HZB(i), y_ZZB(i))
        Next i
    End Sub
    '*********************************************************************************************************************************************
    '                                         在指定的AxiPlotX对象上的0,1号Channel画管柱安全包络线
    '输入参数：iPlotX1：绘制信封曲线的iplotx对象
    '          burst_lim：抗内压强度，MPa
    '          tension_lim：抗拉强度，kN
    '          collapse_lim：抗挤强度，MPa
    '          compression_lim：抗压挤强度，kN
    '          xgm_s：屈服强度，MPa
    '          waijing：油管外径，mm
    '          neijing：油管内径，mm
    '          drawmode:  绘图开关，整形变量，等于0按真实轴力画，等于1按等效轴力画  

    '输出参数：无，直接在指定对象上绘制图形
    '*********************************************************************************************************************************************
    Public Sub draw_qdjx_blqx(ByRef iPlotX1 As AxiPlotLibrary.AxiPlotX, ByVal burst_lim As Double, ByVal tension_lim As Double, ByVal collapse_lim As Double, ByVal compression_lim As Double, ByVal xgm_s As Double, ByVal waijing As Double, ByVal neijing As Double, ByVal drawmode As Integer)
        Dim x_hzb() As Double
        Dim y_zzb() As Double
        Dim i As Integer
        Dim point_count As Integer
        Dim A_tubing As Double
        Dim dlt_Fz As Double
        Dim Fz As Double
        Dim xgm_st As Double         '根据管体屈服强度（kN）算出的屈服强度（MPa）
        Dim xgm_z As Double          '非弯曲引起的轴向应力的分量
        Dim ped As Double            '在轴向应力和内压作用下的组合加载当量等级、等效屈服强度
        Dim pi As Double             '管内压力
        Dim temp_A As Double
        Dim temp_B As Double
        Dim temp_C As Double
        Dim temp_D As Double
        Dim temp_F As Double
        Dim temp_L As Double
        Dim ALF As Double
        Dim PAI As Double
        Dim minZZB As Double
        Dim maxZZB As Double
        Dim maxHZB As Double
        Dim minHZB As Double
        Dim XAxis_count As Integer
        Dim YAxis_count As Integer

        PAI = 3.1415926535897931
        XAxis_count = 0
        YAxis_count = 0
        '截面面积
        A_tubing = 0.25 * (waijing * waijing - neijing * neijing)
        'pi = Val(BindingSource9.Current("管内压力MPa").ToString)
        pi = collapse_lim
        '***************************************************************************************************************
        '1.算画单双轴强度极限
        ' 公式来源：
        '    (1) wellCAT帮助文件：2015 API TR 5C3 Collapse Formula Update WELLCAT Workflow
        '    (2) 00 高难度复杂井完井（试油）油套管柱力学分析要点及若干工程问题简析（2011年5月11日西安完井会议）.ppt
        ' 说明：
        '   （1）当（1）中pi取为0时，（1）、（2）中公式一致，对每一工况（如此某点的内压不同），pi不一定为零，画出的图差别不大。
        '   （2）公式中乘以屈服极限不行
        '   （3）20220316按（2）式调通，20220317（1）、（2）式代码合并
        '    (3) 20251011,按GBT20657-2022石油天然气工业套管油管钻杆和用作套管或油管的管线管性能公式及计算再次检查修正计算，
        '***************************************************************************************************************
        xgm_st = 1000 * tension_lim / A_tubing
        Fz = 0.0#
        dlt_Fz = 100
        i = 0
        Do
            xgm_z = 1000 * Fz / A_tubing
            '乘以屈服极限不行：ped = -1 * xgm_st * (Math.Sqrt(1 - 0.75 * ((xgm_z + pi) / xgm_st) * ((xgm_z + pi) / xgm_st)) - 0.5 * ((xgm_z + pi) / xgm_st))
            'ped = collapse_lim * (Math.Sqrt(1 - 0.75 * ((xgm_z + pi) / xgm_st) * ((xgm_z + pi) / xgm_st)) - 0.5 * ((xgm_z + pi) / xgm_st))
            If Fz = 0 Then
                ped = collapse_lim
            End If
            Fz = Fz + dlt_Fz
            i = i + 1
            'Loop While ((xgm_z + pi) >= 0) And Fz < tension_lim
        Loop While Fz < tension_lim
        point_count = i
        Fz = 0.0#
        ReDim x_hzb(point_count + 5)
        ReDim y_zzb(point_count + 5)
        For i = 0 To point_count - 1 Step 1
            x_hzb(i) = Fz
            xgm_z = 1000 * Fz / A_tubing
            '乘以屈服极限不行：y_zzb(i) = -1 * xgm_st * (Math.Sqrt(1 - 0.75 * ((xgm_z + pi) / xgm_st) * ((xgm_z + pi) / xgm_st)) - 0.5 * ((xgm_z + pi) / xgm_st))
            '20251011,按GBT20657-2022石油天然气工业套管油管钻杆和用作套管或油管的管线管性能公式及计算再次检查修正计算
            If (xgm_z + pi) >= 0 Then
                y_zzb(i) = collapse_lim * (Math.Sqrt(1 - 0.75 * ((xgm_z + pi) / xgm_st) * ((xgm_z + pi) / xgm_st)) - 0.5 * ((xgm_z + pi) / xgm_st))
            Else
                y_zzb(i) = collapse_lim
            End If
            If Fz = 0 Then
                y_zzb(i) = collapse_lim
            End If
            Fz = Fz + dlt_Fz
        Next i
        Fz = tension_lim
        x_hzb(i) = Fz
        xgm_z = 1000 * Fz / A_tubing
        y_zzb(i) = collapse_lim * (Math.Sqrt(1 - 0.75 * ((xgm_z + pi) / xgm_st) * ((xgm_z + pi) / xgm_st)) - 0.5 * ((xgm_z + pi) / xgm_st))
        If Fz = 0 Then
            y_zzb(i) = collapse_lim
        End If
        i = i + 1
        x_hzb(i) = tension_lim
        y_zzb(i) = burst_lim
        i = i + 1
        x_hzb(i) = compression_lim
        y_zzb(i) = burst_lim
        i = i + 1
        x_hzb(i) = compression_lim
        y_zzb(i) = collapse_lim
        i = i + 1
        x_hzb(i) = 0
        y_zzb(i) = collapse_lim
        i = i + 1
        x_hzb(i) = 0
        y_zzb(i) = collapse_lim
        maxZZB = y_zzb(0)
        minZZB = y_zzb(0)
        maxHZB = x_hzb(0)
        minHZB = x_hzb(0)
        For i = 0 To point_count + 4
            If y_zzb(i) > maxZZB Then maxZZB = y_zzb(i)
            If y_zzb(i) < minZZB Then minZZB = y_zzb(i)
            If x_hzb(i) > maxHZB Then maxHZB = x_hzb(i)
            If x_hzb(i) < minHZB Then minHZB = x_hzb(i)
        Next i
        iPlotX1.get_Channel(0).Clear()
        Call draw_blqx_at_ch(iPlotX1, 0, x_hzb, y_zzb, point_count + 5)
        '*********************************************************************************************
        '2.算画三轴强度极限
        '*********************************************************************************************
        ALF = 0
        ReDim x_hzb(3600)
        ReDim y_zzb(3600)
        i = 0
        Do
            '用真实轴力算
            temp_A = 1000.0 * Math.Cos(ALF * PAI / 180) / (0.25 * PAI * (waijing * waijing - neijing * neijing))
            temp_F = temp_A
            If drawmode = 1 Then
                '用等效轴力算
                temp_A = 4000.0 * Math.Cos(ALF * PAI / 180) / (PAI * (waijing * waijing - neijing * neijing)) - neijing * neijing * Math.Sin(ALF * PAI / 180) / (waijing * waijing - neijing * neijing)
                temp_F = 4000.0 * Math.Cos(ALF * PAI / 180) / (PAI * (waijing * waijing - neijing * neijing)) - waijing * waijing * Math.Sin(ALF * PAI / 180) / (waijing * waijing - neijing * neijing)
            End If
            
            temp_B = Math.Sin(ALF * PAI / 180) * (waijing * waijing + neijing * neijing) / (waijing * waijing - neijing * neijing)
            temp_C = -1 * Math.Sin(ALF * PAI / 180)
            temp_D = 2 * waijing * waijing * Math.Sin(ALF * PAI / 180) / (waijing * waijing - neijing * neijing)
            If ALF <= 180 Then
                temp_L = Math.Sqrt(2 * xgm_s * xgm_s / ((temp_A - temp_B) * (temp_A - temp_B) + (temp_B - temp_C) * (temp_B - temp_C) + (temp_C - temp_A) * (temp_C - temp_A)))
            Else
                temp_L = Math.Sqrt(2 * xgm_s * xgm_s / ((temp_F - temp_D) * (temp_F - temp_D) + temp_D * temp_D + temp_F * temp_F))
            End If
            x_hzb(i) = temp_L * Math.Cos(ALF * PAI / 180)
            y_zzb(i) = temp_L * Math.Sin(ALF * PAI / 180)
            ALF = ALF + 0.1
            i = i + 1
        Loop While ALF <= 360.0
        x_hzb(i) = x_hzb(0)
        y_zzb(i) = y_zzb(0)
        For i = 0 To 3600
            If y_zzb(i) > maxZZB Then maxZZB = y_zzb(i)
            If y_zzb(i) < minZZB Then minZZB = y_zzb(i)
            If x_hzb(i) > maxHZB Then maxHZB = x_hzb(i)
            If x_hzb(i) < minHZB Then minHZB = x_hzb(i)
        Next i
        iPlotX1.get_Channel(1).Clear()
        Call draw_blqx_at_ch(iPlotX1, 1, x_hzb, y_zzb, 3601)

        '计算X、Y轴最大、最小刻度值=坐标点最大最小值向外延伸8%
        iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(System.Math.Round(minHZB, 3) - System.Math.Sign(minHZB) * System.Math.Round(minHZB * 0.08, 3), 3)
        iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(maxHZB, 3) + System.Math.Sign(maxHZB) * System.Math.Round(maxHZB * 0.08, 3)
        If iPlotX1.get_XAxis(XAxis_count).Span > 0 And iPlotX1.get_XAxis(XAxis_count).Min < 0 Then
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span - iPlotX1.get_XAxis(XAxis_count).Min, 3)
        Else
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span - System.Math.Sign(iPlotX1.get_XAxis(XAxis_count).Min) * iPlotX1.get_XAxis(XAxis_count).Min, 3)
        End If
        iPlotX1.get_YAxis(YAxis_count).Min = System.Math.Round(minZZB, 3) - System.Math.Sign(minZZB) * System.Math.Round(minZZB * 0.08, 3)
        iPlotX1.get_YAxis(YAxis_count).Span = System.Math.Round(maxZZB, 3) + System.Math.Sign(maxZZB) * System.Math.Round(maxZZB * 0.08, 3)
        If iPlotX1.get_YAxis(YAxis_count).Span > 0 And iPlotX1.get_YAxis(YAxis_count).Min < 0 Then
            iPlotX1.get_YAxis(YAxis_count).Span = iPlotX1.get_YAxis(YAxis_count).Span - iPlotX1.get_YAxis(YAxis_count).Min
        Else
            iPlotX1.get_YAxis(YAxis_count).Span = iPlotX1.get_YAxis(YAxis_count).Span - System.Math.Sign(iPlotX1.get_YAxis(YAxis_count).Min) * iPlotX1.get_YAxis(YAxis_count).Min
        End If
        '将坐标原点和坐标长度转为5的整倍数()
        If (System.Math.Abs(iPlotX1.get_XAxis(XAxis_count).Min) Mod 5 < 2.5) Then
            iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(Int(iPlotX1.get_XAxis(XAxis_count).Min / 10) * 10, 0)
        Else
            iPlotX1.get_XAxis(XAxis_count).Min = System.Math.Round(Int(iPlotX1.get_XAxis(XAxis_count).Min / 10) * 10 + 5, 0)
        End If
        If (System.Math.Abs(iPlotX1.get_XAxis(XAxis_count).Span) Mod 5 < 2.5) Then
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span / 10, 0) * 10
        Else
            iPlotX1.get_XAxis(XAxis_count).Span = System.Math.Round(System.Math.Round(iPlotX1.get_XAxis(XAxis_count).Span / 10, 0) * 10 + 5, 0)
        End If
    End Sub
End Module