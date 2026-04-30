Option Strict Off
Option Explicit On
Imports System.Drawing
Imports System.Data.OleDb
Module M_subWellStruction
    '***************************************************************************************************************************************************
    '                                                            井身结构及管柱结构绘制模块
    ' 程序升级记事：
    '                                                                                           秦彦斌 2019年10月30日最后整理
    '    因为 "当应用程序从 Visual Basic 6.0 升级到 Visual Basic 2008 时，图形方法并不升级，并会在代码中插入警告。由于 GDI 和 GDI+ 之间的巨大差异，
    '现有的所有图形代码都需要重写。（摘自VS2008帮助:ms-help://MS.VSCC.v90/MS.MSDNQTR.v90.chs/dv_vbvers/html/24cd2d55-ebf1-42d6-b755-00e9001f1cb8.htm）"                               
    '2019年10月30日，重写井身结构及管柱结构绘制模块。
    '
    ' 输入参数：
    ' picControl            PictureBox对象
    ' opt_select            开关变量，opt_select=1时仅画井身结构，opt_select=2时画井身结构+管柱结构
    '
    '***************************************************************************************************************************************************
    Public Sub drawWellStruction(ByRef picControl As PictureBox, ByVal opt_select As Short)
        Dim SQL_command As String
        Dim EXECOleDbCommand As OleDbCommand
        Dim RECreader As OleDbDataReader
        Dim RECreader1 As OleDbDataReader
        Dim RECreader2 As OleDbDataReader
        Dim cn As System.Data.OleDb.OleDbConnection
        Dim mycolor As Color

        Dim gr As Graphics
        Dim drawFont As New System.Drawing.Font("Arial", 8)
        Dim mpen As New Pen(Color.Black)

        Dim casinglayer_max As Short
        Dim i As Short
        Dim j As Short
        Dim casinglayer_min As Short
        Dim casingDo_max As Double
        Dim casing_start1 As Double
        Dim casing_end1 As Double
        Dim drill_start As Double
        Dim drill_end As Double
        Dim casing_Do1 As Double
        Dim casing_t1 As Double
        Dim casing_L As Double
        Dim drill_D As Double
        Dim shnfsh As Double
        Dim zzjsh As Double
        Dim max_depth As Double
        Dim min_zhijing As Double
        Dim depth As Double
        Dim y_scale As Double
        Dim x_scale As Double
        Dim x1 As Integer
        Dim x2 As Integer
        Dim y1 As Integer
        Dim y2 As Integer
        Dim rctw As Integer
        Dim rcth As Integer
        Dim c1 As Integer
        Dim c2 As Integer
        Dim c3 As Integer

        gr = picControl.CreateGraphics
        gr.Clear(Color.White)
        '****************************************************************************************************************************************************
        '画中心线
        '****************************************************************************************************************************************************
        x1 = picControl.Size.Width / 2
        x2 = x1
        y1 = 0
        y2 = picControl.Size.Height
        mpen.Width = 1
        mpen.Color = Color.Black
        mpen.DashStyle = Drawing2D.DashStyle.DashDot
        ' Create a custom dash pattern.
        mpen.DashPattern = New Single() {15.0F, 2.0F, 3.0F, 2.0F}
        gr.DrawLine(mpen, x1, y1, x2, y2)
        '****************************************************************************************************************************************************
        ' 从油气井表中取得完钻井深
        '****************************************************************************************************************************************************
        max_depth = 0.0
        min_zhijing = 0.0
        '用ADO.NET给井基本数据赋值
        cn = New System.Data.OleDb.OleDbConnection(use_AdoConString)
        cn.Open()
        SQL_command = "select 地理位置,构造位置,井别,设计井深m,完钻井深m,完钻层位,完钻钻头尺寸mm from 油气井表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader.Read Then
            max_depth = Val(RECreader.Item("完钻井深m").ToString)
            min_zhijing = Val(RECreader.Item("完钻钻头尺寸mm").ToString)
        End If
        RECreader.Close()
        SQL_command = "select * from 套管数据表 where 井号='" & well_name & "'"
        EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
        RECreader1 = EXECOleDbCommand.ExecuteReader()
        EXECOleDbCommand.Dispose()
        If RECreader1.Read Then
            SQL_command = "select distinct min(套管层数) as 最小套管层数, max(套管层数) as 最大套管层数, max(套管外径mm) as 最大套管套管外径, max(套管下深m) as 最大下深 from 套管数据表 where 井号='" & well_name & "'"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            RECreader2 = EXECOleDbCommand.ExecuteReader()
            EXECOleDbCommand.Dispose()
            If RECreader2.Read Then
                casinglayer_max = Val(RECreader2.Item("最大套管层数").ToString)
                casinglayer_min = Val(RECreader2.Item("最小套管层数").ToString)
                casingDo_max = Val(RECreader2.Item("最大套管套管外径").ToString)
                y_scale = (picControl.Size.Height / max_depth) * 0.98
                x_scale = (picControl.Size.Width / casingDo_max) * 0.6
                drill_start = 0.0
                drill_end = 0.0
                '为了文字不被覆盖，先画图
                For i = casinglayer_min To casinglayer_max
                    SQL_command = "select  * from  套管数据表 where 井号='" & well_name & "'and 套管层数 =" & i & " order by 套管段数 "
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    j = 1
                    zzjsh = 0.0
                    While RECreader.Read() '读取数据
                        casing_start1 = Val(RECreader.Item("悬挂深度m").ToString)
                        casing_end1 = Val(RECreader.Item("套管下深m").ToString)
                        casing_Do1 = Val(RECreader.Item("套管外径mm").ToString)
                        casing_t1 = Val(RECreader.Item("套管壁厚mm").ToString)
                        drill_D = IIf(RECreader.Item("钻头尺寸mm").ToString = "", casing_Do1 + 25.4, Val(RECreader.Item("钻头尺寸mm").ToString))
                        If drill_D = 0.0 Then
                            drill_D = casing_Do1 + 25.4
                        End If
                        shnfsh = Val(RECreader.Item("水泥返深m").ToString)
                        If j = 1 Then
                            zzjsh = Val(RECreader.Item("完钻深度m").ToString)
                            drill_end = zzjsh
                        End If
                        Do
                            c1 = Val(Int((255 - 0 + 1) * Rnd(i) + 0))
                            c2 = Val(Int((255 - 0 + 1) * Rnd(j + i) + 0))
                            c3 = Val(Int((255 - 0 + 1) * Rnd(i * j) + 0))
                        Loop While c1 > 100 And c2 > 100 And c3 > 100
                        mycolor = Color.FromArgb(255, c2, c3, c1)
                        mpen.Width = casing_t1 * x_scale
                        mpen.Color = mycolor
                        mpen.DashStyle = Drawing2D.DashStyle.Solid
                        '画右边水泥返深
                        x1 = picControl.Size.Width / 2 - x_scale * drill_D / 2 - 0.3 * mpen.Width
                        y1 = shnfsh * y_scale
                        rctw = x_scale * (drill_D - casing_Do1) / 2
                        rcth = y_scale * (casing_end1 - shnfsh)
                        gr.FillRectangle(Brushes.LightGray, x1, y1, rctw, rcth)
                        '画左边水泥返深
                        x1 = picControl.Size.Width / 2 + x_scale * casing_Do1 / 2 + 0.3 * mpen.Width
                        y1 = shnfsh * y_scale
                        rctw = x_scale * (drill_D - casing_Do1) / 2
                        rcth = y_scale * (casing_end1 - shnfsh)
                        gr.FillRectangle(Brushes.LightGray, x1, y1, rctw, rcth)
                        '画套管右边竖线
                        x1 = picControl.Size.Width / 2 - x_scale * casing_Do1 / 2
                        x2 = x1
                        y1 = casing_start1 * y_scale
                        y2 = casing_end1 * y_scale
                        gr.DrawLine(mpen, x1, y1, x2, y2)
                        '画套管左边竖线
                        x1 = picControl.Size.Width / 2 + x_scale * casing_Do1 / 2
                        x2 = x1
                        y1 = casing_start1 * y_scale
                        y2 = casing_end1 * y_scale
                        gr.DrawLine(mpen, x1, y1, x2, y2)
                        j = j + 1
                    End While
                    RECreader.Close()
                    '画右边套管鞋
                    mpen.LineJoin = Drawing2D.LineJoin.Round
                    mpen.Width = 2
                    x1 = picControl.Size.Width / 2 - x_scale * casing_Do1 / 2
                    x2 = x1 - 6
                    y1 = casing_end1 * y_scale
                    y2 = y1
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    y1 = y2 - 5
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    '画左边套管鞋
                    x1 = picControl.Size.Width / 2 + x_scale * casing_Do1 / 2
                    x2 = x1 + 6
                    y1 = casing_end1 * y_scale
                    y2 = y1
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    y1 = y2 - 5
                    gr.DrawLine(mpen, x1, y1, x2, y2)

                    '画井眼右边竖线
                    mpen.Width = 0.5
                    mpen.Color = Color.LightGray
                    x1 = picControl.Size.Width / 2 - x_scale * drill_D / 2
                    x2 = x1
                    y1 = drill_start * y_scale
                    y2 = drill_end * y_scale
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    '画井眼左边竖线
                    x1 = picControl.Size.Width / 2 + x_scale * drill_D / 2
                    x2 = x1
                    y1 = drill_start * y_scale
                    y2 = drill_end * y_scale
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    drill_start = drill_end
                Next i
                If drill_end < max_depth Then
                    x1 = picControl.Size.Width / 2 - x_scale * min_zhijing / 2
                    x2 = x1
                    y1 = drill_start * y_scale
                    y2 = max_depth * y_scale
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                    '画井眼左边竖线
                    x1 = picControl.Size.Width / 2 + x_scale * min_zhijing / 2
                    x2 = x1
                    y1 = drill_start * y_scale
                    y2 = max_depth * y_scale
                    gr.DrawLine(mpen, x1, y1, x2, y2)
                End If
                '为了文字不被覆盖，先画图，然后写文字
                For i = casinglayer_min To casinglayer_max
                    SQL_command = "select  * from  套管数据表 where 井号='" & well_name & "'and 套管层数 =" & i & " order by 套管段数 "
                    EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
                    RECreader = EXECOleDbCommand.ExecuteReader()
                    EXECOleDbCommand.Dispose()
                    j = 1
                    While RECreader.Read() '读取数据
                        casing_start1 = Val(RECreader.Item("悬挂深度m").ToString)
                        casing_end1 = Val(RECreader.Item("套管下深m").ToString)
                        casing_Do1 = Val(RECreader.Item("套管外径mm").ToString)
                        casing_t1 = Val(RECreader.Item("套管壁厚mm").ToString)
                        drill_D = Val(RECreader.Item("钻头尺寸mm").ToString)
                        shnfsh = Val(RECreader.Item("水泥返深m").ToString)
                        '在右边写悬挂深度
                        If casing_start1 <> 0 Then
                            x1 = picControl.Size.Width / 2 - 1.6 * x_scale * casingDo_max / 2
                            y1 = Math.Abs(casing_start1 - 100) * y_scale
                            gr.DrawString(Trim(CStr(casing_start1)) & "m", drawFont, Brushes.DarkBlue, x1, y1)
                        End If
                        '在左边写下入深度
                        x1 = picControl.Size.Width / 2 + 0.85 * x_scale * casingDo_max / 2
                        y1 = Math.Abs(casing_end1 - 100) * y_scale
                        gr.DrawString(Trim(CStr(casing_end1)) & "m", drawFont, Brushes.DarkBlue, x1, y1)
                        j = j + 1
                    End While
                    RECreader.Close()
                Next i
            End If
            RECreader2.Close()
        End If
        RECreader1.Close()
        '****************************************************************************************************************************************************
        ' 绘制井底横线及标注
        '****************************************************************************************************************************************************
        mycolor = Color.FromArgb(255, 0, 0, 0)
        mpen.Width = 2
        mpen.Color = mycolor
        mpen.DashStyle = Drawing2D.DashStyle.Solid
        casing_end1 = max_depth
        x1 = picControl.Size.Width / 2 - x_scale * min_zhijing / 2
        x2 = picControl.Size.Width / 2 + x_scale * min_zhijing / 2
        y1 = casing_end1 * y_scale
        y2 = y1
        gr.DrawLine(mpen, x1, y1, x2, y2)
        x1 = picControl.Size.Width / 2 - 1.6 * x_scale * casingDo_max / 2
        y1 = casing_end1 * y_scale - 8
        gr.DrawString(Trim(CStr(casing_end1)) & "m", drawFont, Brushes.DarkBlue, x1, y1)

        If opt_select = 2 Then
            '****************************************************************************************************************************************************
            ' 绘制当前工况管柱结构图
            '****************************************************************************************************************************************************
            SQL_command = "select * from 管柱数据表 where 井号='" & well_name & "' and 作业名称='" & zuoye_name & "' order by 元件序号"
            EXECOleDbCommand = New OleDbCommand(SQL_command, cn)
            RECreader = EXECOleDbCommand.ExecuteReader()
            EXECOleDbCommand.Dispose()
            depth = 0.0#
            j = 1
            While RECreader.Read() '读取数据
                casing_start1 = depth
                casing_L = Val(RECreader.Item("元件长度m").ToString)
                casing_end1 = casing_L + depth
                casing_Do1 = Val(RECreader.Item("元件外径mm").ToString)
                casing_t1 = 0.5 * (Val(RECreader.Item("元件外径mm").ToString) - Val(RECreader.Item("元件内径mm").ToString))
                Do
                    c1 = Val(Int((255 - 0 + 1) * Rnd(i) + 0))
                    c2 = Val(Int((255 - 0 + 1) * Rnd(j + i) + 0))
                    c3 = Val(Int((255 - 0 + 1) * Rnd(i * j) + 0))
                Loop While c1 > 100 And c2 > 100 And c3 > 100
                mycolor = Color.FromArgb(255, c2, c3, c1)
                mpen.Width = casing_t1 * x_scale
                mpen.Color = mycolor
                mpen.DashStyle = Drawing2D.DashStyle.Solid
                '画右边竖线
                x1 = picControl.Size.Width / 2 - x_scale * casing_Do1 / 2
                x2 = x1
                y1 = casing_start1 * y_scale
                y2 = casing_end1 * y_scale
                gr.DrawLine(mpen, x1, y1, x2, y2)
                '画左边竖线
                x1 = picControl.Size.Width / 2 + x_scale * casing_Do1 / 2
                x2 = x1
                y1 = casing_start1 * y_scale
                y2 = casing_end1 * y_scale
                gr.DrawLine(mpen, x1, y1, x2, y2)
                '画水平线
                mpen.Width = 2
                x1 = picControl.Size.Width / 2 - x_scale * casing_Do1 / 2
                x2 = picControl.Size.Width / 2 + x_scale * casing_Do1 / 2
                y1 = casing_end1 * y_scale
                y2 = y1
                gr.DrawLine(mpen, x1, y1, x2, y2)
                depth = depth + casing_L
            End While
            RECreader.Close()
        End If
        cn.Close()
        drawFont.Dispose()
        mpen.Dispose()
    End Sub
End Module