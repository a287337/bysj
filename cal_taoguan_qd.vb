Option Strict Off
Option Explicit On
Module M_cal_taoguan_qd
	'************************************************************************************************************************************************************************
	'                                            管体最大三轴等效应力计算
	'
    ' 算法来源：GB/T20657-2011/ISO/TR10400:2007-石油天然气工业套管、油管、钻杆和用作套管或油管的管线管性能公式和计算:P12
	' 说明：
	'    (1) 考虑到井下管柱在弯曲井眼中受弯矩，在螺旋弯曲时存在螺旋弯矩，故参数传递设计成传入弯矩以便正确计算弯曲应力。
	'    (2) 井下管柱还会承受由套管或井壁作用的侧向力，该力会在管柱中产生剪切应力，故也需传入参与计算。
	' 编 制 人：秦彦斌
	' 编制日期：2013年3月16日
	'           2013年5月29日修改：增加三个应力为拉力的判断后，再取三轴等效应力和三个拉应力，共4个应力中的最大值。
	'           2013年遵照窦QQ指示（见刘古3井计算分析情况.doc），dwall=d0-2*t
    '           2017年8月，课题组认真核对公式。
    '           2025年8月，对照GB/T20657-2022，再次核对公式。
	'
	'输入参数：
	'   d0          管体名义外径                                 单位：mm
	'   t           套管的名义壁厚                               单位：mm
	'   e           弹性模量                                     单位：MPa，一般钢材E=206842MPa
	'   kwall       表示管壁厚度名义公差的参数，若偏差12.5% 则kwall=0.875，一般取kwall=0.875
	'   c           管体弯曲曲率，管材中心线弯曲半径的倒数       单位：rad/m
	'   pi          内压                                         单位：MPa
	'   po          外压                                         单位：MPa
	'   Fa          轴向载荷                                     单位：N
	'   N           侧向接触力                                   单位：N
	'   W           弯矩                                         单位：Nm
	'   Tr          扭矩                                         单位：Nm
	'
	'函数返回：管体最大三轴等效应力或者在三向拉应力时比三轴等效应力大的主应力,实型数，单位：MPa
	'************************************************************************************************************************************************************************
	Function cal_xgm_e(ByVal d0 As Double, ByVal t As Double, ByVal e As Double, ByVal kwall As Double, ByVal c As Double, ByVal pi As Double, ByVal po As Double, ByVal Fa As Double, ByVal N As Double, ByVal W As Double, ByVal Tr As Double) As Double
        Dim xgm_e3 As Double
        Dim xgm_e1 As Double
        Dim xgm_b As Double
        Dim xgm_b1 As Double
        Dim xgm_b2 As Double
        Dim xgm_h1 As Double
        Dim xgm_r2 As Double
        Dim xgm_r As Double
        Dim Jp As Double
        Dim d As Double
        Dim r As Double
        Dim dwall As Double
        Dim Ap As Double
        Dim xgm_a As Double
        Dim xgm_r1 As Double
        Dim xgm_h As Double
        Dim xgm_h2 As Double
        Dim tao_ha As Double
        Dim xgm_e2 As Double
        Dim xgm_e4 As Double
        Dim d_I As Double
		Dim pai As Double
		Dim ls_bl As Double
		pai = 3.14159265358979
		cal_xgm_e = 0#
		
        d = d0 - 2 * t
		'2013年遵照窦QQ指示（见刘古3井计算分析情况.doc），dwall=d0-2*t
		'dwall = d0 - 2 * kwall * t
		dwall = d
		Ap = 0.25 * pai * (d0 * d0 - d * d)
		Jp = pai * (d0 ^ 4 - d ^ 4) / 32#
		d_I = pai * (d0 ^ 4 - d ^ 4) / 64#
		
		'轴力产生的拉应力
		xgm_a = Fa / Ap
		'计算外壁上的应力
		r = d0 / 2#
		xgm_r = ((pi * dwall * dwall - po * d0 * d0) - (pi - po) * dwall * dwall * d0 * d0 / (4# * r * r)) / (d0 * d0 - dwall * dwall)
		xgm_h = ((pi * dwall * dwall - po * d0 * d0) + (pi - po) * dwall * dwall * d0 * d0 / (4# * r * r)) / (d0 * d0 - dwall * dwall)
		tao_ha = 1000# * Tr * r / Jp + N / Ap
        xgm_b = -1000.0# * W * r / d_I
        ls_bl = xgm_r * xgm_r + xgm_h * xgm_h + (xgm_a + xgm_b) * (xgm_a + xgm_b) - xgm_r * xgm_h - xgm_r * (xgm_a + xgm_b) - xgm_h * (xgm_a + xgm_b) + 3.0# * tao_ha * tao_ha
		If ls_bl <= 0 Then
			xgm_e1 = 0
		Else
			xgm_e1 = System.Math.Sqrt(ls_bl)
		End If
		xgm_b = 1000.0# * W * r / d_I
        ls_bl = xgm_r * xgm_r + xgm_h * xgm_h + (xgm_a + xgm_b) * (xgm_a + xgm_b) - xgm_r * xgm_h - xgm_r * (xgm_a + xgm_b) - xgm_h * (xgm_a + xgm_b) + 3.0# * tao_ha * tao_ha
		If ls_bl <= 0 Then
			xgm_e2 = 0
		Else
			xgm_e2 = System.Math.Sqrt(ls_bl)
		End If
		xgm_r1 = xgm_r
		xgm_h1 = xgm_h
        xgm_b1 = xgm_b
		'计算内壁上的应力
		r = d / 2#
		xgm_r = ((pi * dwall * dwall - po * d0 * d0) - (pi - po) * dwall * dwall * d0 * d0 / (4# * r * r)) / (d0 * d0 - dwall * dwall)
		xgm_h = ((pi * dwall * dwall - po * d0 * d0) + (pi - po) * dwall * dwall * d0 * d0 / (4# * r * r)) / (d0 * d0 - dwall * dwall)
		tao_ha = 1000# * Tr * r / Jp + N / Ap
        xgm_b = -1000.0# * W * r / d_I
        ls_bl = xgm_r * xgm_r + xgm_h * xgm_h + (xgm_a + xgm_b) * (xgm_a + xgm_b) - xgm_r * xgm_h - xgm_r * (xgm_a + xgm_b) - xgm_h * (xgm_a + xgm_b) + 3.0# * tao_ha * tao_ha
		If ls_bl <= 0 Then
			xgm_e3 = 0
		Else
			xgm_e3 = System.Math.Sqrt(ls_bl)
		End If
        xgm_b = 1000.0# * W * r / d_I
        ls_bl = xgm_r * xgm_r + xgm_h * xgm_h + (xgm_a + xgm_b) * (xgm_a + xgm_b) - xgm_r * xgm_h - xgm_r * (xgm_a + xgm_b) - xgm_h * (xgm_a + xgm_b) + 3# * tao_ha * tao_ha
		If ls_bl <= 0 Then
			xgm_e4 = 0
		Else
			xgm_e4 = System.Math.Sqrt(ls_bl)
		End If
		xgm_r2 = xgm_r
		xgm_h2 = xgm_h
        xgm_b2 = xgm_b

		'取4个等效应力中的最大值
		cal_xgm_e = xgm_e1
		If xgm_e2 > cal_xgm_e Then
            cal_xgm_e = xgm_e2
        End If
		If xgm_e3 > cal_xgm_e Then
            cal_xgm_e = xgm_e3
        End If
		If xgm_e4 > cal_xgm_e Then
            cal_xgm_e = xgm_e4
        End If

		'当三向拉应力时，应用第一强度理论
        If xgm_r1 > 0 And xgm_h1 > 0 And (xgm_a + xgm_b1) > 0 Then
            If xgm_r1 > cal_xgm_e Then
                cal_xgm_e = xgm_r1
            End If
            If xgm_h1 > cal_xgm_e Then
                cal_xgm_e = xgm_h1
            End If
            If (xgm_a + xgm_b1) > cal_xgm_e Then
                cal_xgm_e = xgm_a + xgm_b1
            End If
        End If
        If xgm_r2 > 0 And xgm_h2 > 0 And (xgm_a + xgm_b2) > 0 Then
            If xgm_r2 > cal_xgm_e Then
                cal_xgm_e = xgm_r2
            End If
            If xgm_h2 > cal_xgm_e Then
                cal_xgm_e = xgm_h2
            End If
            If (xgm_a + xgm_b2) > cal_xgm_e Then
                cal_xgm_e = xgm_a + xgm_b2
            End If
        End If
	End Function
	'************************************************************************************************************************************************************************
	'                    抗内压5                   用均匀壁厚模型API公式计算磨损套管的剩余抗内压强度
	'
	' 算法思路：认为磨损后的套管的壁厚是均匀磨损，用API5CT公式按新的减薄的壁厚重新计算强度。
	'           注：2012年7月20日,按照窦的意思，此非均匀壁厚模型，是“秦彦斌算法”。2012年8月11日编入系统
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kny        套管的初始抗内压强度 MPa
	'
	'函数返回：磨损套管的剩余抗内压强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykyqd_API5CT(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kny As Double) As Double
        Dim pi As Double
		Dim pi0 As Double
		Dim sy_t As Double
		If dlt_t >= t Then
			cal_sykyqd_API5CT = 0
			Exit Function
		End If
		sy_t = t - dlt_t
		pi0 = cal_kangneiya_qd(yp, d0, t)
        pi = cal_kangneiya_qd(yp, d0, sy_t)
		cal_sykyqd_API5CT = cs_kny * pi / pi0
	End Function
	
	'************************************************************************************************************************************************************************
	'                   抗内压4                用月牙形模型Song公式计算磨损套管的剩余抗内压强度
	'
	' 算法公式来源：Song J S , Bowen Jerry, Klementich Frank. The internal pressure capacity of crescent-shaped wear casing.SPE 23902, 1992。
	' 完整公式描述见:
	'      (1)法炜论文46页,
	'      (2)2001年，韩勇博士论文，《钻杆接头与套管摩擦磨损问题的理论与试验研究》电子版第75页
	'
	' 本段程序主要依据韩勇博士论文和SPE23902编写。
	'      经编程验算：在SPE23902中，内表面应力公式中，Mo应为Mi。
	'                  在韩勇博士论文中，程序中变量c、e、Mi、Mo算式与原文不一致。
	'
	'                                               -------------------------------秦彦斌 2012年4月7日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  BHA_d        钻杆接头外径    mm
	'  cs_kny        套管的初始抗内压强度 MPa
	'
	'函数返回：磨损套管的剩余抗内压强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykyqd_Song(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal BHA_d As Double, ByVal cs_kny As Double) As Double
		'r2 套管外径 r1 套管内径 r 钻杆接头外径  remain_t剩余壁厚  k套管中心与磨损圆心间距离 k=r2-r-remain_t
        Dim pi20, pi2, M4, a40, a3, c, a, r1, r2, r, k, b, e, a4, M3, pi1, pi10, pic1 As Double
		Dim pic2 As Double
		Dim f0 As Double
		Dim f1 As Double
		
		If dlt_t >= t Then
			cal_sykyqd_Song = 0
			Exit Function
		End If
		
        r2 = 0.5 * d0
		r1 = r2 - t
		r = 0.5 * BHA_d
		k = r1 - r + dlt_t
		'磨损后强度计算
		a = System.Math.Sqrt((r2 * r2 - (r - k) * (r - k)) * (r2 * r2 - (r + k) * (r + k))) / 2# / k '韩勇博士论文68页，a
		b = (r2 * r2 - r * r - k * k) / 2# / k - a '韩勇博士论文69页，b
		f0 = System.Math.Log((2 * a + b + k + r2) / (b + k + r2)) '韩勇博士论文68页，kco
		f1 = System.Math.Log((2 * a + b + r) / (b + r)) '韩勇博士论文68页，kci
		c = (dlt_t / (r1 + r2)) * ((1.36 - System.Math.Exp((-1#) * dlt_t / t)) ^ (-2)) '韩勇博士论文68页,式3.4，3.5系数
		e = 1# / ((sinh(f1)) ^ 2 - (sinh(f0)) ^ 2) '韩勇博士论文68页,式3.4，3.5系数
		a3 = 2# * r1 * r1 / (r2 * r2 - r1 * r1) 'song原始公式，韩勇博士论文68页,式3.4，3.5系数，一致
		a4 = r1 * r1 * ((k + r) * (k + r) + r2 * r2) / ((k + r) * (k + r)) / (r2 * r2 - r1 * r1) 'song原始公式，韩勇博士论文68页,式3.4，3.5系数，一致
		M3 = -2# * (sinh(f0) * sinh(f0) - sinh(f1) * (1 + cosh(f0)) / (sinh(f1 - f0))) '韩勇博士论文68页,Mo中pi系数
		M4 = -2# * (sinh(f0) * sinh(f0) - sinh(f0) * (1 + cosh(f1)) / (sinh(f1 - f0))) '韩勇博士论文68页,Mi中pi系数
		
		'c = (dlt_t / (r1 + r2)) ^ ((1.36 - Exp((-1#) * dlt_t / t)) ^ (-2))                         'song原始公式
		'e = 1# / ((sinh(f1)) ^ 2 + (sinh(f0)) ^ 2)                                                 'song原始公式
		'a3 = 2# * r1 * r1 / (r2 * r2 - r1 * r1)                                                    'song原始公式
		'a4 = r1 * r1 * ((k + r) * (k + r) + r2 * r2) / ((k + r) * (k + r)) / (r2 * r2 - r1 * r1)   'song原始公式
		'M3 = sinh(f0) * sinh(f0) - sinh(f1) * (1 + cosh(f0)) / (sinh(f1 - f0))                     'song原始公式（5）
		'M4 = sinh(f0) * sinh(f0) - sinh(f0) * (1 + cosh(f1)) / (sinh(f1 - f0))                     'song原始公式（6）
		
		
		'未磨损时c=0，a4为
		a40 = (r2 * r2 + r1 * r1) / (r2 * r2 - r1 * r1) 'song原始公式，韩勇博士论文68页,一致
		pi1 = yp / (a3 + c * e * M3)
		pi2 = yp / (a4 + c * e * M4)
		pi10 = yp / a3
		pi20 = yp / a40
		
		'分别计算剩余抗内压强度
		pic1 = cs_kny * pi1 / pi10
		pic2 = cs_kny * pi2 / pi20
		
		'取小值返回
		If pic1 < pic2 Then
            cal_sykyqd_Song = pic1
        Else
            cal_sykyqd_Song = pic2
        End If
	End Function
	
	'**********************************************************************************************************************************************************
	'              抗内压3     用Klever-Stewart提出的未磨损套管塑性破坏内压强度设计模型计算磨损套管的剩余抗内压强度
	'
	' 算法公式来源：
	'     廖华林，管志川，马广军，冯光通.深井超深井内壁磨损套管剩余强度计算[J].工程力学.2010，27(2):251-256.中的Klever-Stewart算法,(29)式
	' 算法计算情况：
	'     2013年2月19日，认真核对程序与文献的一致性后，用不同算法比较，发现本算法远没有文中所述的比API算法合理，算出的抗内压值远比API算法
	' 算出的值小！！软件宜放弃此算法。
	'
	' 考证：根据上述文献追朔，该公式来源于：
	'   ISO/TR 10400-2007, International organization for standardization. Petroleum and natural gas industries-equations and calculations
	' for the properties of casing,tubing, drill pipe and line pipe used as casing or tubing [S]. Geneva: ISO Copyright Office, 2007.
	'   ISO/TR 10400-2007又指向：
	'   [1] KLEVER, F.J. and STEWART, G., Analytical Burst Strength Prediction of OCTG With and Without Defects,SPE 48329, 1998
	'   [2] STEWART, G. and KLEVER, F.J., Accounting for Flaws in the Burst Strength of OCTG, SPE 48330, 1998
	'
	'                                               -------------------------------秦彦斌 2012年2月12日,2013年2月19日第2次考证
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kny        套管的初始抗内压强度 MPa
	'
	'函数返回：磨损套管的剩余抗内压强度  MPa    ,实型数
	'*********************************************************************************************************************************************************
	Function cal_sykyqd_K_Stewart(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kny As Double) As Double
		Dim nj1 As Double 'n+1,n--n为套管材料应力-应变强度硬化因子，其取值可根据套管材料实际试验曲线来确定，或用经验公式n=0.1693－1.1774×10^4σy，无因次。
		Dim ka As Double 'ka为内压强度系数，调质和13Cr材料套管取1.0，旋转校直套管取2.0，未知时取2.0；
        Dim pi0 As Double
        Dim pi As Double
		
		If dlt_t >= t Then
			cal_sykyqd_K_Stewart = 0
			Exit Function
		End If
		
		ka = 2#
		nj1 = 0.1693 - yp * 1.1774 * 0.0001 + 1
		'计算剩余值
		pi = (2# * yp * (t - ka * dlt_t) * (0.5 ^ nj1 + (1 / System.Math.Sqrt(3)) ^ nj1)) / (d0 - t + ka * dlt_t)
		'计算原始值
        pi0 = (2.0# * yp * t * (0.5 ^ nj1 + (1 / System.Math.Sqrt(3)) ^ nj1)) / (d0 - t)
		cal_sykyqd_K_Stewart = cs_kny * pi / pi0
		If cal_sykyqd_K_Stewart < 0 Then
			cal_sykyqd_K_Stewart = 0
		End If
	End Function
	
	
	'************************************************************************************************************************************************************************
	'                     抗内压2             用偏心圆筒模型小增算法计算磨损套管的剩余抗内压强度
	'
	' 算法公式来源： (2) 窦益华，张福祥，王维君等．井下套管磨损深度及剩余强度分析[J]．石油钻采工艺，2007，29（4）：36-39。
	'                (2) 王小增硕士论文，井下套管磨损及剩余强度分析，式（4.24）-式（4.26）,式（4.47），式（4.55）-式（4.57）
	'                                               -------------------------------秦彦斌 2012年3月30日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kny        套管的初始抗内压强度 MPa
	'
	'函数返回：磨损套管的剩余抗内压强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykyqd_pxyt(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kny As Double) As Double
        Dim pi0, fi, kci, r0, a, e, r1, kco, m, pi As Double
		Dim dbt As Double
		
		If dlt_t >= t Then
			cal_sykyqd_pxyt = 0
			Exit Function
		End If
		
		dbt = d0 / t
		r0 = 0.5 * d0
		r1 = r0 - (t + t - dlt_t) / 2
		e = 0.5 * dlt_t
		a = System.Math.Sqrt(r1 ^ 4 - 2# * e * e * r1 * r1 + r0 ^ 4 - 2# * e * e * r0 * r0 - 2# * r0 * r0 * r1 * r1 - e ^ 4) / (2# * e)
		kco = arcsinh(a / r0)
		kci = arcsinh(a / r1)
		m = (a / r0) * (a / r0) + (a / r1) * (a / r1)
		fi = ((2# * (a / r0) + sinh(kco + kci)) / sinh(kci - kco) - 1 - 2# * (a / r0) * (a / r0)) / m + 1 '小增论文(4.47)式f4
		fi = System.Math.Abs(fi)
		pi = yp / fi
		pi0 = yp * (dbt - 1) / (0.5 * dbt * dbt - dbt + 1)
		cal_sykyqd_pxyt = cs_kny * pi / pi0
	End Function
	
	'************************************************************************************************************************************************************************
	'                    抗内压1                   用均匀壁厚模型常规算法计算磨损套管的剩余抗内压强度
	' 注：算法思路为强度降低与壁厚磨损成正比
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kny        套管的初始抗内压强度 MPa
	'
	'函数返回：磨损套管的剩余抗内压强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykyqd_jybh(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kny As Double) As Double
		Dim sy_t As Double
		If dlt_t >= t Then
			cal_sykyqd_jybh = 0
			Exit Function
		End If
		sy_t = t - dlt_t
		cal_sykyqd_jybh = cs_kny * sy_t / t
	End Function
	
	'***************************************************************************************************************************
	'                    抗挤7                   用均匀壁厚模型API公式计算磨损套管的剩余抗挤强度
	'
	' 算法思路：认为磨损后的套管的壁厚是均匀磨损，用API5CT公式按新的减薄的壁厚重新计算强度。
	'           注：按照窦的意思，此非均匀壁厚模型，是“秦彦斌算法”,2012年7月20日。2012年8月6日编入系统
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kj        套管的初始抗挤强度 MPa
	'
	'函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
	'***************************************************************************************************************************
	Function cal_sykjqd_API5CT(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
        Dim py As Double
		Dim pocr As Double
		Dim kocr As Double '内壁磨损套管抗挤毁强度降低系数
		Dim sy_t As Double
		
		If dlt_t >= t Then
			cal_sykjqd_API5CT = 0
			Exit Function
		End If
		
		sy_t = t - dlt_t
        py = cal_kangji_qd(yp, d0, t, 0, 0)
        pocr = cal_kangji_qd(yp, d0, sy_t, 0, 0)
		kocr = pocr / py
		cal_sykjqd_API5CT = cs_kj * kocr
	End Function
	
	'************************************************************************************************************************************************************************
	'                 抗挤6                                用视磨损为缺陷ISO 10400方法计算磨损套管的剩余抗挤强度
	'
	' 算法公式来源：
	'    曾德智，龚龙祥，付建红，施太和，胡勇．全井段套管磨损量预测与磨损后抗挤强度计算方法研究．钢管2010年7月第39卷增刊
	'    ISO 10400: 2004. Petroleumand natural gas industries formula and calculation for casing, tubing, drill pipe and line pipe Properties [s]
	'    孙永兴,林元华,舒玉春,张智,施太和,ISO 10400油套管强度新模型,石油钻探技术,2008年1月，第36卷第1期
	'
	'                                                                              -------------------------------秦彦斌 2012年3月29日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kj        套管的初始抗挤强度 MPa
	'
	'函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykjqd_ISO10400(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
        Dim dbt, hn, pe, xgmr, dlt0, epsl, Hult, py, mu, e, po0 As Double
		Dim po As Double
        xgmr = 0.0#
		hn = 0#
		e = 206800#
		mu = 0.3
		If dlt_t >= t Then
			cal_sykjqd_ISO10400 = 0
			Exit Function
		End If
		dbt = d0 / t
		epsl = 200# * dlt_t / (2# * t - dlt_t) 'epsl
		dlt0 = 200# * dlt_t / (2# * d0 - 4# * t + dlt_t) 'e
		py = 2# * yp * (dbt - 1) / (dbt * dbt) * (1 + 1.47 / (dbt - 1))
		pe = 2# * e / (1 - mu * mu) / (dbt * (dbt - 1) * (dbt - 1))
		Hult = 0
		po0 = (pe + py - System.Math.Sqrt((pe - py) * (pe - py) + 4# * Hult * pe * py)) / (2# * (1 - Hult))
		Hult = 0.127 * dlt0 + 0.003 * epsl - 0.44 * xgmr / yp + hn
		po = (pe + py - System.Math.Sqrt((pe - py) * (pe - py) + 4# * Hult * pe * py)) / (2# * (1 - Hult))
		cal_sykjqd_ISO10400 = cs_kj * po / po0
	End Function
    '************************************************************************************************************************************************************************
    '                 抗挤5                                           用偏心圆筒模型Kuriyama公式计算磨损套管的剩余抗挤强度
    '
    ' 算法公式来源：
    '     Kuriyama Yukihisa;Tsukano Yasushi;Mimaki Toshitaro,et al. Effect of wear and bending on casing collapse strength(SPE 24597)[R],Proceedings-SPE Annual
    ' Technical Conference and Exhibition,Proceedings of the 1992 SPE Annual Technical Conference and Exhibition,Washington DC,1992,Houston:Society of
    ' Petroleum Engineers Inc,1992:543-552。该文献的中文翻译见：彭维健译，夏月泉校，磨损和弯曲对套管挤毁强度的影响，国外钻井技术，1993年第六期，P53-59.
    '
    '    注：该文献中a的计算公式与图中注的不符，应该是a=b-(t+tmin)/2
    '
    '                                                                              -------------------------------秦彦斌 2012年3月24日
    '
    '输入参数：
    '   yp          套管的屈服强度  MPa
    '   d0          套管的公称外径  mm
    '   t           套管的公称壁厚  mm
    '  dlt_t        套管的磨损深度  mm
    '  cs_kj        套管的初始抗挤强度 MPa
    '
    '函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
    '************************************************************************************************************************************************************************
	Function cal_sykjqd_SPE24597(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
        Dim po, b, a, c, po0 As Double

		If dlt_t >= t Then
			cal_sykjqd_SPE24597 = 0
			Exit Function
		End If
        b = 0.5 * d0
		a = b - t
		c = 0#
		po0 = yp * ((a * a + b * b) / (2# * b * b)) * ((b * b + a * a - c * c) * (b * b + a * a - c * c) - 4# * a * a * b * b) / ((b * b - c * c) * (b * b - c * c) - a * a * (a - 2# * c) * (a - 2# * c))
		a = b - t + 0.5 * dlt_t
		c = 0.5 * dlt_t
		po = yp * ((a * a + b * b) / (2# * b * b)) * ((b * b + a * a - c * c) * (b * b + a * a - c * c) - 4# * a * a * b * b) / ((b * b - c * c) * (b * b - c * c) - a * a * (a - 2# * c) * (a - 2# * c))
		cal_sykjqd_SPE24597 = cs_kj * po / po0
	End Function
    '************************************************************************************************************************************************************************
    '                 抗挤4                                           用视磨损为缺陷Tamano方法计算磨损套管的剩余抗挤强度
    '
    ' 算法公式来源：见下面的考证
    '
    ' 考证：
    '     1、1983年，此算法可能最早发表于：Tamano.T.,Mimaki.T.et al.:"A New Empirical Formula for Collapse Resistance of Commercial Casing",
    ' Proceedings of the second International Offshore Mechanics and Arctic Engineering Symposium.ETCE(1983),pp.489
    '     因为在文献：
    '     2、1992年，Kuriyama Yukihisa;Tsukano Yasushi;Mimaki Toshitaro,et al. Effect of wear and bending on casing collapse strength(SPE 24597)[R],Proceedings-SPE Annual
    ' Technical Conference and Exhibition,Proceedings of the 1992 SPE Annual Technical Conference and Exhibition,Washington DC,1992,Houston:Society of
    ' Petroleum Engineers Inc,1992:543-552。该文献的中文翻译见：彭维健译，夏月泉校，磨损和弯曲对套管挤毁强度的影响，国外钻井技术，1993年第六期，P53-59.
    ' 中，引用了Tamano的估算真实套管挤毁强度的经验公式。
    '
    '     3、2005年，文献【曾德智，林元华，施太和等．磨损套管抗挤强度的新算法研究[J]．钻井工程，2005，25（2）：78～81】中，引用了Kuriyama的文章公式，与SPE 24597
    ' 对照，文中g的计算公式中系数0.00228应为0.00456。

    '     4、2007年，博士论文"复杂井况下套管的可靠性与风险评估研究.kdh"中第20-21（电子32页）页，“2.2.1制造及磨损等缺陷对套管强度的影响”节引用了Kuriyama的文章公
    ' 式，其中g的计算公式中系数与Kuriyama的文章公式一致。
    '
    '     5、2010年2月， 文献：【廖华林，管志川，马广军，冯光通 深井超深井内壁磨损套管剩余强度计算 工程力学 2010年2月】中，直接给出了含缺陷套管的抗挤强度计算公式，
    ' 未注出处。此公式与Kuriyama的文章公式一致，但泊松比u的指数2丢失，且文中g的计算公式中系数为0.00228，与2005年曾德智一样，估计是抄来的。
    '
    '     6、2010年7月，文献【曾德智，龚龙祥，付建红，施太和，胡勇．全井段套管磨损量预测与磨损后抗挤强度计算方法研究．钢管2010年7月第39卷增刊】中，再次引用了
    ' Kuriyama的文章公式，文中g的计算公式中系数为0.00228，与2005年文献一样，属炒冷饭。文中还引用了 ISO 10400 给出的考虑套管缺陷的抗挤强度计算公式
    '     注：ISO 10400:2004 Petroleum and natural gas industries-formula and calculation for casing,tubing,drill pipe and line pipe properties［S］.
    '
    '
    '    根据前面文献中不圆度e的定义，e=2(Dmax-Dmin)/(Dmax+Dmin),而Dmin=D-2t,Dmax=D-2t+dlt_t,则e=2dlt_t/(2D-4t+dlt_t),前述文献3中图1、文献5中图2、文献6中图4均错误，误导人！
    '
    '                                                                              -------------------------------秦彦斌 2012年3月24日
    '
    '输入参数：
    '   yp          套管的屈服强度  MPa
    '   d0          套管的公称外径  mm
    '   t           套管的公称壁厚  mm
    '  dlt_t        套管的磨损深度  mm
    '  cs_kj        套管的初始抗挤强度 MPa
    '
    '函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
    '************************************************************************************************************************************************************************
	Function cal_sykjqd_quexian(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
        Dim po0, e, pe, xgmr, dlt0, epsl, g, py, mu, dbt, po As Double

		If dlt_t >= t Then
			cal_sykjqd_quexian = 0
			Exit Function
		End If
		xgmr = 0#
		e = 206800#
		mu = 0.3
		dbt = d0 / t
		epsl = 200# * dlt_t / (2# * t - dlt_t) '200=2*100,100为百分比
		dlt0 = 200# * dlt_t / (2# * d0 - 4# * t + dlt_t)
		py = 2# * yp * (dbt - 1) / (dbt * dbt) * (1 + 1.47 / (dbt - 1))
		pe = 2# * e / (1 - mu * mu) / (dbt * (dbt - 1) * (dbt - 1))
		g = 0
		po0 = 0.5 * (pe + py - System.Math.Sqrt((pe - py) * (pe - py) + g * pe * py))
		g = 0.3232 * dlt0 + 0.00456 * epsl - 0.5648 * xgmr / yp
		po = 0.5 * (pe + py - System.Math.Sqrt((pe - py) * (pe - py) + g * pe * py))
		cal_sykjqd_quexian = cs_kj * po / po0
	End Function
	
	'************************************************************************************************************************************************************************
	'                    抗挤3          用月牙形模型song公式计算磨损套管的剩余抗挤强度
	'
	' 算法公式来源：Song J S , Bowen Jerry, Klementich Frank. The internal pressure capacity of crescent-shaped wear casing.SPE 23902, 1992。
	' 完整公式描述见
	'      (1)法炜论文46页,
	'      (2)2001年，韩勇博士论文，《钻杆接头与套管摩擦磨损问题的理论与试验研究》电子版第75页
	'
	' 本段程序主要依据韩勇博士论文编写。经编程验算：在SPE23902中，内表面应力公式中，Mo应为Mi。在韩勇博士论文中，程序中变量e值算式错误。
	'
	' 此处程序有问题，算出的结果不符合“哈数”！ 2012年8月11日于济南
	'                                               -------------------------------秦彦斌 2012年4月7日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  BHA_d        钻杆接头外径    mm
	'  cs_kj        套管的初始抗挤强度 MPa
	'
	'函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
	'************************************************************************************************************************************************************************
	Function cal_sykjqd_song(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal BHA_d As Double, ByVal cs_kj As Double) As Double
		'r2 套管外径 r1 套管内径 r 钻杆接头外径  remain_t剩余壁厚  k套管中心与磨损圆心间距离 k=r2-r-remain_t
        Dim poc2, po20, po2, M2, a20, a1, c, a, remain_t, r1, r2, r, k, b, e, a2, M1, po1, po10, poc1 As Double
        Dim f0 As Double
		Dim f1 As Double
		
		If dlt_t >= t Then
			cal_sykjqd_song = 0
			Exit Function
		End If
		
		'计算剩余值
		r2 = 0.5 * d0
		r1 = r2 - t
		r = 0.5 * BHA_d
		remain_t = t - dlt_t
		k = r2 - r - remain_t
		'有磨损强度计算
		a = System.Math.Sqrt((r2 * r2 - (r - k) * (r - k)) * (r2 * r2 - (r + k) * (r + k))) / 2# / k '韩勇博士论文68页，a
		b = (r2 * r2 - r * r - k * k) / 2# / k - a '韩勇博士论文69页，b
		f0 = System.Math.Log((2 * a + b + k + r2) / (b + k + r2)) '韩勇博士论文68页，kec0
		f1 = System.Math.Log((2 * a + b + r) / (b + r)) '韩勇博士论文68页，kec1
		c = (dlt_t / (r2 + r1)) * ((1.36 - System.Math.Exp((-1#) * dlt_t / t)) ^ (-2)) '韩勇博士论文68页，式3.4,式3.5系数
		e = 1# / (sinh(f1) * sinh(f1) - sinh(f0) * sinh(f0)) '韩勇博士论文68页，式3.4,式3.5系数
		a1 = ((r2 * r2 + r1 * r1) / (r2 * r2 - r1 * r1)) 'song原始公式(3)，韩勇博士论文68页，式3.4,式3.5系数，一致
		a2 = r2 * r2 * ((k + r) * (k + r) + r1 * r1) / ((k + r) * (k + r)) / (r2 * r2 - r1 * r1) 'song原始公式(4)，韩勇博士论文68页，式3.4,式3.5系数，一致
		M1 = -2# * (sinh(f1) * sinh(f1) + sinh(f1) * (1 + cosh(f0)) / (sinh(f1 - f0))) 'song原始公式(5)，韩勇博士论文68页，Mo系数，一致
		M2 = -2# * (sinh(f1) * sinh(f1) + sinh(f0) * (1 + cosh(f1)) / (sinh(f1 - f0))) 'song原始公式(6)，韩勇博士论文68页，Mo系数，一致
		
		'c = (dlt_t / (r2 + r1)) ^ ((1.36 - Exp((-1#) * dlt_t / t)) ^ (-2))                             'song原始公式
		'e = 1# / (sinh(f1) * sinh(f1) + sinh(f0) * sinh(f0))                                           'song原始公式(3)、(4)中系数
		'a1 = ((r2 * r2 + r1 * r1) / (r2 * r2 - r1 * r1))                                               'song原始公式(3)
		'a2 = r2 * r2 * ((k + r) * (k + r) + r1 * r1) / ((k + r) * (k + r)) / (r2 * r2 - r1 * r1)       'song原始公式(4)
		'M1 = -2# * (sinh(f1) * sinh(f1) + sinh(f1) * (1 + cosh(f0)) / (sinh(f1 - f0)))                 'song原始公式(5)
		'M2 = -2# * (sinh(f1) * sinh(f1) + sinh(f0) * (1 + cosh(f1)) / (sinh(f1 - f0)))                 'song原始公式(6)
		
		'无磨损强度计算时，c=0
		a20 = 2# * r2 * r2 / (r2 * r2 - r1 * r1)
		po1 = yp / (a1 + c * M1 * e + c)
		po2 = yp / (a2 + c * M2 * e + c)
		po10 = yp / a1
		po20 = yp / a20
		poc1 = cs_kj * po10 / po1
		poc2 = cs_kj * po20 / po2
		
		'判断取值
		If poc1 < poc2 Then
            cal_sykjqd_song = poc1
        Else
            cal_sykjqd_song = poc2
        End If
	End Function
	
    '************************************************************************************************************************************************************************
    '                    抗挤2                   用偏心圆筒模型小增算法计算磨损套管的剩余抗挤强度
	'
	' 算法公式来源： (1) 窦益华，张福祥，王维君等．井下套管磨损深度及剩余强度分析[J]．石油钻采工艺，2007，29（4）：36-39。
	'                (2) 王小增硕士论文，井下套管磨损及剩余强度分析，式（4.24）-式（4.26）,式（4.36），式（4.58）-式（4.61）
	'                                               -------------------------------秦彦斌 2012年2月12日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kj        套管的初始抗挤强度 MPa
	'
	'函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
    '************************************************************************************************************************************************************************
    Function cal_sykjqd_pxyt(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
        Dim fo, kci, r0, a, d0bt, e, r1, kco, m, po0 As Double
        Dim po As Double

        If dlt_t >= t Then
            cal_sykjqd_pxyt = 0
            Exit Function
        End If

        d0bt = d0 / t
        r0 = 0.5 * d0
        r1 = r0 - (t + t - dlt_t) / 2
        e = 0.5 * dlt_t
        a = System.Math.Sqrt(r1 ^ 4 - 2.0# * e * e * r1 * r1 + r0 ^ 4 - 2.0# * e * e * r0 * r0 - 2.0# * r0 * r0 * r1 * r1 - e ^ 4) / (2.0# * e)
        kco = arcsinh(a / r0)
        kci = arcsinh(a / r1)
        m = (a / r0) * (a / r0) + (a / r1) * (a / r1)
        fo = (-2.0# * (a / r0) - sinh(kco + kci)) / sinh(kci - kco) + 1 - 2.0# * (a / r1) * (a / r1) '小增论文(4.37)式f2
        fo = System.Math.Abs(fo / m)
        po0 = 2.0# * yp * (d0bt - 1) / (d0bt * d0bt)
        po = yp / fo
        cal_sykjqd_pxyt = cs_kj * po / po0
    End Function
	
	'***************************************************************************************************************************
	'                    抗挤1                   用均匀壁厚模型常规算法计算磨损套管的剩余抗挤强度
	'算法思路：壁厚降低的比例即为强度降低的比例
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
	'  dlt_t        套管的磨损深度  mm
	'  cs_kj        套管的初始抗挤强度 MPa
	'
	'函数返回：磨损套管的剩余抗挤强度  MPa    ,实型数
	'***************************************************************************************************************************
	Function cal_sykjqd_jybh(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal dlt_t As Double, ByVal cs_kj As Double) As Double
		Dim sy_t As Double
		If dlt_t >= t Then
            cal_sykjqd_jybh = 0
            Exit Function
        End If
		sy_t = t - dlt_t
		cal_sykjqd_jybh = cs_kj * sy_t / t
	End Function

	'***************************************************************************************************************************
	'                        抗挤0                      用API公式计算套管的抗挤强度
	'
    '     API根据套管挤压破坏形式按厚径比分为四种计算方法。本函数就是实现此四种方法以计算套管的抗挤强度。
    '     算法公式来源：根本是API，最新标准为：《GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》

    '     算法考证： 
    '         《油气井管柱力学与工程》 高德利 中国石油大学出版社 ISBN：9787563621583，第145-147页所述公式实质来自API，书中将美国惯用单位
    '       制换算为公制，计算系数转换是否正确不好说，至少反复查明：书中第146页式（4-4-5）书写错误，应为：pt= yp * (f / (d0/t) - g)
    '          2012年3月，对照GB/T 20657-2006对程序中的公式及参数核实、核对，计算基本正确。
    '          2025年3月，对照GB/T 20657-2022再次对程序中的公式及参数核实、核对。考虑到来源公式中参数单位采用美国惯用单位制，为保证公式
    '       的一致性，2025年3月17日修改程序，先将传入的公有制参数换算成英制，然后用原公式计算，最后再将计算结果换回公制。
    '
    '                                                       -------------------------------秦彦斌 2025年3月17日最后修订
    '输入参数：
    '   yp          套管的屈服强度  MPa
    '   d0          套管的公称外径  mm
    '   t           套管的公称壁厚  mm
    '   xgmz         轴向应力       MPa
    '   pi           内压           MPa
    '
    '函数返回：套管的抗挤强度  MPa    ,实型数
    '***************************************************************************************************************************
    '***************************************************************************************************************************
    '                                       以下摘自wellCAT帮助文件
    '                                                      -------------------------------秦彦斌 2021年9月23日
    '  topic:2015 API TR 5C3 Collapse Formula Update WELLCAT Workflow
    '     fycom=fymn*(sqrt(1-0.75*((xgma+pi)/fymn)*(xgma+pi)/fymn))-0.5*(xgma+pi)/fymn)    (for (xgma+pi)>=0)     (42)
    '
    ' fycom is the combined loading equivalent grade, the equivalent yield strength in the presence of axial stress and internal pressure;
    ' fymn is the specified minimum yield strength;
    ' σa is the component of axial stress not due to bending; 
    ' pi is the internal pressure;
    ' pc is the collapse resistance.

    'Collapse resistance equation factors and D/t ranges for the combined loading equivalent grade are then calculated by means of 
    'Equations (36), (38), (40), (44) or (49), (45) or (50), (46) or (51), (47) or (52), and (48) or (53). Using equation factors 
    'for the combined loading equivalent grade, collapse resistance under axial stress and internal pressure is calculated by means of 
    'Equations (35), (37), (39) and (41). 

    'API collapse resistance equations are not valid for the yield strength of combined loading equivalent grade (fycom) less than 24 000 psi.

    'Effect of internal pressure on collapse 

    'The external pressure equivalent of external pressure and internal pressure is determined by means of Equation (43) in API 5C3
    ' is not applicable anymore according to the Addendum, Oct. 2015 where the original equation (43) in the first edition 2008 has
    ' been deleted in the addendum, October 2015. Instead, the differential pressure of external pressure and internal pressure is used.

    'References
    'TR 5C3/ISO 10400:2007, Technical Report on Equations and Calculations or Casing, Tubing, and Line Pipe Used as Casing or Tubing; 
    'and Performance Properties Tables for Casing and Tubing, First Edition, December 2008 Addendum, October 2015 

    'CLINEDINST, W.O., Calculating Collapse Resistance under Axial Stress using Existing API Collapse Formulas and the Strain Energy
    ' of Distortion Theory of Yielding, report prepared for the American Petroleum Institute, December 1, 1980

    'API 5C3 Addendum, issued in October 2015, upon API 5C3 First Edition, 2008
    '***************************************************************************************************************************
    Function cal_kangji_qd(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double, ByVal xgmz As Double, ByVal pi As Double) As Double
        Dim d0bt_yp, d0bt, d0bt_pt As Double
        Dim d0bt_te As Double '径厚比及径厚比分界点值
        Dim k, b, a, c, f As Double
        Dim g As Double '计算系数A,B,C,K,F,G
        Dim yp1 As Double '单位为kPa的套管的屈服强度
        Dim bba As Double
        Dim xgmzbyp As Double 'B/A,xgmz/yp

        '返回值赋初值，若计算错误，返回此值。
        cal_kangji_qd = 0.0#
        'xgmzbyp = xgmz / yp
        '单轴向拉力及轴向拉力+内压对挤毁计算可合并
        xgmzbyp = (xgmz + pi) / yp
        '折减的屈服强度。因为1psi（磅/平方英寸）=6.8947572932kPa(千帕)，故先乘以1000转化成kPa，再除以6.8947572932转化成psi。
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P26-27页 （42）、（43）式
        yp1 = yp * 1000.0# * (System.Math.Sqrt(1 - 0.75 * xgmzbyp) - 0.5 * xgmzbyp)
        '将屈服强度转换成英制psi，以后用它算
        k = yp1 / 6.8947572932
        d0bt = d0 / t
        '用原始英制单位，原始公式，算完后再转换回来，免得系数混乱----20250317
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P27页 （49）式
        a = 2.8762 + 0.10679 * 10 ^ (-5) * k + 0.21301 * 10 ^ (-10) * k * k - 0.53132 * 10 ^ (-16) * k * k * k
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P28页 （50）式
        b = 0.026233 + 0.50609 * 10 ^ (-6) * k
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P28页 （51）式
        c = -465.93 + 0.030867 * k - 0.10483 * 10 ^ (-7) * k * k + 0.36989 * 10 ^ (-13) * k * k * k
        bba = b / a
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P28页 （52）式
        f = 46.95 * 10 ^ 6 * ((3 * bba / (2 + bba)) ^ 3) / (k * ((3 * bba / (2 + bba) - bba) * (1 - 3 * bba / (2 + bba)) ^ 2))
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P28页 （53）式
        g = f * bba
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P22页（36）式,P100页（E.2）式
        d0bt_yp = (a - 2 + System.Math.Sqrt((a - 2) * (a - 2) + 8 * (b + c / k))) / (2 * (b + c / k))
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P23页（38）式,P101页（E.4）式
        d0bt_pt = (k * (a - f)) / (c + k * (b - g))
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P24页（40）式,P103页（E.6）式
        d0bt_te = (2 + bba) / (3 * bba)
        '1 拉梅公式，屈服强度挤毁公式
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P22页（35）式,P100页（E.1）式
        If d0bt <= d0bt_yp Then
            cal_kangji_qd = 2 * k * ((d0bt - 1) / (d0bt * d0bt))
        End If
        '2 塑性挤毁公式
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P23页（37）式,P101页（E.3）式
        If d0bt > d0bt_yp And d0bt <= d0bt_pt Then
            cal_kangji_qd = k * (a / d0bt - b) - c
        End If
        '3 塑弹性(过渡)挤毁公式
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P24页（39）式,P103页（E.5）式
        If d0bt > d0bt_pt And d0bt <= d0bt_te Then
            cal_kangji_qd = k * (f / d0bt - g)
        End If
        '4 弹性挤毁公式
        'GB/T 20657-2022 石油天然气工业套管、油管、钻杆和用作套管或油管得管线管性能公式及计算》P25页（41）式,P104页（E.7）式
        If d0bt > d0bt_te Then
            cal_kangji_qd = (46.95 * 10 ^ 6) / (d0bt * (d0bt - 1) ^ 2)
        End If
        '转化回MPa。1psi（磅/平方英寸）=6.8947572932kPa(千帕)
        cal_kangji_qd = System.Math.Round(6.894757 * cal_kangji_qd / 1000.0#, 4)
    End Function


    '***************************************************************************************************************************
    '                       抗内压1                   用GB/T 20657-2022公式计算开口套管的抗内压强度
    '
    ' 算法公式来源： 石油天然气工业套管、油管、钻杆和用作套管或油管的管线管性能公式及计算[S]. GB/T 20657-2022.  第13页 公式（9）
    '    
    '      历史上API采用的巴洛(Barlow)管体屈服公式基于单轴(非三轴),是米塞斯(Mises)屈服条件的近似公式，包括近似的体环向应力作用。
    ' 事实是巴洛(Barlow)公式近似计算了环向立力，然后据此近似计算出屈服强度。这个近似比6.6.1中讨论的立梅(Lame)屈服公式精度要低。
    ' 因为巴洛(Barlow)公式忽略了轴向应力，导致管端封堵、管端开口或管端有拉伸载荷的管子结果没有区别。
    ' ---摘自 中华人民共和国国家标准. 石油天然气工业套管、油管、钻杆和用作套管或油管的管线管性能公式及计算[S]. GB/T 20657-2022.
    '
    '                                               -------------------------------秦彦斌 2025年2月25日最后修订
    '
    '输入参数：
    '   yp          套管的屈服强度  MPa
    '   d0          套管的公称外径  mm
    '   t           套管的公称壁厚  mm
    '
    '函数返回：套管的抗内压强度  MPa    ,实型数
    '***************************************************************************************************************************
    Function cal_kangneiya_qd1(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double) As Double
        '管端开口的厚壁管抗内压计算公式。石油天然气工业套管、油管、钻杆和用作套管或油管的管线管性能公式及计算[S]. GB/T 20657-2022.  第13页 公式（9）
        'cal_kangneiya_qd1 = yp * (d0 * d0 - (d0 - 2 * 0.875 * t) * (d0 - 2 * 0.875 * t)) / System.Math.Sqrt(3 * d0 ^ 4 + (d0 - 2 * 0.875 * t) ^ 4)
        '封堵管厚壁管抗内压计算公式。石油天然气工业套管、油管、钻杆和用作套管或油管的管线管性能公式及计算[S]. GB/T 20657-2022.  第13页 公式（8）
        cal_kangneiya_qd1 = yp / System.Math.Sqrt((3 * d0 ^ 4 + (d0 - 2 * 0.875 * t) ^ 4) / (d0 ^ 2 - (d0 - 2 * 0.875 * t) ^ 2) ^ 2 + (d0 - 2 * 4) ^ 4 / (d0 ^ 2 - (d0 - 2 * t) ^ 2) ^ 2) - 2 * (d0 - 2 * t) ^ 2 * (d0 - 2 * 0.875 * t) ^ 2 / ((d0 ^ 2 - (d0 - 2 * t) ^ 2) * (d0 ^ 2 - (d0 - 2 * 0.875 * t) ^ 2))
        '巴洛公式
        '*  / System.Math.Sqrt(3 * d0 ^ 4 + (d0 - 2 * 0.875 * t) ^ 4)
    End Function

    '***************************************************************************************************************************
    '                       抗内压0                       用API公式计算套管的抗内压强度
    '
    ' Barlow公式，它是由承受内压力的薄壁管周向应力公式，并考虑到壁厚不均匀因素而得到的。
    '                                                                                ---摘自 《油气井管柱力学与工程》
    ' 算法公式来源： 《油气井管柱力学与工程》 高德利 中国石油大学出版社 ISBN：9787563621583，第147页。
    '    
    '                                               -------------------------------秦彦斌 2012年2月3日最后修订
    '
    '输入参数：
    '   yp          套管的屈服强度  MPa
    '   d0          套管的公称外径  mm
    '   t           套管的公称壁厚  mm
    '
    '函数返回：套管的抗内压强度  MPa    ,实型数
    '***************************************************************************************************************************
    Function cal_kangneiya_qd(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double) As Double
        cal_kangneiya_qd = 0.875 * 2 * yp * t / d0
    End Function

	
	
	'***************************************************************************************************************************
	'                     抗拉                       用API公式计算套管的抗拉强度
	'
	' 算法公式来源： 《油气井管柱力学与工程》 高德利 中国石油大学出版社 ISBN：9787563621583，第148页。
	'                                               -------------------------------秦彦斌 2012年2月3日
	'
	'输入参数：
	'   yp          套管的屈服强度  MPa
	'   d0          套管的公称外径  mm
	'   t           套管的公称壁厚  mm
    '
	'函数返回：套管的抗拉强度  kN    ,实型数
	'***************************************************************************************************************************
	Function cal_kangla_qd(ByVal yp As Double, ByVal d0 As Double, ByVal t As Double) As Double
		cal_kangla_qd = 0.25 * 3.1415926535 * (d0 * d0 - (d0 - 2 * t) * (d0 - 2 * t)) * yp / 1000#
	End Function
End Module