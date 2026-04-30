Option Strict Off
Option Explicit On
Module M_通用数值分析
	'***************************************************************************************************************************
	'                                     VB没提供的反余弦函数
	'***************************************************************************************************************************
	Function Arccos(ByVal x As Double) As Double
		'Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
		If x = 1 Then
			Arccos = 0#
		Else
			Arccos = System.Math.Atan(-x / System.Math.Sqrt(1 - x * x)) + 2 * System.Math.Atan(1)
		End If
	End Function
	
	'***************************************************************************************************************************
	'                                     VB没提供的反正弦函数
	'***************************************************************************************************************************
	Function Arcsin(ByVal x As Double) As Double
		'Arcsin(X) = Atn(X / Sqr(-X * X + 1))
		If x = 1 Then
			Arcsin = 3.1415926535 / 2#
		Else
			Arcsin = System.Math.Atan(x / System.Math.Sqrt(1 - x * x))
		End If
	End Function
	
	'***************************************************************************************************************************
	' 数学公式：Hyperbolic Sine（双曲正弦）   HSin(x) = (Exp(x) - Exp(-x)) / 2
	'***************************************************************************************************************************
	Public Function sinh(ByVal x As Double) As Double
		sinh = (System.Math.Exp(x) - System.Math.Exp(-x)) / 2
	End Function
	
	
	'***************************************************************************************************************************
	' 数学公式：Hyperbolic Cosine（双曲余弦） HCos(x) = (Exp(x) + Exp(-x)) / 2
	'***************************************************************************************************************************
	Public Function cosh(ByVal x As Double) As Double
		cosh = (System.Math.Exp(x) + System.Math.Exp(-x)) / 2
	End Function
	
	
	'***************************************************************************************************************************
	' 数学公式：Inverse Hyperbolic Sine（反双曲正弦） HArcsin(X) = Log(X + Sqr(X * X + 1))
	'***************************************************************************************************************************
	Public Function arcsinh(ByVal x As Double) As Double
		arcsinh = System.Math.Log(x + System.Math.Sqrt(x * x + 1))
	End Function
	
	'***************************************************************************************************************************
	' 线性插值函数
	'***************************************************************************************************************************
	Public Function linear_interpolation(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x As Double) As Double
		If x = x1 Then
			linear_interpolation = y1
			Exit Function
		End If
		If x = x2 Then
			linear_interpolation = y2
			Exit Function
		End If
		If x1 = x2 Then
			linear_interpolation = (y1 + y2) / 2#
			Exit Function
		End If
		linear_interpolation = y2 + (x - x2) * (y1 - y2) / (x1 - x2)
	End Function
End Module