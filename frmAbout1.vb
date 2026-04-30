'**********************************************************************************************************************************
' 20190717-从VB6升级到VS2008完成
'**********************************************************************************************************************************
Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAbout1
	Inherits System.Windows.Forms.Form
	
	' 注册表关键字安全选项...
	Const READ_CONTROL As Integer = &H20000
	Const KEY_QUERY_VALUE As Integer = &H1
	Const KEY_SET_VALUE As Integer = &H2
	Const KEY_CREATE_SUB_KEY As Integer = &H4
	Const KEY_ENUMERATE_SUB_KEYS As Integer = &H8
	Const KEY_NOTIFY As Integer = &H10
	Const KEY_CREATE_LINK As Integer = &H20
	Const KEY_ALL_ACCESS As Double = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
	
	' 注册表关键字 ROOT 类型...
	Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Const ERROR_SUCCESS As Short = 0
	Const REG_SZ As Short = 1 ' 独立的空的终结字符串
	Const REG_DWORD As Short = 4 ' 32位数字
	
	Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
	Const gREGVALSYSINFOLOC As String = "MSINFO"
	Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
	Const gREGVALSYSINFO As String = "PATH"
	
	Private Declare Function RegOpenKeyEx Lib "advapi32"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegQueryValueEx Lib "advapi32"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer
	
	
	Private Sub cmdSysInfo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSysInfo.Click
		Call StartSysInfo()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	
	Public Sub StartSysInfo()
		On Error GoTo SysInfoErr
		
		Dim rc As Integer
        Dim SysInfoPath As String

        SysInfoPath = " "
        ' 试图从注册表中获得系统信息程序的路径及名称...
		If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
			' 试图仅从注册表中获得系统信息程序的路径...
		ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then 
			' 已知32位文件版本的有效位置
			'UPGRADE_WARNING: Dir 有新行为。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"”
			If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
				SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
				
				' 错误 - 文件不能被找到...
			Else
				GoTo SysInfoErr
			End If
			' 错误 - 注册表相应条目不能被找到...
		Else
			GoTo SysInfoErr
		End If
		
		Call Shell(SysInfoPath, AppWinStyle.NormalFocus)
		
		Exit Sub
SysInfoErr: 
		MsgBox("此时系统信息不可用", MsgBoxStyle.OKOnly)
	End Sub
	
	Public Function GetKeyValue(ByRef KeyRoot As Integer, ByRef KeyName As String, ByRef SubKeyRef As String, ByRef KeyVal As String) As Boolean
		Dim i As Integer ' 循环计数器
		Dim rc As Integer ' 返回代码
		Dim hKey As Integer ' 打开的注册表关键字句柄
        Dim KeyValType As Integer ' 注册表关键字数据类型
		Dim tmpVal As String ' 注册表关键字值的临时存储器
		Dim KeyValSize As Integer ' 注册表关键自变量的尺寸
		'------------------------------------------------------------
		' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
		'------------------------------------------------------------
		rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' 处理错误...
		
		tmpVal = New String(Chr(0), 1024) ' 分配变量空间
		KeyValSize = 1024 ' 标记变量尺寸
		
		'------------------------------------------------------------
		' 检索注册表关键字的值...
		'------------------------------------------------------------
		rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize) ' 获得/创建关键字值
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' 处理错误
		
		If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then ' Win95 外接程序空终结字符串...
			tmpVal = VB.Left(tmpVal, KeyValSize - 1) ' Null 被找到,从字符串中分离出来
		Else ' WinNT 没有空终结字符串...
			tmpVal = VB.Left(tmpVal, KeyValSize) ' Null 没有被找到, 分离字符串
		End If
		'------------------------------------------------------------
		' 决定转换的关键字的值类型...
		'------------------------------------------------------------
		Select Case KeyValType ' 搜索数据类型...
			Case REG_SZ ' 字符串注册关键字数据类型
				KeyVal = tmpVal ' 复制字符串的值
			Case REG_DWORD ' 四字节的注册表关键字数据类型
				For i = Len(tmpVal) To 1 Step -1 ' 将每位进行转换
					KeyVal = KeyVal & Hex(Asc(Mid(tmpVal, i, 1))) ' 生成值字符。 By Char。
				Next 
				KeyVal = VB6.Format("&h" & KeyVal) ' 转换四字节的字符为字符串
		End Select
		
		GetKeyValue = True ' 返回成功
		rc = RegCloseKey(hKey) ' 关闭注册表关键字
		Exit Function ' 退出
		
GetKeyError: ' 错误发生后将其清除...
		KeyVal = "" ' 设置返回值到空字符串
		GetKeyValue = False ' 返回失败
		rc = RegCloseKey(hKey) ' 关闭注册表关键字
    End Function

    Private Sub frmAbout1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '由于窗口的显示有两种方式：模态显示（showdialog）和非模态显示（show），本软件用非模态显示，显示前禁用主菜单，结束后应该恢复允许使用主菜单
        zct_main.MainMenu1.Enabled = True
    End Sub
	
	Private Sub frmAbout1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.lblCompanyProduct.Text = sofe_name_l1
        Me.Label2.Text = sofe_name_l2
        Me.lblCompany.Text = sofe_danwei_l1
        Me.Label1.Text = sofe_danwei_l2
        Me.lblWarning.Text = Trim(sofe_time & sofe_ver)
        Me.Text = "关于" & sofe_name
    End Sub
End Class