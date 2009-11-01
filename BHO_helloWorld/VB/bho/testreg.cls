VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (dest As Any, source As Any, ByVal numBytes As Long)

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
   (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
   
'//注册表 API 函数声明
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   ByVal lpbData As Any, ByVal cbData As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
   (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, _
   ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
   phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
   lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
   lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, _
   ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
   lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
   (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
   (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal ipValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Byte, ByVal cbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
   ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
   lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValueInt Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long
    
'//注册表结构
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

'//注册表访问权
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = &H3F

'//打开/建立选项
Const REG_OPTION_NON_VOLATILE = 0&
Const REG_OPTION_VOLATILE = &H1

'//Key 创建/打开
Const REG_CREATED_NEW_KEY = &H1
Const REG_OPENED_EXISTING_KEY = &H2

'//预定义存取类型
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF

'//严格代码定义
Const ERROR_SUCCESS = 0&
Const ERROR_ACCESS_DENIED = 5
Const ERROR_NO_MORE_ITEMS = 259
Const ERROR_MORE_DATA = 234 '//  错误

'//注册表值类型列举
Private Enum RegDataTypeEnum
'   REG_NONE = (0)                         '// No value type
   REG_SZ = (1)                           '// Unicode nul terminated string
   REG_EXPAND_SZ = (2)                    '// Unicode nul terminated string w/enviornment var
   REG_BINARY = (3)                       '// Free form binary
   REG_DWORD = (4)                        '// 32-bit number
   REG_DWORD_LITTLE_ENDIAN = (4)          '// 32-bit number (same as REG_DWORD)
   REG_DWORD_BIG_ENDIAN = (5)             '// 32-bit number
'   REG_LINK = (6)                         '// Symbolic Link (unicode)
   REG_MULTI_SZ = (7)                     '// Multiple, null-delimited, double-null-terminated Unicode strings
'   REG_RESOURCE_LIST = (8)                '// Resource list in the resource map
'   REG_FULL_RESOURCE_DESCRIPTOR = (9)     '// Resource list in the hardware description
'   REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum
   
'//注册表基本键值列表
Public Enum RootKeyEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA_WIN2K_ONLY = &H80000004 '//仅Win2k
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

'// for specifying the type of data to save
Public Enum RegValueTypes
   eInteger = vbInteger
   eLong = vbLong
   eString = vbString
   eByteArray = vbArray + vbByte
End Enum

'//保存时指定类型
Public Enum RegFlags
   IsExpandableString = 1
   IsMultiString = 2
   'IsBigEndian = 3 '// 无指针同样不要设置大Endian值
End Enum

Private Const ERR_NONE = 0


Function SetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, _
   ByVal ValueName As String, ByVal Value As Variant, valueType As RegValueTypes, _
   Optional Flag As RegFlags = 0) As Boolean
   
   Dim handle As Long
   Dim lngValue As Long
   Dim strValue As String
   Dim binValue() As Byte
   Dim length As Long
   Dim retVal As Long
   
   Dim SecAttr As SECURITY_ATTRIBUTES '//键的安全设置
   '//设置新键值的名称和默认安全设置
   SecAttr.nLength = Len(SecAttr) '//结构大小
   SecAttr.lpSecurityDescriptor = 0 '//默认安全权限
   SecAttr.bInheritHandle = True '//设置的默认值

   '// 打开或创建键
   'If RegOpenKeyEx(hKey, KeyName, 0, KEY_ALL_ACCESS, handle) Then Exit Function
   retVal = RegCreateKeyEx(hKey, KeyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, handle, retVal)
   If retVal Then Exit Function

   '//3种数据类型
   Select Case VarType(Value)
      Case vbByte, vbInteger, vbLong '// 若是字节, Integer值或Long值...
         lngValue = Value
         retVal = RegSetValueExLong(handle, ValueName, 0, REG_DWORD, lngValue, Len(lngValue))
      
      Case vbString '// 字符串, 扩展环境字符串或多段字符串...
         strValue = Value
         Select Case Flag
            Case IsExpandableString
               retVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, 255)
            Case IsMultiString
               retVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, 255)
            Case Else '// 正常 REG_SZ 字符串
               retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, 255)
         End Select
      
      Case vbArray + vbByte '// 如果是字节数组...
         binValue = Value
         length = UBound(binValue) - LBound(binValue) + 1
         retVal = RegSetValueExByte(handle, ValueName, 0, REG_BINARY, binValue(0), length)
      
      Case Else '// 如果其它类型
         RegCloseKey handle
         'Err.Raise 1001, , "不支持的值类型"
   
   End Select

   '// 返回关闭结果
   RegCloseKey handle
   
   '// 返回写入成功结果
   SetRegistryValue = (retVal = 0)

End Function


Function GetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, _
   ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
   
   Dim handle As Long
   Dim resLong As Long
   Dim resString As String
   Dim resBinary() As Byte
   Dim length As Long
   Dim retVal As Long
   Dim valueType As Long

   Const KEY_READ = &H20019
   
   '// 默认结果
   GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
   
   '// 打开键, 不存在则退出
   If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
   
   '// 准备 1K  resBinary 用于接收
   length = 1024
   ReDim resBinary(0 To length - 1) As Byte
   
   '// 读注册表值
   retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
   
   '// 若resBinary 太小则重读
   If retVal = ERROR_MORE_DATA Then
      '// resBinary放大,且重新读取
      ReDim resBinary(0 To length - 1) As Byte
      retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
      length)
   End If
   
   '// 返回相应值类型
   Select Case valueType
      Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
         '// REG_DWORD 和 REG_DWORD_LITTLE_ENDIAN 相同
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = resLong
      
      Case REG_DWORD_BIG_ENDIAN
         '// Big Endian's 用在非-Windows环境, 如Unix系统, 本地计算机远程访问
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = SwapEndian(resLong)
      
      Case REG_SZ, REG_EXPAND_SZ
         resString = Space$(length - 1)
         CopyMemory ByVal resString, resBinary(0), length - 1
         If valueType = REG_EXPAND_SZ Then
            '// 查询对应的环境变量
            GetRegistryValue = ExpandEnvStr(resString)
         Else
            GetRegistryValue = resString
         End If

      Case REG_MULTI_SZ
         '// 复制时需指定2个空格符
         resString = Space$(length - 2)
         CopyMemory ByVal resString, resBinary(0), length - 2
         GetRegistryValue = resString

      Case Else ' 包含 REG_BINARY
         '// resBinary 调整
         If length <> UBound(resBinary) + 1 Then
            ReDim Preserve resBinary(0 To length - 1) As Byte
         End If
      GetRegistryValue = resBinary()
   
   End Select
   
   '// 关闭
   RegCloseKey handle

End Function


Public Function DeleteRegistryValueOrKey(ByVal hKey As RootKeyEnum, RegKeyName As String, _
   ValueName As String) As Boolean
'//删除注册表值和键,如果成功返回True

   Dim lRetval As Long      '//打开和输出注册表键的返回值
   Dim lRegHWND As Long     '//打开注册表键的句柄
   Dim sREGSZData As String '//把获取值放入缓冲区
   Dim lSLength As Long     '//缓冲区大小.  改变缓冲区大小要在调用之后
   
   '//打开键
   lRetval = RegOpenKeyEx(hKey, RegKeyName, 0, KEY_ALL_ACCESS, lRegHWND)
   
   '//成功打开
   If lRetval = ERR_NONE Then
      '//删除指定值
      lRetval = RegDeleteValue(lRegHWND, ValueName)  '//如果已存在则先删除
      
      '//如出现错误则删除值并返回False
      If lRetval <> ERR_NONE Then Exit Function
      
      '//注意: 如果成功打开仅关闭注册表键
      lRetval = RegCloseKey(lRegHWND)
     
      '//如成功关闭则返回 True 或者其它错误
      If lRetval = ERR_NONE Then DeleteRegistryValueOrKey = True
      
   End If

End Function


Private Function ExpandEnvStr(sData As String) As String
'// 查询环境变量和返回定义值
'// 如： %PATH% 则返回 "c:\;c:\windows;"

   Dim c As Long, s As String
   
   s = "" '// 不支持Windows 95
   
   '// get the length
   c = ExpandEnvironmentStrings(sData, s, c)
   
   '// 展开字符串
   s = String$(c - 1, 0)
   c = ExpandEnvironmentStrings(sData, s, c)
   
   '// 返回环境变量
   ExpandEnvStr = s
   
End Function


Private Function SwapEndian(ByVal dw As Long) As Long
'// 转换大DWord 到小 DWord
   
   CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
   CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1

End Function





