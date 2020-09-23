Attribute VB_Name = "RegistryAccess"
Option Explicit
' Possible registry data types
Public Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum
' Registry value type definitions
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
' Registry section definitions
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
' Codes returned by Reg API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1
Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const KEY_CREATE_LINK = &H20
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = (KEY_READ)
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))


' This routine allows you to get values from anywhere in the Registry, it currently
' only handles string and double word values.

Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
  On Error Resume Next
    lResult = RegOpenKeyEx(ByVal Group, ByVal Section, ByVal &O0, &H1, lKeyValue)
'    If lResult > 0 Then
'      Dim Buffer As String
'      'Create a string buffer
'      Buffer = Space(200)
'      'Set the error number
'      SetLastError lResult
'      'Format the message string
'      FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal &O0, GetLastError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
'      'Show the message
'      MsgBox Buffer
' End If
 sValue = Space$(2048)
 lValueLength = Len(sValue)
 lResult = RegQueryValueEx(ByVal lKeyValue, ByVal Key, ByVal 0&, ByVal lDataTypeValue, ByVal sValue, lValueLength)
 If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   sValue = Left$(sValue, lValueLength - 1)
Else
  sValue = ""
End If
On Error GoTo 0
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sValue
End Function

' This routine allows you to write values into the entire Registry, it currently
' only handles string and double word values.
Public Sub WriteRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String, ByVal ValType As InTypes, ByVal Value As Variant)
Dim lResult As Long
Dim lKeyValue As Long
Dim InLen As Long
Dim lNewVal As Long
Dim sNewVal As String
On Error Resume Next
 RegCreateKeyEx Group, Section, 0, "REG_SZ", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lKeyValue, lResult
 If ValType = ValDWord Then
   lNewVal = CLng(Value)
   InLen = 4
   lResult = RegSetValueExLong(lKeyValue, Key, 0&, ValType, lNewVal, InLen)
 Else
   sNewVal = Value
   InLen = Len(sNewVal)
   lResult = RegSetValueExString(lKeyValue, Key, 0&, 1&, sNewVal, InLen)
 End If
 lResult = RegFlushKey(lKeyValue)
 lResult = RegCloseKey(lKeyValue)
End Sub


