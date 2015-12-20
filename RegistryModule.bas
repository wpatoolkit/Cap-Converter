Attribute VB_Name = "RegistryModule"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
   
'Possible registry data types
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

Public Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
End Type

Public Type ACL
 AclRevision As Byte
 Sbz1 As Byte
 AclSize As Integer
 AceCount As Integer
 Sbz2 As Integer
End Type

Public Type SECURITY_DESCRIPTOR
 Revision As Byte
 Sbz1 As Byte
 Control As Long
 Owner As Long
 Group As Long
 Sacl As ACL
 Dacl As ACL
End Type

Public Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type

'Reg Create Type Values
Public Const REG_OPTION_RESERVED As Long = 0           ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE As Long = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE As Long = 1           ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK As Long = 2        ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE As Long = 4     ' open for backup or restore

'Reg Data Types
Public Const REG_NONE As Long = 0                       ' No value type
Public Const REG_SZ As Long = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ As Long = 2                  ' Unicode nul terminated string
Public Const REG_BINARY As Long = 3                     ' Free form binary
Public Const REG_DWORD As Long = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN As Long = 5           ' 32-bit number
Public Const REG_LINK As Long = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ As Long = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST As Long = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
Public Const REG_CREATED_NEW_KEY As Long = &H1                      ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY As Long = &H2                      ' Existing Key opened
Public Const REG_WHOLE_HIVE_VOLATILE As Long = &H1                      ' Restore whole hive volatile
Public Const REG_REFRESH_HIVE As Long = &H2                      ' Unwind changes to last flush
Public Const REG_NOTIFY_CHANGE_NAME As Long = &H1                      ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES As Long = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET As Long = &H4                      ' Time stamp
Public Const REG_NOTIFY_CHANGE_SECURITY As Long = &H8
Public Const REG_LEGAL_CHANGE_FILTER As Long = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION As Long = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

'Security Mask constants
Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Reg Key Security Options
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

'Registry section definitions
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

'Codes returned by Reg API calls
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

'Wrapper for the RegDeleteKey Win32 API
Public Sub RegDeleteSubKey(ByVal Group As Long, ByVal section As String)
 On Error GoTo RegDeleteSubKeyError
 Dim RC As Long, hKey As Long
 RC = RegOpenKeyEx(Group, vbNullChar, 0&, KEY_ALL_ACCESS, hKey)
 RC = RegDeleteKey(hKey, section)
 RC = RegCloseKey(hKey)
 Exit Sub
RegDeleteSubKeyError:
 Call DisplayRegError(RC, "RegDeleteSubKey")
End Sub

'Wrapper for the RegDeleteValue Win32 API
Public Sub DeleteValue(ByVal Group As Long, ByVal section As String, ByVal Key As String)
 On Error GoTo DeleteValueError
 Dim RC As Long, hKey As Long
 RC = RegOpenKey(Group, section, hKey)
 RC = RegDeleteValue(hKey, Key)
 RC = RegCloseKey(hKey)
 Exit Sub
DeleteValueError:
 Call DisplayRegError(RC, "DeleteValue")
End Sub

'Wrapper for the RegOpenKeyEx and RegCloseKey Win32 API
Public Function RegKeyExists(ByVal RootKey As Long, ByVal SubKey As String) As Boolean
 On Error GoTo RegKeyExistsError
 Dim RC As Long 'return code
 Dim hKey As Long 'key handle
 RC = RegOpenKeyEx(RootKey, SubKey, 0&, KEY_QUERY_VALUE, hKey)
 If RC = ERROR_NONE Then
  Call RegCloseKey(hKey)
  RegKeyExists = True
  Exit Function
 ElseIf RC = ERROR_BADKEY Then
  Call RegCloseKey(hKey)
  RegKeyExists = False
  Exit Function
 Else
  Call DisplayRegError(RC, "RegOpenKeyEx")
 End If
 Exit Function
RegKeyExistsError:
 Call DisplayRegError(RC, "RegKeyExists")
End Function

'Wrapper for the RegOpenKeyEx, RegQueryValueEx, and RegCloseKey Win32 API
Public Function RegValueExists(ByVal RootKey As Long, ByVal SubKey As String, ByVal ValueName As String) As Boolean
 On Error GoTo RegValueExistsError
 If (RegKeyExists(RootKey, SubKey) = False) Then
  RegValueExists = False
  Exit Function
 End If
 Dim RC As Long 'return code
 Dim hKey As Long 'key handle
 Dim lDataTypeValue As Long
 Dim sValue As String
 Dim lValueLength As Long
 RC = RegOpenKeyEx(RootKey, SubKey, 0&, KEY_QUERY_VALUE, hKey)
 If RC = ERROR_NONE Then
  RC = RegQueryValueEx(hKey, ValueName, 0&, lDataTypeValue, sValue, lValueLength)
  If RC = ERROR_NONE Then
   Call RegCloseKey(hKey)
   RegValueExists = True
   Exit Function
  ElseIf RC = ERROR_BADKEY Then
   Call RegCloseKey(hKey)
   RegValueExists = False
   Exit Function
  Else
   Call DisplayRegError(RC, "RegQueryValueEx")
   Call RegCloseKey(hKey)
   Exit Function
  End If
  Call RegCloseKey(hKey)
 Else
  Call DisplayRegError(RC, "RegOpenKeyEx")
  Call RegCloseKey(hKey)
 End If
 Exit Function
RegValueExistsError:
 Call DisplayRegError(RC, "RegValueExists")
 RegValueExists = False
End Function

'Wrapper for the RegCreateKeyEx and RegSetValueEx Win32 API
Public Sub WriteRegistry(ByVal Group As Long, ByVal section As String, ByVal Key As String, ByVal ValType As InTypes, ByVal Value As String)
 On Error GoTo WriteRegistryError
 Dim RC As Long
 Dim hKey As Long
 Dim lDisp As Long
 RC = RegCreateKeyEx(Group, section, 0&, vbNullString, 0&, KEY_ALL_ACCESS, 0&, hKey, lDisp)
 If RC <> 0 Then GoTo WriteRegistryError
 RC = RegSetValueExString(hKey, Key, 0&, REG_SZ, ByVal Value, Len(Value))
 If RC <> 0 Then GoTo WriteRegistryError
 Exit Sub
WriteRegistryError:
 Call DisplayRegError(RC, "WriteRegistry")
End Sub

'Wrapper for the RegOpenKey, RegQueryValueEx, and RegCloseKey Win32 API
Public Function ReadRegistry(ByVal Group As Long, ByVal section As String, ByVal Key As String) As String
 On Error Resume Next
 Dim RC As Long, hKey As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
 Dim TStr1 As String, TStr2 As String
 Dim i As Integer
 RC = RegOpenKey(Group, section, hKey)
 If (RC = 0) And (Err.Number = 0) Then
  sValue = Space$(2048)
  lValueLength = Len(sValue)
  RC = RegQueryValueEx(hKey, Key, 0&, lDataTypeValue, sValue, lValueLength)
  If (RC = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
    td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
    sValue = Format$(td, "000")
   End If
   If lDataTypeValue = REG_BINARY Then
    'Return a binary field as a hex string (2 chars per byte)
    TStr2 = ""
    For i = 1 To lValueLength
     TStr1 = Hex(Asc(Mid(sValue, i, 1)))
     If Len(TStr1) = 1 Then TStr1 = "0" & TStr1
     TStr2 = TStr2 + TStr1
    Next
    sValue = TStr2
   Else
    sValue = Left$(sValue, lValueLength - 1)
   End If
  Else
   sValue = ""
  End If
 End If
 RC = RegCloseKey(hKey)
 ReadRegistry = sValue
End Function

Public Sub DisplayRegError(ByVal errNum As Integer, ByVal func As String)
 Select Case errNum
  Case 0
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_NONE" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 1
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_BADDB" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 2
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_BADKEY" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 3
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_CANTOPEN" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 4
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_CANTREAD" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 5
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_CANTWRITE" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 6
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_OUTOFMEMORY" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 7
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_INVALID_PARAMETER" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 8
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_ACCESS_DENIED" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 87
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_INVALID_PARAMETERS" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case 259
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: ERROR_NO_MORE_ITEMS" & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
  Case Else
   MsgBox "Error Accessing the Registry." & vbNewLine & vbNewLine & "Error Number: " & errNum & vbNewLine & "Error Description: Unknown." & vbNewLine & "Calling Function: " & func, vbExclamation + vbOKOnly
 End Select
End Sub
