Attribute VB_Name = "RegistryModule"
Option Explicit

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'Security Mask constants
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))
   
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

'Registry value type definitions
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8

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

'This routine deletes a specified key (and all its subkeys and values if on Win95) from the registry.
'Be very careful using this function.
'
'DeleteSubkey HKEY_CURRENT_USER, "Software\My Name\My App"
Public Function DeleteSubkey(ByVal Group As Long, ByVal section As String) As String
 Dim lResult As Long, lKeyValue As Long
 On Error GoTo DeleteSubkeyError
 lResult = RegOpenKeyEx(Group, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue)
 lResult = RegDeleteKey(lKeyValue, section)
 lResult = RegCloseKey(lKeyValue)
 Exit Function
DeleteSubkeyError:
 Call DisplayError(lResult, "DeleteSubkey")
End Function

Public Function RegKeyExists(ByVal RootKey As Long, ByVal SubKey As String) As Boolean
 On Error Resume Next
 Dim RC As Long 'return code
 Dim hKey As Long 'key handle
 RC = RegOpenKey(RootKey, SubKey, hKey)
 If RC = ERROR_NONE Then
  RegKeyExists = True
  Call RegCloseKey(hKey)
  Exit Function
 ElseIf RC = ERROR_BADKEY Then
  RegKeyExists = False
  Call RegCloseKey(hKey)
  Exit Function
 Else
  Call DisplayError(RC, "RegOpenKeyEx")
 End If
End Function

'This routine allows you to write values into the Registry
Public Sub WriteRegistry(ByVal Group As Long, ByVal section As String, ByVal Key As String, ByVal ValType As InTypes, ByVal Value As String)
 On Error GoTo ErrTrapRegCreate
 Dim result As Long
 Dim handleToKey As Long
 Dim InLen As Long
 Dim lNewVal As Long
 Dim sNewVal As String
 result = RegCreateKeyEx(Group, section, 0&, vbNullString, 0, KEY_ALL_ACCESS, 0&, handleToKey, result)
 If result <> 0 Then GoTo ErrTrapRegCreate
 result = RegSetValueExString(handleToKey, Key, 0, REG_SZ, ByVal Value, Len(Value))
 If result <> 0 Then GoTo ErrTrapRegCreate
 Exit Sub
ErrTrapRegCreate:
 Call DisplayError(result, "WriteRegistry")
End Sub

'This routine allows you to read values from the Registry
Public Function ReadRegistry(ByVal Group As Long, ByVal section As String, ByVal Key As String) As String
 On Error Resume Next
 Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
 Dim TStr1 As String, TStr2 As String
 Dim i As Integer
 lResult = RegOpenKey(Group, section, lKeyValue)
 sValue = Space$(2048)
 lValueLength = Len(sValue)
 lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
 If (lResult = 0) And (Err.Number = 0) Then
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
 lResult = RegCloseKey(lKeyValue)
 ReadRegistry = sValue
End Function

'This routine deletes a specified value from below a specified subkey.
'Be very careful using this function.
'
'DeleteValue HKEY_CURRENT_USER, "Software\My Name\My App", "NewSubKey"
Public Function DeleteValue(ByVal Group As Long, ByVal section As String, ByVal Key As String) As String
 On Error GoTo DeleteValueError
 Dim lResult As Long, lKeyValue As Long
 lResult = RegOpenKey(Group, section, lKeyValue)
 lResult = RegDeleteValue(lKeyValue, Key)
 lResult = RegCloseKey(lKeyValue)
 Exit Function
DeleteValueError:
 Call DisplayError(lResult, "DeleteValue")
End Function

Public Function RegValueExists(ByVal RootKey As Long, ByVal SubKey As String, ByVal ValueName As String) As Boolean
 On Error Resume Next
 Dim RC As Long 'return code
 Dim hKey As Long 'key handle
 Dim lDataTypeValue As Long
 Dim sValue As String
 Dim lValueLength As Long
 RC = RegOpenKey(RootKey, SubKey, hKey)
 If RC = ERROR_NONE Then
  RC = RegQueryValueEx(hKey, ValueName, 0&, lDataTypeValue, sValue, lValueLength)
  If RC = ERROR_NONE Then
   RegValueExists = True
   Call RegCloseKey(hKey)
   Exit Function
  ElseIf RC = ERROR_BADKEY Then
   RegValueExists = False
   Call RegCloseKey(hKey)
   Exit Function
  Else
   Call DisplayError(RC, "RegQueryValueEx")
   Call RegCloseKey(hKey)
   Exit Function
  End If
  Call RegCloseKey(hKey)
 Else
  Call DisplayError(RC, "RegOpenKey")
  Call RegCloseKey(hKey)
 End If
End Function

Public Sub DisplayError(ByVal errNum As Integer, ByVal func As String)
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
