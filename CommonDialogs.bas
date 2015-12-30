Attribute VB_Name = "CommonDialogs"
Option Explicit
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As OPENFILENAME) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Public Type OPENFILENAME
 lStructSize As Long         'The size of this struct (Use the Len function)
 hwndOwner As Long           'The hWnd of the owner window. The dialog will be modal to this window
 hInstance As Long           'The instance of the calling thread. You can use the App.hInstance here.
 lpstrFilter As String       'Use this to filter what files are shown in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
 lpstrCustomFilter As String 'The pattern the user has chosen is saved here if you pass a non empty string. I never use this one
 nMaxCustFilter As Long      'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
 nFilterIndex As Long        'What filter (of lpstrFilter) is showed when the user opens the dialog.
 lpstrFile As String         'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
 nMaxFile As Long            'The length of lpstrFile + 1
 lpstrFileTitle As String    'The name of the file. Should be MAX_PATH character long
 nMaxFileTitle As Long       'The length of lpstrFileTitle + 1
 lpstrInitialDir As String   'The path to the initial path. If you pass an empty string the initial path is the current path.
 lpstrTitle As String        'The caption of the dialog.
 FLAGS As Long               'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
 nFileOffset As Integer      'Points to the what character in lpstrFile where the actual filename begins (zero based)
 nFileExtension As Integer   'Same as nFileOffset except that it points to the file extention.
 lpstrDefExt As String       'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
 lCustData As Long           'Only used if you provide a Hook procedure
 lpfnHook As Long            'Pointer to the hook procedure
 lpTemplateName As String    'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

Public Type OFSTRUCT
 cBytes As Byte
 fFixedDisk As Byte
 nErrCode As Integer
 Reserved1 As Integer
 Reserved2 As Integer
 szPathName(128&) As Byte
End Type

Public Const OF_READ As Long = &H0 'Opens a file for reading only.
Public Const OF_WRITE As Long = &H1 'Opens a file for write access only.
Public Const OF_READWRITE As Long = &H2 'Opens a file with read/write permissions.
Public Const OF_SHARE_COMPAT As Long = &H0 'For MS-DOS–based file systems, opens a file with compatibility mode, allows any process on a specified computer to open the file any number of times.
Public Const OF_SHARE_DENY_NONE As Long = &H40 'Opens a file without denying read or write access to other processes. On MS-DOS-based file systems, if the file has been opened in compatibility mode by any other process, the function fails.
Public Const OF_SHARE_DENY_READ As Long = &H30 'Opens a file and denies read access to other processes. On MS-DOS-based file systems, if the file has been opened in compatibility mode, or for read access by any other process, the function fails.
Public Const OF_SHARE_DENY_WRITE As Long = &H20 'Opens a file and denies write access to other processes. On MS-DOS-based file systems, if a file has been opened in compatibility mode, or for write access by any other process, the function fails.
Public Const OF_SHARE_EXCLUSIVE As Long = &H10 'Opens a file with exclusive mode, and denies both read/write access to other processes. If a file has been opened in any other mode for read/write access, even by the current process, the function fails.
Public Const OF_CANCEL As Long = &H800 'Ignored. To produce a dialog box containing a Cancel button, use OF_PROMPT.
Public Const OF_CREATE As Long = &H1000 'Creates a new file. If the file exists, it is truncated to zero (0) length.
Public Const OF_DELETE As Long = &H200 'Deletes a file.
Public Const OF_EXIST As Long = &H4000 'Opens a file and then closes it. Use this to test for the existence of a file.
Public Const OF_PARSE As Long = &H100 'Fills the OFSTRUCT structure, but does not do anything else.
Public Const OF_PROMPT As Long = &H2000 'Displays a dialog box if a requested file does not exist. A dialog box informs a user that the system cannot find a file, and it contains Retry and Cancel buttons. The Cancel button directs OpenFile to return a file-not-found error message.
Public Const OF_REOPEN As Long = &H8000 'Opens a file by using information in the reopen buffer.
Public Const OF_VERIFY As Long = &H400 'Verifies that the date and time of a file are the same as when it was opened previously. This is useful as an extra check for read-only files.
Public Const HFILE_ERROR As Long = -1&

Public Type BrowseInfo
 hwndOwner      As Long
 pIDLRoot       As Long
 pszDisplayName As Long
 lpszTitle      As Long
 ulFlags        As Long
 lpfnCallback   As Long
 lParam         As Long
 iImage         As Long
End Type

Public Function file_exists(ByVal file_name As String) As Boolean
 On Error Resume Next
 Dim lResult As Long
 Dim open_file_struct As OFSTRUCT
 lResult = OpenFile(file_name, open_file_struct, OF_EXIST)
 file_exists = IIf(lResult <> -1&, True, False)
End Function

Public Function is_file(str As String) As Boolean
 Dim fso As Scripting.FileSystemObject
 Set fso = New Scripting.FileSystemObject
 Dim return_value As Boolean
 return_value = fso.FileExists(str)
 Set fso = Nothing
 is_file = return_value
End Function

Public Function is_folder(str As String) As Boolean
 Dim fso As Scripting.FileSystemObject
 Set fso = New Scripting.FileSystemObject
 Dim return_value As Boolean
 return_value = fso.FolderExists(str)
 Set fso = Nothing
 is_folder = return_value
End Function

Public Function get_path_from_file(fileName As String) As String
 Dim pos As Integer
 pos = InStrRev(fileName, "\")
 If pos > 0 Then
  get_path_from_file = Left$(fileName, pos)
 Else
  get_path_from_file = ""
 End If
End Function

Public Function get_file_from_path(fileName As String) As String
 fileName = Left$(fileName, lstrlen(StrPtr(fileName)))
 Dim pos As Integer
 pos = InStrRev(fileName, "\")
 If pos > 0 Then
  get_file_from_path = Right$(fileName, Len(fileName) - pos)
 Else
  get_file_from_path = ""
 End If
End Function

Public Function ShowOpenFileDialog(owner_hwnd As Long, caption As String) As String
 Dim OFN As OPENFILENAME
 OFN.lStructSize = Len(OFN)
 OFN.hwndOwner = owner_hwnd
 OFN.hInstance = App.hInstance
 OFN.lpstrTitle = IIf(caption <> "", caption, "Choose File")
 If (InStr(caption, "HCCAP") > 0) Then
  OFN.lpstrFilter = "HCCAP Files (*.hccap)" & Chr$(0) & "*.hccap" & Chr$(0)
  OFN.lpstrDefExt = "hccap"
 ElseIf (InStr(caption, "CAP") > 0) Then
  OFN.lpstrFilter = "CAP Files (*.cap;*.pcap;*.dmp)" & Chr$(0) & "*.cap;*.pcap;*.dmp" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
  OFN.lpstrDefExt = "cap"
 Else
  OFN.lpstrFilter = "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
 End If
 OFN.nFilterIndex = 1
 OFN.lpstrFile = String(257, 0)
 OFN.nMaxFile = Len(OFN.lpstrFile) - 1
 OFN.lpstrFileTitle = OFN.lpstrFile
 OFN.nMaxFileTitle = OFN.nMaxFile
 OFN.lpstrInitialDir = IIf((last_path <> ""), last_path, App.path)
 OFN.FLAGS = 0
 If GetOpenFileName(OFN) Then
  last_path = get_path_from_file(Trim$(OFN.lpstrFile))
  ShowOpenFileDialog = IIf(is_file(Trim$(OFN.lpstrFile)), Trim$(OFN.lpstrFile), "")
 Else
  ShowOpenFileDialog = ""
 End If
End Function

Public Function ShowSaveFileDialog(owner_hwnd As Long, caption As String) As String
 Dim OFN As OPENFILENAME
 OFN.lStructSize = Len(OFN)
 OFN.hwndOwner = owner_hwnd
 OFN.hInstance = App.hInstance
 OFN.lpstrTitle = IIf(caption <> "", caption, "Save File As")
 If (InStr(caption, "HCCAP") > 0) Then
  OFN.lpstrFilter = "HCCAP Files (*.hccap)" & Chr$(0) & "*.hccap" & Chr$(0)
  OFN.lpstrDefExt = "hccap"
 ElseIf (InStr(caption, "CAP") > 0) Then
  OFN.lpstrFilter = "CAP Files (*.cap;*.pcap;*.dmp)" & Chr$(0) & "*.cap;*.pcap;*.dmp" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
  OFN.lpstrDefExt = "cap"
 Else
  OFN.lpstrFilter = "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
 End If
 OFN.nFilterIndex = 1
 OFN.lpstrFile = String(257, 0)
 OFN.nMaxFile = Len(OFN.lpstrFile) - 1
 OFN.lpstrFileTitle = OFN.lpstrFile
 OFN.nMaxFileTitle = OFN.nMaxFile
 OFN.lpstrInitialDir = IIf((last_path <> ""), last_path, App.path)
 OFN.FLAGS = 0
 If GetSaveFileName(OFN) Then
  last_path = get_path_from_file(Trim$(OFN.lpstrFile))
  ShowSaveFileDialog = Trim$(OFN.lpstrFile)
 Else
  ShowSaveFileDialog = ""
 End If
End Function
