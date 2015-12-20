VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Cap Converter"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   6990
      Width           =   2295
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "&Convert"
      Default         =   -1  'True
      Height          =   735
      Left            =   960
      TabIndex        =   9
      Top             =   6990
      Width           =   2295
   End
   Begin VB.CommandButton btnOutputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   1410
      Width           =   495
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton btnInputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtInputFile 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label lblOutputFile 
      Caption         =   "Output File:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   2730
   End
   Begin VB.Label lblDirectoryMode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Directory Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3405
      TabIndex        =   2
      Top             =   210
      Width           =   1455
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " | "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2985
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFileMode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "File Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      TabIndex        =   0
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblInputFile 
      Caption         =   "Input File:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2370
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Output Options:
'[ ] Extract first found handshake only
'[ ] Combine all found handshakes into single hccap
'[ ] Create separate hccaps for each found handshake
'
'File Association Options:
'[ ] Add to context menu of .HCCAP files
'[ ] Add to context menu of .CAP files
'[ ] Set as default application for .HCCAP files
'[ ] Set as default application for .CAP files

Private Sub btnInputFile_Click()
If lblFileMode.ForeColor = &H80000012 Then 'Browse for File
 Dim file_to_open As String
 Dim OFN As OPENFILENAME
 OFN.lStructSize = Len(OFN)
 OFN.hwndOwner = Me.hwnd
 OFN.hInstance = App.hInstance
 OFN.lpstrTitle = "Choose Input File"
 OFN.lpstrFilter = "CAP Files (*.cap;*.pcap;*.dmp)" & Chr$(0) & "*.cap;*.pcap;*.dmp" & Chr$(0) & "HCCAP Files (*.hccap)" & Chr$(0) & "*.hccap" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
 OFN.lpstrDefExt = "cap"
 OFN.nFilterIndex = 1
 OFN.lpstrFile = String(257, 0)
 OFN.nMaxFile = Len(OFN.lpstrFile) - 1
 OFN.lpstrFileTitle = OFN.lpstrFile
 OFN.nMaxFileTitle = OFN.nMaxFile
 OFN.lpstrInitialDir = IIf((last_path <> ""), last_path, App.path)
 OFN.flags = 0
 If GetOpenFileName(OFN) Then
  last_path = get_path_from_file(Trim$(OFN.lpstrFile))
  file_to_open = IIf(is_file(Trim$(OFN.lpstrFile)), Trim$(OFN.lpstrFile), "")
 End If
 If (file_to_open <> "") Then
  txtInputFile.Text = file_to_open
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  last_path = fso.GetParentFolderName(txtInputFile.Text)
  txtOutputFile.Text = fso.GetParentFolderName(txtInputFile.Text) & "\" & fso.GetBaseName(txtInputFile.Text) & "." & IIf(fso.GetExtensionName(txtInputFile.Text) = "hccap", "cap", "hccap")
  Set fso = Nothing
 End If
Else 'Browse for Directory
 Dim tBrowseInfo As BrowseInfo
 tBrowseInfo.hwndOwner = Me.hwnd
 tBrowseInfo.lpszTitle = lstrcat("Choose Input Directory", "")
 tBrowseInfo.ulFlags = 1 + 2 + &H4&
 Dim tmpLong As Long
 tmpLong = SHBrowseForFolder(tBrowseInfo)
 If (tmpLong) Then
  Dim sBuffer As String
  sBuffer = Space(260)
  SHGetPathFromIDList tmpLong, sBuffer
  sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  txtInputFile.Text = sBuffer
  txtOutputFile.Text = txtInputFile.Text & IIf(Right$(txtInputFile.Text, 1) <> "\", "\", "") & "HCCAPS"
 End If
End If
End Sub

Private Sub btnOutputFile_Click()
If lblFileMode.ForeColor = &H80000012 Then 'Browse for File
 Dim file_to_save As String
 Dim OFN As OPENFILENAME
 OFN.lStructSize = Len(OFN)
 OFN.hwndOwner = Me.hwnd
 OFN.hInstance = App.hInstance
 OFN.lpstrTitle = "Save File As"
 OFN.lpstrFilter = "HCCAP Files (*.hccap)" & Chr$(0) & "*.hccap" & Chr$(0) & "CAP Files (*.cap;*.pcap;*.dmp)" & Chr$(0) & "*.cap;*.pcap;*.dmp" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
 OFN.lpstrDefExt = "hccap"
 OFN.nFilterIndex = 1
 OFN.lpstrFile = String(257, 0)
 OFN.nMaxFile = Len(OFN.lpstrFile) - 1
 OFN.lpstrFileTitle = OFN.lpstrFile
 OFN.nMaxFileTitle = OFN.nMaxFile
 OFN.lpstrInitialDir = IIf((last_path <> ""), last_path, App.path)
 OFN.flags = 0
 If GetSaveFileName(OFN) Then
  last_path = get_path_from_file(Trim$(OFN.lpstrFile))
  file_to_save = Trim$(OFN.lpstrFile)
 End If
 If (file_to_save <> "") Then
  txtOutputFile.Text = file_to_save
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  last_path = fso.GetParentFolderName(file_to_save)
  Set fso = Nothing
 End If
Else
 Dim tBrowseInfo As BrowseInfo
 tBrowseInfo.hwndOwner = Me.hwnd
 tBrowseInfo.lpszTitle = lstrcat("Choose Output Directory", "")
 tBrowseInfo.ulFlags = 1 + 2 + &H4&
 Dim tmpLong As Long
 tmpLong = SHBrowseForFolder(tBrowseInfo)
 If (tmpLong) Then
  Dim sBuffer As String
  sBuffer = Space(260)
  SHGetPathFromIDList tmpLong, sBuffer
  sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  txtOutputFile.Text = sBuffer
 End If
End If
End Sub

Private Sub Form_Load()
 RemoveMenu GetSystemMenu(Me.hwnd, 0), 2, &H400& 'prevent resizing
 
 'command line args
 If (Command$ <> "") Then
  Dim cmdline_file As String
  cmdline_file = Command$
  If Mid$(cmdline_file, 1, 1) = """" Then
   cmdline_file = Mid$(cmdline_file, 2) 'remove starting quote
  End If
  If Mid$(cmdline_file, Len(cmdline_file), 1) = """" Then
   cmdline_file = Mid$(cmdline_file, 1, Len(cmdline_file) - 1) 'remove ending quote
  End If
  If InStr(cmdline_file, "\") = 0 Then
   cmdline_file = App.path & IIf(Right$(App.path, 1) <> "\", "\", "") & cmdline_file 'prefix starting path
  End If
  If (is_file(cmdline_file)) Then 'file exists
   Dim tmp_hccap_record As hccap_record
   Dim fso As Scripting.FileSystemObject
   Set fso = New Scripting.FileSystemObject
   If Right$(cmdline_file, 6) = ".hccap" Then
     txtInputFile.Text = cmdline_file
     txtOutputFile.Text = fso.GetParentFolderName(txtInputFile.Text) & "\" & fso.GetBaseName(txtInputFile.Text) & ".cap"
   ElseIf Right$(cmdline_file, 4) = ".cap" Or Right$(cmdline_file, 5) = ".pcap" Or Right$(cmdline_file, 4) = ".dmp" Then
     txtInputFile.Text = cmdline_file
     txtOutputFile.Text = fso.GetParentFolderName(txtInputFile.Text) & "\" & fso.GetBaseName(txtInputFile.Text) & ".hccap"
   End If
   Set fso = Nothing
  End If
 End If
End Sub

Private Sub lblDirectoryMode_Click()
 If lblDirectoryMode.ForeColor = &HFF0000 Then 'blue
  lblDirectoryMode.ForeColor = &H80000012 'black
  lblDirectoryMode.FontUnderline = False
  lblFileMode.ForeColor = &HFF0000 'blue
  lblFileMode.FontUnderline = True
  lblInputFile.caption = "Input Directory:"
  lblOutputFile.caption = "Output Directory:"
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  If (is_file(txtInputFile.Text) = True) Then
   txtInputFile.Text = fso.GetParentFolderName(txtInputFile.Text)
  End If
  If (is_folder(txtInputFile.Text) = True) Then
   txtOutputFile.Text = txtInputFile.Text & IIf(Right$(txtInputFile.Text, 1) <> "\", "\", "") & "HCCAPS"
  End If
  Set fso = Nothing
 End If
End Sub

Private Sub lblFileMode_Click()
 If lblFileMode.ForeColor = &HFF0000 Then 'blue
  lblFileMode.ForeColor = &H80000012 'black
  lblFileMode.FontUnderline = False
  lblDirectoryMode.ForeColor = &HFF0000 'blue
  lblDirectoryMode.FontUnderline = True
  lblInputFile.caption = "Input File:"
  lblOutputFile.caption = "Output File:"
  txtInputFile.Text = ""
  txtOutputFile.Text = ""
 End If
End Sub

Private Sub lblDirectoryMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If lblDirectoryMode.ForeColor = &HFF0000 Then
  SetCursor LoadCursor(0, 32649&)
 End If
End Sub

Private Sub lblFileMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblFileMode.ForeColor = &HFF0000 Then
  SetCursor LoadCursor(0, 32649&)
 End If
End Sub

Private Sub btnExit_Click()
 Unload Me
End Sub
