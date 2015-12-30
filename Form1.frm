VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cap Converter"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10335
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrev 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Top             =   375
      Width           =   495
   End
   Begin VB.CommandButton btnNext 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   23
      Top             =   375
      Width           =   495
   End
   Begin VB.CommandButton btnWriteCAP 
      Caption         =   "Save As CAP..."
      Height          =   645
      Left            =   3810
      TabIndex        =   2
      Top             =   9465
      Width           =   1770
   End
   Begin VB.CommandButton btnReadHCCAP 
      Caption         =   "Open HCCAP..."
      Height          =   645
      Left            =   2010
      TabIndex        =   1
      Top             =   9465
      Width           =   1770
   End
   Begin VB.CommandButton btnWriteHCCAP 
      Caption         =   "Save As HCCAP..."
      Height          =   645
      Left            =   5610
      TabIndex        =   3
      Top             =   9465
      Width           =   1770
   End
   Begin VB.CommandButton btnReadCAP 
      Caption         =   "Open CAP..."
      Height          =   645
      Left            =   210
      TabIndex        =   0
      Top             =   9465
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   " HCCAP Info "
      Height          =   9045
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   225
      Width           =   7125
      Begin VB.TextBox txtKEYMIC 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   47
         TabIndex        =   22
         Top             =   8370
         Width           =   6615
      End
      Begin VB.TextBox txtKEYVER 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   5
         TabIndex        =   20
         Top             =   7650
         Width           =   6615
      End
      Begin VB.TextBox txtEAPOLSIZE 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   3
         TabIndex        =   18
         Top             =   6930
         Width           =   6615
      End
      Begin VB.TextBox txtEAPOL 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   240
         MaxLength       =   782
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   4650
         Width           =   6615
      End
      Begin VB.TextBox txtANONCE 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         MaxLength       =   96
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3690
         Width           =   6615
      End
      Begin VB.TextBox txtSNONCE 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         MaxLength       =   96
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2730
         Width           =   6615
      End
      Begin VB.TextBox txtSTA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   17
         TabIndex        =   10
         Top             =   2010
         Width           =   6615
      End
      Begin VB.TextBox txtBSSID 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   17
         TabIndex        =   8
         Top             =   1290
         Width           =   6615
      End
      Begin VB.TextBox txtESSID 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   36
         TabIndex        =   6
         Top             =   570
         Width           =   6615
      End
      Begin VB.Label lblCounter 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   195
         Left            =   2760
         TabIndex        =   25
         Top             =   255
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblKEYMIC 
         AutoSize        =   -1  'True
         Caption         =   "Key MIC (16 bytes) (Bytes 376 to 391):"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   8130
         Width           =   2715
      End
      Begin VB.Label lblKEYVER 
         AutoSize        =   -1  'True
         Caption         =   "Key Version (4 bytes) (Bytes 372 to 375):"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   7410
         Width           =   2865
      End
      Begin VB.Label lblEAPOLSIZE 
         AutoSize        =   -1  'True
         Caption         =   "EAPOL Size (4 bytes) (Bytes 368 to 371):"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   6690
         Width           =   2895
      End
      Begin VB.Label lblEAPOL 
         AutoSize        =   -1  'True
         Caption         =   "EAPOL (256 bytes) (Bytes 112 to 367):"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   4410
         Width           =   2730
      End
      Begin VB.Label lblANONCE 
         AutoSize        =   -1  'True
         Caption         =   "ANONCE (32 bytes) (Bytes 80 to 111):"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   3450
         Width           =   2700
      End
      Begin VB.Label lblSNONCE 
         AutoSize        =   -1  'True
         Caption         =   "SNONCE (32 bytes) (Bytes 48 to 79):"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2490
         Width           =   2610
      End
      Begin VB.Label lblSTA 
         AutoSize        =   -1  'True
         Caption         =   "STATION MAC (6 bytes) (Bytes 42 to 47):"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1770
         Width           =   2940
      End
      Begin VB.Label lblBSSID 
         AutoSize        =   -1  'True
         Caption         =   "BSSID (6 bytes) (Bytes 36 to 41):"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1050
         Width           =   2325
      End
      Begin VB.Label lblESSID 
         AutoSize        =   -1  'True
         Caption         =   "ESSID (36 bytes) (Bytes 0 to 35):"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   330
         Width           =   2325
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnuAddHCCAPContext 
         Caption         =   "Add to context menu of .HCCAP files"
      End
      Begin VB.Menu mnuAddCAPContext 
         Caption         =   "Add to context menu of .CAP files"
      End
      Begin VB.Menu mnuMakeHCCAPDefault 
         Caption         =   "Set as default application for .HCCAP files"
      End
      Begin VB.Menu mnuMakeCAPDefault 
         Caption         =   "Set as default application for .CAP files"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNext_Click()
 If (num_hccap_records > 1) Then
  current_index = current_index + 1
  lblCounter.caption = current_index + 1 & "/" & num_hccap_records
  If num_hccap_records - 1 > current_index Then
   btnNext.Enabled = True
  Else
   btnNext.Enabled = False
  End If
  If 0 < current_index Then
   btnPrev.Enabled = True
  Else
   btnPrev.Enabled = False
  End If
  txtESSID.Text = tmp_hccap_records(current_index).ESSID
  txtBSSID.Text = tmp_hccap_records(current_index).BSSID
  txtSTA.Text = tmp_hccap_records(current_index).STATION_MAC
  txtSNONCE.Text = tmp_hccap_records(current_index).SNONCE
  txtANONCE.Text = tmp_hccap_records(current_index).ANONCE
  txtEAPOL.Text = tmp_hccap_records(current_index).EAPOL
  txtEAPOLSIZE.Text = tmp_hccap_records(current_index).EAPOL_SIZE
  txtKEYVER.Text = tmp_hccap_records(current_index).KEY_VERSION
  txtKEYMIC.Text = tmp_hccap_records(current_index).KEY_MIC
 End If
End Sub

Private Sub btnPrev_Click()
 If (num_hccap_records > 1) Then
  current_index = current_index - 1
  lblCounter.caption = current_index + 1 & "/" & num_hccap_records
  If num_hccap_records - 1 > current_index Then
   btnNext.Enabled = True
  Else
   btnNext.Enabled = False
  End If
  If 0 < current_index Then
   btnPrev.Enabled = True
  Else
   btnPrev.Enabled = False
  End If
  txtESSID.Text = tmp_hccap_records(current_index).ESSID
  txtBSSID.Text = tmp_hccap_records(current_index).BSSID
  txtSTA.Text = tmp_hccap_records(current_index).STATION_MAC
  txtSNONCE.Text = tmp_hccap_records(current_index).SNONCE
  txtANONCE.Text = tmp_hccap_records(current_index).ANONCE
  txtEAPOL.Text = tmp_hccap_records(current_index).EAPOL
  txtEAPOLSIZE.Text = tmp_hccap_records(current_index).EAPOL_SIZE
  txtKEYVER.Text = tmp_hccap_records(current_index).KEY_VERSION
  txtKEYMIC.Text = tmp_hccap_records(current_index).KEY_MIC
 End If
End Sub

Private Sub btnWriteCAP_Click()
If txtBSSID.Text = "" Then
 MsgBox "BSSID cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtBSSID.SetFocus
 Exit Sub
ElseIf txtSTA.Text = "" Then
 MsgBox "STATION MAC cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtSTA.SetFocus
 Exit Sub
ElseIf txtSNONCE.Text = "" Then
 MsgBox "SNONCE cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtSNONCE.SetFocus
 Exit Sub
ElseIf txtANONCE.Text = "" Then
 MsgBox "ANONCE cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtANONCE.SetFocus
 Exit Sub
ElseIf txtEAPOL.Text = "" Then
 MsgBox "EAPOL cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtEAPOL.SetFocus
 Exit Sub
ElseIf txtEAPOLSIZE.Text = "" Then
 MsgBox "EAPOL Size cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtEAPOLSIZE.SetFocus
 Exit Sub
ElseIf txtKEYMIC.Text = "" Then
 MsgBox "Key MIC cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtKEYMIC.SetFocus
 Exit Sub
End If
Dim file_to_save As String
file_to_save = ShowSaveFileDialog(Me.hwnd, "Save CAP As")
Call write_this_cap(file_to_save)
End Sub

Private Sub write_this_cap(file_to_save As String)
If (file_to_save <> "") Then
 Dim msgbox_result As VbMsgBoxResult
 If is_file(file_to_save) = True Then
  msgbox_result = MsgBox("A file named """ & Left$(file_to_save, lstrlen(StrPtr(file_to_save))) & """ already exists. Replace?", vbExclamation + vbYesNo, "Warning")
  If (msgbox_result = vbNo) Then
   Call btnWriteCAP_Click
   Exit Sub
  End If
 End If
 If (num_hccap_records > 1) Then
  msgbox_result = MsgBox(current_file & " contains multiple handshakes." & vbCrLf & "Would you like to save all of them to a single CAP file?" & vbCrLf & vbCrLf & "YES: will save all handshakes to a single cap file" & vbCrLf & "NO: will save only the currently selected handshake" & vbCrLf & "CANCEL: will not save anything", vbQuestion + vbYesNoCancel, "Create multi cap?")
  If (msgbox_result = vbYes) Then
   Call WriteCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, False)
  ElseIf (msgbox_result = vbNo) Then
   Call WriteCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, True)
  End If
 Else
  Call WriteCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, True)
 End If
End If
End Sub

Private Sub read_this_cap(file_to_open As String)
If (file_to_open <> "") Then
 btnReadCAP.Enabled = False
 btnReadHCCAP.Enabled = False
 btnWriteCAP.Enabled = False
 btnWriteHCCAP.Enabled = False
 current_file = get_file_from_path(file_to_open)
 tmp_hccap_records = ReadCAP(file_to_open)
 If (num_hccap_records > 0) Then
  lblCounter.Visible = True
  txtESSID.Text = tmp_hccap_records(0).ESSID
  txtBSSID.Text = tmp_hccap_records(0).BSSID
  txtSTA.Text = tmp_hccap_records(0).STATION_MAC
  txtSNONCE.Text = tmp_hccap_records(0).SNONCE
  txtANONCE.Text = tmp_hccap_records(0).ANONCE
  txtEAPOL.Text = tmp_hccap_records(0).EAPOL
  txtEAPOLSIZE.Text = tmp_hccap_records(0).EAPOL_SIZE
  txtKEYVER.Text = tmp_hccap_records(0).KEY_VERSION
  txtKEYMIC.Text = tmp_hccap_records(0).KEY_MIC
 End If
 current_index = 0
 lblCounter.caption = current_index + 1 & "/" & num_hccap_records
 If ((current_index = 0) And (num_hccap_records = 1)) Or (num_hccap_records = 0) Then
  lblCounter.Visible = False
 End If
 If num_hccap_records > 1 Then
  btnNext.Enabled = True
  btnPrev.Enabled = False
 Else
  btnNext.Enabled = False
  btnPrev.Enabled = False
 End If
 btnReadCAP.Enabled = True
 btnReadHCCAP.Enabled = True
 btnWriteCAP.Enabled = True
 btnWriteHCCAP.Enabled = True
 btnWriteHCCAP.SetFocus
End If
End Sub

Private Sub btnReadCAP_Click()
 Dim file_to_open As String
 file_to_open = ShowOpenFileDialog(Me.hwnd, "Choose CAP File")
 Call read_this_cap(file_to_open)
End Sub

Private Sub read_this_hccap(file_to_open As String)
If (file_to_open <> "") Then
 current_file = get_file_from_path(file_to_open)
 tmp_hccap_records = ReadHCCAP(file_to_open)
 If (num_hccap_records > 0) Then
  lblCounter.Visible = True
  txtESSID.Text = tmp_hccap_records(0).ESSID
  txtBSSID.Text = tmp_hccap_records(0).BSSID
  txtSTA.Text = tmp_hccap_records(0).STATION_MAC
  txtSNONCE.Text = tmp_hccap_records(0).SNONCE
  txtANONCE.Text = tmp_hccap_records(0).ANONCE
  txtEAPOL.Text = tmp_hccap_records(0).EAPOL
  txtEAPOLSIZE.Text = tmp_hccap_records(0).EAPOL_SIZE
  txtKEYVER.Text = tmp_hccap_records(0).KEY_VERSION
  txtKEYMIC.Text = tmp_hccap_records(0).KEY_MIC
 End If
 current_index = 0
 lblCounter.caption = current_index + 1 & "/" & num_hccap_records
 If ((current_index = 0) And (num_hccap_records = 1)) Or (num_hccap_records = 0) Then
  lblCounter.Visible = False
 End If
 If num_hccap_records > 1 Then
  btnNext.Enabled = True
  btnPrev.Enabled = False
 Else
  btnNext.Enabled = False
  btnPrev.Enabled = False
 End If
 btnWriteCAP.SetFocus
End If
End Sub

Private Sub btnReadHCCAP_Click()
 Dim file_to_open As String
 file_to_open = ShowOpenFileDialog(Me.hwnd, "Choose HCCAP File")
 Call read_this_hccap(file_to_open)
End Sub

Private Sub btnWriteHCCAP_Click()
If txtBSSID.Text = "" Then
 MsgBox "BSSID cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtBSSID.SetFocus
 Exit Sub
ElseIf txtSTA.Text = "" Then
 MsgBox "STATION MAC cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtSTA.SetFocus
 Exit Sub
ElseIf txtSNONCE.Text = "" Then
 MsgBox "SNONCE cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtSNONCE.SetFocus
 Exit Sub
ElseIf txtANONCE.Text = "" Then
 MsgBox "ANONCE cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtANONCE.SetFocus
 Exit Sub
ElseIf txtEAPOL.Text = "" Then
 MsgBox "EAPOL cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtEAPOL.SetFocus
 Exit Sub
ElseIf txtEAPOLSIZE.Text = "" Then
 MsgBox "EAPOL Size cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtEAPOLSIZE.SetFocus
 Exit Sub
ElseIf txtKEYMIC.Text = "" Then
 MsgBox "Key MIC cannot be blank.", vbCritical + vbOKOnly, "Invalid Input"
 txtKEYMIC.SetFocus
 Exit Sub
End If
Dim file_to_save As String
file_to_save = ShowSaveFileDialog(Me.hwnd, "Save HCCAP As")
Call write_this_hccap(file_to_save)
End Sub

Private Sub write_this_hccap(file_to_save As String)
If (file_to_save <> "") Then
 Dim msgbox_result As VbMsgBoxResult
 If is_file(file_to_save) = True Then
  msgbox_result = MsgBox("A file named """ & Left$(file_to_save, lstrlen(StrPtr(file_to_save))) & """ already exists. Replace?", vbExclamation + vbYesNo, "Warning")
  If (msgbox_result = vbNo) Then
   Call btnWriteHCCAP_Click
   Exit Sub
  End If
 End If
 If (num_hccap_records > 1) Then
  msgbox_result = MsgBox(current_file & " contains multiple handshakes." & vbCrLf & "Would you like to save all of them to a single HCCAP file?" & vbCrLf & vbCrLf & "YES: will save all handshakes to a single multi hccap file" & vbCrLf & "NO: will save only the currently selected handshake" & vbCrLf & "CANCEL: will not save anything", vbQuestion + vbYesNoCancel, "Create multi hccap?")
  If (msgbox_result = vbYes) Then
   Call WriteHCCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, False)
  ElseIf (msgbox_result = vbNo) Then
   Call WriteHCCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, True)
  End If
 Else
  Call WriteHCCAP(file_to_save, txtESSID.Text, txtBSSID.Text, txtSTA.Text, txtSNONCE.Text, txtANONCE.Text, txtEAPOL.Text, txtEAPOLSIZE.Text, txtKEYVER.Text, txtKEYMIC.Text, True)
 End If
End If
End Sub

Private Sub Form_Load()
 On Error Resume Next
 'txtESSID.Text = "hashcat.net"
 'txtBSSID.Text = "B0:48:7A:D6:76:E2"
 'txtSTA.Text = "00:25:CF:2D:B4:89"
 'txtSNONCE.Text = "70 00 3E 0A D1 1B C0 A9 E4 86 79 45 9E BC BF FD" & vbNewLine & "7E E7 56 97 62 8C 37 13 65 D7 A0 5E 1B 35 D7 D8"
 'txtANONCE.Text = "2F 0F 76 4C 66 32 D5 57 9C 57 C3 A9 FE 06 7A 84" & vbNewLine & "5E 22 D6 43 59 41 C1 84 38 45 DB 34 A2 F8 0D DE"
 'txtEAPOL.Text = "01 03 00 75 02 01 0A 00 00 00 00 00 00 00 00 00" & vbNewLine & "01 70 00 3E 0A D1 1B C0 A9 E4 86 79 45 9E BC BF" & vbNewLine & "FD 7E E7 56 97 62 8C 37 13 65 D7 A0 5E 1B 35 D7" & vbNewLine & "D8 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 16 30 14 01 00 00 0F AC 04 01 00 00 0F AC" & vbNewLine & "04 01 00 00 0F AC 02 00 00"
 'txtEAPOLSIZE.Text = "121"
 'txtKEYVER.Text = "2"
 'txtKEYMIC.Text = "D9 F3 B5 B6 F7 44 C6 62 51 84 58 AC 6C C7 9F 11"
 
 RemoveMenu GetSystemMenu(Me.hwnd, 0), 2, &H400& 'prevent resizing
 
 If GetSetting("CapConverter", "Preferences", "AlwaysOnTop", "") = "1" Then
  mnuAlwaysOnTop.Checked = True
  SetWindowPos Me.hwnd, -1&, 0&, 0&, 0&, 0&, 2& Or 1&
 End If
 
 mnuAddHCCAPContext.Checked = IIf(RegKeyExists(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter"), True, False)
 mnuAddCAPContext.Checked = IIf(RegKeyExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter") Or RegKeyExists(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter"), True, False)
 mnuMakeHCCAPDefault.Checked = IIf(RegKeyExists(HKEY_CLASSES_ROOT, ".hccap\shell\open"), True, False)
 mnuMakeCAPDefault.Checked = IIf(RegKeyExists(HKEY_CLASSES_ROOT, ".pcap\shell\open") Or RegValueExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "bak"), True, False)
  
 'open file passed on command line
 If (Command$ <> "") Then
  Dim cmdline_file As String
  cmdline_file = Command$
  If Mid$(cmdline_file, 1, 1) = """" Then
   cmdline_file = Mid$(cmdline_file, 2) 'remove starting quote
  End If
  If Mid$(cmdline_file, Len(cmdline_file), 1) = """" Then
   cmdline_file = Mid$(cmdline_file, 1, Len(cmdline_file) - 1) 'remove ending quote
  End If
  If (is_file(cmdline_file)) Then 'file exists
   If Right$(cmdline_file, 6) = ".hccap" Then
    Call read_this_hccap(cmdline_file)
   ElseIf Right$(cmdline_file, 4) = ".cap" Or Right$(cmdline_file, 5) = ".pcap" Or Right$(cmdline_file, 4) = ".dmp" Then
    Call read_this_cap(cmdline_file)
   End If
  End If
 End If
 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 PopupMenu mnuPopUp, vbPopupMenuRightButton
End If
End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Form_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) Then
 Dim fso As Scripting.FileSystemObject
 Set fso = New Scripting.FileSystemObject
 Dim sFilename As String
 Do While Data.Files.Count > 0
  If Right$(Data.Files.Item(Data.Files.Count), 6) = ".hccap" Then
   Call read_this_hccap(Data.Files.Item(Data.Files.Count))
   Call write_this_cap(fso.GetParentFolderName(Data.Files.Item(Data.Files.Count)) & "\" & fso.GetBaseName(Data.Files.Item(Data.Files.Count)) & ".cap")
  ElseIf Right$(Data.Files.Item(Data.Files.Count), 4) = ".cap" Or Right$(Data.Files.Item(Data.Files.Count), 5) = ".pcap" Or Right$(Data.Files.Item(Data.Files.Count), 4) = ".dmp" Then
   Call read_this_cap(Data.Files.Item(Data.Files.Count))
   Call write_this_hccap(fso.GetParentFolderName(Data.Files.Item(Data.Files.Count)) & "\" & fso.GetBaseName(Data.Files.Item(Data.Files.Count)) & ".hccap")
  ElseIf is_folder(Data.Files.Item(Data.Files.Count)) = True Then 'folder
   sFilename = Dir(Data.Files.Item(Data.Files.Count) & "\")
   Do While sFilename > ""
    If Right$(sFilename, 6) = ".hccap" Then
     Call read_this_hccap(Data.Files.Item(Data.Files.Count) & "\" & sFilename)
     If is_file(Data.Files.Item(Data.Files.Count) & "\" & fso.GetBaseName(sFilename) & ".cap") = False Then
      Call write_this_cap(Data.Files.Item(Data.Files.Count) & "\" & fso.GetBaseName(sFilename) & ".cap")
     End If
    ElseIf Right$(sFilename, 4) = ".cap" Or Right$(sFilename, 5) = ".pcap" Or Right$(sFilename, 4) = ".dmp" Then
     Call read_this_cap(Data.Files.Item(Data.Files.Count) & "\" & sFilename)
     Call write_this_hccap(Data.Files.Item(Data.Files.Count) & "\" & fso.GetBaseName(sFilename) & ".hccap")
    End If
    sFilename = Dir()
   Loop
  End If
  Call Data.Files.Remove(Data.Files.Count)
 Loop
 Set fso = Nothing
End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 PopupMenu mnuPopUp, vbPopupMenuRightButton
End If
End Sub

'Add to context menu of .HCCAP files
Private Sub mnuAddHCCAPContext_Click()
 mnuAddHCCAPContext.Checked = Not mnuAddHCCAPContext.Checked
 'remove context menu
 If (mnuAddHCCAPContext.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter") = True) Then
  Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter\command")
  Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter")
 'add context menu
 ElseIf (mnuAddHCCAPContext.Checked = True) Then
  Dim path As String
  path = App.path
  If Right$(path, 1) <> "\" Then path = path & "\"
  path = path & App.EXEName & ".exe"
  Call WriteRegistry(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter", "", ValString, "&Open With Cap Converter")
  Call WriteRegistry(HKEY_CLASSES_ROOT, ".hccap\shell\Open With Cap Converter\command", "", ValString, """" & path & """" & " %1")
 End If
End Sub

'Add to context menu of .CAP files
Private Sub mnuAddCAPContext_Click()
 mnuAddCAPContext.Checked = Not mnuAddCAPContext.Checked
 Dim path As String
 path = App.path
 If Right$(path, 1) <> "\" Then path = path & "\"
 path = path & App.EXEName & ".exe"
 'wireshark is installed
 If RegKeyExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell") = True Then
  'remove context menu
  If (mnuAddCAPContext.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter") = True) Then
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter")
  'add context menu
  ElseIf (mnuAddCAPContext.Checked = True) Then
   Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter", "", ValString, "&Open With Cap Converter")
   Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter\command", "", ValString, """" & path & """" & " %1")
  End If
 'wireshark is not installed
 Else
  'remove context menu
  If (mnuAddCAPContext.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter") = True) Then
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\Open With Cap Converter")
  'add context menu
  ElseIf (mnuAddCAPContext.Checked = True) Then
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter", "", ValString, "&Open With Cap Converter")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".cap\shell\Open With Cap Converter\command", "", ValString, """" & path & """" & " %1")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter", "", ValString, "&Open With Cap Converter")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".pcap\shell\Open With Cap Converter\command", "", ValString, """" & path & """" & " %1")
  End If
 End If
End Sub

Private Sub mnuAlwaysOnTop_Click()
If (mnuAlwaysOnTop.Checked = False) Then 'make the window topmost
 mnuAlwaysOnTop.Checked = True
 SetWindowPos Me.hwnd, -1&, 0&, 0&, 0&, 0&, 2& Or 1&
 SaveSetting "CapConverter", "Preferences", "AlwaysOnTop", "1"
ElseIf (mnuAlwaysOnTop.Checked = True) Then
 mnuAlwaysOnTop.Checked = False
 SetWindowPos Me.hwnd, -2&, 0&, 0&, 0&, 0&, 2& Or 1&
 DeleteSetting "CapConverter"
 'DeleteSetting "CapConverter", "Preferences", "AlwaysOnTop"
End If
End Sub

'Set as default application for .HCCAP files
Private Sub mnuMakeHCCAPDefault_Click()
 mnuMakeHCCAPDefault.Checked = Not mnuMakeHCCAPDefault.Checked
 'remove association
 If (mnuMakeHCCAPDefault.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, ".hccap\shell\open") = True) Then
  Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".hccap\shell\open\command")
  Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".hccap\shell\open")
 'add association
 ElseIf (mnuMakeHCCAPDefault.Checked = True) Then
  Dim path As String
  path = App.path
  If Right$(path, 1) <> "\" Then path = path & "\"
  path = path & App.EXEName & ".exe"
  Call WriteRegistry(HKEY_CLASSES_ROOT, ".hccap\shell\open", "", ValString, "")
  Call WriteRegistry(HKEY_CLASSES_ROOT, ".hccap\shell\open\command", "", ValString, """" & path & """" & " %1")
 End If
End Sub

'Set as default application for .CAP files
Private Sub mnuMakeCAPDefault_Click()
 mnuMakeCAPDefault.Checked = Not mnuMakeCAPDefault.Checked
 Dim prev_value As String
 Dim path As String
 path = App.path
 If Right$(path, 1) <> "\" Then path = path & "\"
 path = path & App.EXEName & ".exe"
 
 'wireshark is installed
 If RegKeyExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell") = True Then
  'remove association
  If (mnuMakeCAPDefault.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open") = True) Then
   prev_value = ReadRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "bak")
   If (Len(prev_value) > 3) Then
    Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "", ValString, prev_value)
    Call DeleteValue(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "bak")
   End If
  'add association
  ElseIf (mnuMakeCAPDefault.Checked = True) Then
   'back up previous value
   prev_value = ReadRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", vbNullString)
   If (Len(prev_value) > 3) Then
    Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "bak", ValString, prev_value)
   End If
   Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open", "", ValString, "")
   Call WriteRegistry(HKEY_CLASSES_ROOT, "wireshark-capture-file\Shell\open\command", "", ValString, """" & path & """" & " %1")
  End If
 'wireshark is not installed
 Else
  'remove context menu
  If (mnuMakeCAPDefault.Checked = False) And (RegKeyExists(HKEY_CLASSES_ROOT, ".cap\shell\open") = True) Then
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\open\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".cap\shell\open")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\open\command")
   Call RegDeleteSubKey(HKEY_CLASSES_ROOT, ".pcap\shell\open")
  'add context menu
  ElseIf (mnuMakeCAPDefault.Checked = True) Then
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".cap\shell\open", "", ValString, "")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".cap\shell\open\command", "", ValString, """" & path & """" & " %1")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".pcap\shell\open", "", ValString, "")
   Call WriteRegistry(HKEY_CLASSES_ROOT, ".pcap\shell\open\command", "", ValString, """" & path & """" & " %1")
  End If
 End If
 
End Sub

Private Sub txtEAPOL_LostFocus()
 Dim eapol_len As Long
 eapol_len = Len(hex_digits_only(txtEAPOL.Text))
 If eapol_len Mod 2 <> 0 Then
  eapol_len = eapol_len + 1
 End If
 txtEAPOLSIZE.Text = eapol_len / 2
 txtEAPOL.Text = UCase$(txtEAPOL.Text)
End Sub

Private Sub txtEAPOLSIZE_LostFocus()
 If Val(txtEAPOLSIZE.Text) > 256 Then
  txtEAPOLSIZE.Text = "256"
 End If
End Sub

Private Sub txtESSID_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtESSID.SelStart = 0
 txtESSID.SelLength = Len(txtESSID.Text)
 End If
End Sub

Private Sub txtBSSID_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtBSSID.SelStart = 0
 txtBSSID.SelLength = Len(txtBSSID.Text)
 End If
End Sub

Private Sub txtBSSID_LostFocus()
 txtBSSID.Text = UCase$(txtBSSID.Text)
End Sub

Private Sub txtSTA_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtSTA.SelStart = 0
 txtSTA.SelLength = Len(txtSTA.Text)
 End If
End Sub

Private Sub txtSTA_LostFocus()
 txtSTA.Text = UCase$(txtSTA.Text)
End Sub

Private Sub txtSNONCE_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtSNONCE.SelStart = 0
 txtSNONCE.SelLength = Len(txtSNONCE.Text)
 End If
End Sub

Private Sub txtSNONCE_LostFocus()
 txtSNONCE.Text = UCase$(txtSNONCE.Text)
End Sub

Private Sub txtANONCE_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtANONCE.SelStart = 0
 txtANONCE.SelLength = Len(txtANONCE.Text)
 End If
End Sub

Private Sub txtANONCE_LostFocus()
 txtANONCE.Text = UCase$(txtANONCE.Text)
End Sub

Private Sub txtEAPOL_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtEAPOL.SelStart = 0
 txtEAPOL.SelLength = Len(txtEAPOL.Text)
 End If
End Sub

Private Sub txtEAPOLSIZE_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtEAPOLSIZE.SelStart = 0
 txtEAPOLSIZE.SelLength = Len(txtEAPOLSIZE.Text)
 End If
End Sub

Private Sub txtEAPOLSIZE_KeyPress(KeyAscii As Integer)
 If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
  KeyAscii = 0
 End If
End Sub

Private Sub txtKEYVER_KeyPress(KeyAscii As Integer)
 If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
  KeyAscii = 0
 End If
End Sub

Private Sub txtKEYVER_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtKEYVER.SelStart = 0
 txtKEYVER.SelLength = Len(txtKEYVER.Text)
 End If
End Sub

Private Sub txtKEYMIC_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = 65 Then
 txtKEYMIC.SelStart = 0
 txtKEYMIC.SelLength = Len(txtKEYMIC.Text)
 End If
End Sub

Private Sub txtKEYMIC_LostFocus()
 txtKEYMIC.Text = UCase$(txtKEYMIC.Text)
End Sub
