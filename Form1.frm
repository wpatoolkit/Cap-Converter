VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Cap Converter"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   4
      Top             =   210
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private last_path As String

Private Sub btnWriteCAP_Click()
On Error GoTo write_cap_error

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

CommonDialog.fileName = ""
CommonDialog.Filter = "CAP Files (*.cap;*.pcap;*.dmp)|*.cap;*.pcap;*.dmp|All files (*.*)|*.*"
CommonDialog.DefaultExt = "cap"
CommonDialog.DialogTitle = "Save CAP As"
CommonDialog.InitDir = IIf((last_path <> ""), last_path, App.Path)
CommonDialog.ShowSave
If (CommonDialog.CancelError = False) And (CommonDialog.fileName <> "") Then
last_path = get_path_from_file(CommonDialog.fileName)
 
Dim ssid_length As Integer
ssid_length = Len(txtESSID.Text)

Dim eapol_length As Integer
Dim actual_eapol_len As Long
actual_eapol_len = Len(hex_digits_only(txtEAPOL.Text))
If actual_eapol_len Mod 2 <> 0 Then
 actual_eapol_len = actual_eapol_len + 1
End If
eapol_length = actual_eapol_len / 2
'eapol_length = CInt(txtEAPOLSIZE.Text)

Dim mac1 As String
mac1 = hex_digits_only(txtBSSID.Text)

Dim mac2 As String
mac2 = hex_digits_only(txtSTA.Text)

Dim anonce As String
anonce = hex_digits_only(txtANONCE.Text)

Dim eapol As String
eapol = hex_digits_only(txtEAPOL.Text)

Dim keymic As String
keymic = Left$(hex_digits_only(txtKEYMIC.Text) & "00000000000000000000000000000000", 32)

'Global Header = 24 bytes
'PACKET #1 HEADER Beacon Frame = 16 bytes
'PACKET #1 DATA   Beacon Frame = 130 + ssid_length bytes (130 + 11 = 141)
'PACKET #2 HEADER Message 1 of 4 = 16 bytes
'PACKET #2 DATA   Message 1 of 4 = 133 bytes
'PACKET #3 HEADER Message 2 of 4 = 16 bytes
'PACKET #3 DATA   Message 2 of 4 = 34 + eapol_length bytes (34 + 121 = 155)

Dim obytes() As Byte
ReDim obytes(0 To 369 + ssid_length + eapol_length - 1) As Byte
Dim current_byte As Integer
Dim i As Integer

'------------------------------------------------------------------------------------------------------
'GLOBAL HEADER (24 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 0 to 3 (magic_number) (global header) (4 bytes) (ex. D4 C3 B2 A1)
'D4 C3 B2 A1 = WinDump (winpcap) capture file (Windows) (little endian)
'4D 3C B2 A1 = WinDump (winpcap) capture file (Windows) (nanosecond-resolution) (little endian)
'A1 B2 C3 D4 = tcpdump (libpcap) capture file (Linux/Unix) (big endian)
'A1 B2 3C 4D = tcpdump (libpcap) capture file (Linux/Unix) (nanosecond-resolution) (big endian)
'34 CD B2 A1 = Extended tcpdump (libpcap) capture file (Linux/Unix) (big endian)
'A1 B2 CD 34 = Extended tcpdump (libpcap) capture file (Linux/Unix) (big endian)
obytes(0) = 212 'D4
obytes(1) = 195 'C3
obytes(2) = 178 'B2
obytes(3) = 161 'A1

'bytes 4 to 5 (version_major) (global header) (2 bytes) (ex. 02 00)
obytes(4) = 2
obytes(5) = 0

'bytes 6 to 7 (version_minor) (global header) (2 bytes) (ex. 04 00)
obytes(6) = 4
obytes(7) = 0

'bytes 8 to 11 (thiszone) (GMT to local correction) (global header) (4 bytes) (ex. 00 00 00 00)
obytes(8) = 0
obytes(9) = 0
obytes(10) = 0
obytes(11) = 0

'bytes 12 to 15 (sigfigs) (accuracy of timestamps) (global header) (4 bytes) (ex. 00 00 00 00)
obytes(12) = 0
obytes(13) = 0
obytes(14) = 0
obytes(15) = 0

'bytes 16 to 19 (snaplen) (max length of captured packets) (global header) (4 bytes) (ex. FF FF 00 00)
obytes(16) = 255 'FF
obytes(17) = 255 'FF
obytes(18) = 0
obytes(19) = 0

'bytes 20 to 23 (network) (Link-Layer Header Type) (global header) (4 bytes) (ex. 69 00 00 00)
'69 (hex) = 105 (dec) = IEEE 802.11 wireless LAN
'77 (hex) = 119 (dec) = Prism monitor mode information followed by an 802.11 header
'7F (hex) = 127 (dec) = Radiotap link-layer information followed by an 802.11 header
'A3 (hex) = 163 (dec) = AVS monitor mode information followed by an 802.11 header
obytes(20) = 105 'IEEE 802.11 wireless LAN
obytes(21) = 0
obytes(22) = 0
obytes(23) = 0
 
'------------------------------------------------------------------------------------------------------
'PACKET #1 HEADER - Beacon Frame (16 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 24 to 27 (ts_sec) (timestamp seconds) (packet header) (4 bytes) (ex. C2 F1 68 55)
obytes(24) = 0
obytes(25) = 0
obytes(26) = 0
obytes(27) = 0

'bytes 28 to 31 (ts_usec) (timestamp microseconds) (packet header) (4 bytes) (ex. 28 48 07 00)
obytes(28) = 0
obytes(29) = 0
obytes(30) = 0
obytes(31) = 0

'bytes 32 to 35 (incl_len) (saved packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(32) = 130 + ssid_length
obytes(33) = 0
obytes(34) = 0
obytes(35) = 0

'bytes 36 to 39 (orig_len) (actual packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(36) = 130 + ssid_length
obytes(37) = 0
obytes(38) = 0
obytes(39) = 0

'------------------------------------------------------------------------------------------------------
'PACKET #1 DATA - Beacon Frame (130 + ssid_length bytes) (130 + 11 = 141)
'------------------------------------------------------------------------------------------------------
obytes(40) = 128 'Beacon frame
obytes(41) = 0   'Flags: 0x00
obytes(42) = 0   'Duration: 0 microseconds
obytes(43) = 0   'Duration: 0 microseconds
obytes(44) = 255 'Receiver Address (byte 1)
obytes(45) = 255 'Receiver Address (byte 2)
obytes(46) = 255 'Receiver Address (byte 3)
obytes(47) = 255 'Receiver Address (byte 4)
obytes(48) = 255 'Receiver Address (byte 5)
obytes(49) = 255 'Receiver Address (byte 6)

'Trasmitter Address (BSSID) (bytes 1 to 6)
For i = 50 To 55
 If Len(mac1) >= ((i - 50) + (i - 50) + 2) Then
  obytes(i) = hex2dec(Mid$(mac1, (i - 50) + (i - 50) + 1, 2))
 End If
Next

'BSSID (bytes 1 to 6)
For i = 56 To 61
 If Len(mac1) >= ((i - 56) + (i - 56) + 2) Then
  obytes(i) = hex2dec(Mid$(mac1, (i - 56) + (i - 56) + 1, 2))
 End If
Next

obytes(62) = 32  'Fragment Number: 0 Sequence Number: 1282
obytes(63) = 80  'Fragment Number: 0 Sequence Number: 1282
obytes(64) = 0   'Timestamp 0x00
obytes(65) = 0   'Timestamp 0x00
obytes(66) = 148 'Timestamp 0x94
obytes(67) = 17  'Timestamp 0x11
obytes(68) = 125 'Timestamp 0x7D
obytes(69) = 0   'Timestamp 0x00
obytes(70) = 0   'Timestamp 0x00
obytes(71) = 0   'Timestamp 0x00
obytes(72) = 100 'Beacon Interval 0x64
obytes(73) = 0   'Beacon Interval 0x00
obytes(74) = 17  'Capabilities Information 0x11
obytes(75) = 4   'Capabilities Information 0x04
obytes(76) = 0   'Tag Number: SSID Parameter Set (0)

'ESSID
obytes(77) = ssid_length 'Tag Length: 11 (SSID LENGTH)
current_byte = 77
If (txtESSID.Text <> "") Then
 For i = 0 To 35
  If Len(txtESSID.Text) >= (i + 1) Then
   obytes(current_byte + 1 + i) = Asc(Mid$(txtESSID.Text, (i + 1), 1))
  End If
 Next
End If
current_byte = current_byte + ssid_length

obytes(current_byte + 1) = 1    'Supported Rates (1)
obytes(current_byte + 2) = 8    'Tag Length: 8
obytes(current_byte + 3) = 130  'Supported Rates: 1(B) (0x82)
obytes(current_byte + 4) = 132  'Supported Rates: 2(B) (0x84)
obytes(current_byte + 5) = 139  'Supported Rates: 5.5(B) (0x8B)
obytes(current_byte + 6) = 150  'Supported Rates: 11(B) (0x96)
obytes(current_byte + 7) = 36   'Supported Rates: 18 (0x24)
obytes(current_byte + 8) = 48   'Supported Rates: 24 (0x30)
obytes(current_byte + 9) = 72   'Supported Rates: 36 (0x48)
obytes(current_byte + 10) = 108 'Supported Rates: 54 (0x6C)
obytes(current_byte + 11) = 3   'Tag Number: DS Parameter Set (3)
obytes(current_byte + 12) = 1   'Tag length: 1
obytes(current_byte + 13) = 1   'Current Channel: 1
obytes(current_byte + 14) = 5   'Tag Number: Traffic Indication Map (TIM) (5)
obytes(current_byte + 15) = 4   'Tag length: 4
obytes(current_byte + 16) = 0   'DTIM Count: 0
obytes(current_byte + 17) = 1   'DTIM Period: 1
obytes(current_byte + 18) = 0   'Bitmap Control: 0x00
obytes(current_byte + 19) = 0   'Partial Virtual Bitmap: 00
obytes(current_byte + 20) = 42  'ERP Information (42)
obytes(current_byte + 21) = 1   'Tag length: 1
obytes(current_byte + 22) = 0   'ERP Information: 0x00
obytes(current_byte + 23) = 47  'Tag Number: ERP Information (47)
obytes(current_byte + 24) = 1   'Tag Length: 1
obytes(current_byte + 25) = 0   'ERP Information: 0x00
obytes(current_byte + 26) = 50  'Tag Number: Extended Supported Rates (50)
obytes(current_byte + 27) = 4   'Tag Length: 4
obytes(current_byte + 28) = 12  'Extended Supported Rates: 6 (0x0C)
obytes(current_byte + 29) = 18  'Extended Supported Rates: 9 (0x12)
obytes(current_byte + 30) = 24  'Extended Supported Rates: 12 (0x18)
obytes(current_byte + 31) = 96  'Extended Supported Rates: 48 (0x60)
obytes(current_byte + 32) = 221 'Tag Number: Vendor Specific (221)
obytes(current_byte + 33) = 9   'Tag Length: 9
obytes(current_byte + 34) = 0   'OUI: 00-10-18 (Broadcom)
obytes(current_byte + 35) = 16  'OUI: 00-10-18 (Broadcom)
obytes(current_byte + 36) = 24  'OUI: 00-10-18 (Broadcom)
obytes(current_byte + 37) = 2   'Vendor Specific Data: 0200f0000000
obytes(current_byte + 38) = 0   'Vendor Specific Data: 0200f0000000
obytes(current_byte + 39) = 240 'Vendor Specific Data: 0200f0000000
obytes(current_byte + 40) = 0   'Vendor Specific Data: 0200f0000000
obytes(current_byte + 41) = 0   'Vendor Specific Data: 0200f0000000
obytes(current_byte + 42) = 0   'Vendor Specific Data: 0200f0000000
obytes(current_byte + 43) = 221 'Tag Number: Vendor Specific (221)
obytes(current_byte + 44) = 22  'Tag Length: 22
obytes(current_byte + 45) = 1   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 46) = 0   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 47) = 0   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 48) = 15  'Vendor Specific Data 0x0F
obytes(current_byte + 49) = 172 'Vendor Specific Data 0xAC
obytes(current_byte + 50) = 4   'Vendor Specific Data 0x04
obytes(current_byte + 51) = 1   'Vendor Specific Data 0x01
obytes(current_byte + 52) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 53) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 54) = 15  'Vendor Specific Data 0x0F
obytes(current_byte + 55) = 172 'Vendor Specific Data 0xAC
obytes(current_byte + 56) = 4   'Vendor Specific Data 0x04
obytes(current_byte + 57) = 1   'Vendor Specific Data 0x01
obytes(current_byte + 58) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 59) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 60) = 15  'Vendor Specific Data 0x0F
obytes(current_byte + 61) = 172 'Vendor Specific Data 0xAC
obytes(current_byte + 62) = 2   'Vendor Specific Data 0x02
obytes(current_byte + 63) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 64) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 65) = 12  'Vendor Specific Data 0x0C
obytes(current_byte + 66) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 67) = 221 'Tag Number: Vendor Specific (221)
obytes(current_byte + 68) = 24  'Tag Length: 24
obytes(current_byte + 69) = 1   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 70) = 0   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 71) = 0   'OUI: 01-00-00 (00:00:00)
obytes(current_byte + 72) = 2   'Vendor Specific Data 0x02
obytes(current_byte + 73) = 1   'Vendor Specific Data 0x01
obytes(current_byte + 74) = 1   'Vendor Specific Data 0x01
obytes(current_byte + 75) = 128 'Vendor Specific Data 0x80
obytes(current_byte + 76) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 77) = 3   'Vendor Specific Data 0x03
obytes(current_byte + 78) = 164 'Vendor Specific Data 0xA4
obytes(current_byte + 79) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 80) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 81) = 39  'Vendor Specific Data 0x27
obytes(current_byte + 82) = 164 'Vendor Specific Data 0xA4
obytes(current_byte + 83) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 84) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 85) = 66  'Vendor Specific Data 0x42
obytes(current_byte + 86) = 67  'Vendor Specific Data 0x43
obytes(current_byte + 87) = 94  'Vendor Specific Data 0x5E
obytes(current_byte + 88) = 0   'Vendor Specific Data 0x00
obytes(current_byte + 89) = 98  'Vendor Specific Data 0x62
obytes(current_byte + 90) = 50  'Vendor Specific Data 0x32
obytes(current_byte + 91) = 47  'Vendor Specific Data 0x2F
obytes(current_byte + 92) = 0   'Vendor Specific Data 0x00

'------------------------------------------------------------------------------------------------------
'PACKET #2 HEADER - Message 1 of 4 (16 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 181 to 184 (ts_sec) (timestamp seconds) (packet header) (4 bytes) (ex. C2 F1 68 55)
obytes(current_byte + 93) = 0
obytes(current_byte + 94) = 0
obytes(current_byte + 95) = 0
obytes(current_byte + 96) = 0

'bytes 185 to 188 (ts_usec) (timestamp microseconds) (packet header) (4 bytes) (ex. 28 48 07 00)
obytes(current_byte + 97) = 0
obytes(current_byte + 98) = 0
obytes(current_byte + 99) = 0
obytes(current_byte + 100) = 0

'bytes 189 to 192 (incl_len) (saved packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(current_byte + 101) = 133
obytes(current_byte + 102) = 0
obytes(current_byte + 103) = 0
obytes(current_byte + 104) = 0

'bytes 193 to 196 (orig_len) (actual packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(current_byte + 105) = 133
obytes(current_byte + 106) = 0
obytes(current_byte + 107) = 0
obytes(current_byte + 108) = 0

'------------------------------------------------------------------------------------------------------
'PACKET #2 DATA - Message 1 of 4 (133 bytes)
'------------------------------------------------------------------------------------------------------
obytes(current_byte + 109) = 136  'QoS Data
obytes(current_byte + 110) = 2    'Flags: 0x02
obytes(current_byte + 111) = 58   'Duration: 314 microseconds
obytes(current_byte + 112) = 1    'Duration: 314 microseconds

current_byte = current_byte + 112

'STATION ADDRESS (bytes 1 to 6)
For i = 0 To 5
 If Len(mac2) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac2, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 6

'Transmitter Address (BSSID) (bytes 1 to 6)
For i = 0 To 5
 If Len(mac1) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac1, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 6

'Source Address (BSSID) (bytes 1 to 6)
For i = 0 To 5
 If Len(mac1) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac1, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 7

obytes(current_byte) = 0        'Fragment Number: 0
obytes(current_byte + 1) = 0    'Sequence Number: 0
obytes(current_byte + 2) = 0    'QoS Control: 0x0000
obytes(current_byte + 3) = 0    'QoS Control: 0x0000
obytes(current_byte + 4) = 170  'DSAP: SNAP (0xAA)
obytes(current_byte + 5) = 170  'SSAP: SNAP (0xAA)
obytes(current_byte + 6) = 3    'Control field: U, func=UI (0x03)
obytes(current_byte + 7) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 8) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 9) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 10) = 136 'Type: 802.1X Authentication (0x888e)
obytes(current_byte + 11) = 142 'Type: 802.1X Authentication (0x888e)
obytes(current_byte + 12) = 2   'Version: 802.1X-2004 (2)
obytes(current_byte + 13) = 3   'Type: Key (3)
obytes(current_byte + 14) = 0   'Length: 95
obytes(current_byte + 15) = 95  'Length: 95
obytes(current_byte + 16) = 2   'Key Descriptor Type: EAPOL RSN Key (2)

'KEYVER (Value of 1 means WPA, else WPA2)
If (Trim$(txtKEYVER.Text) = "1") Then 'WPA
 obytes(current_byte + 17) = 0   'Key Information (byte 1) 0x0089 = RC4 Cipher, HMAC-MD5 MIC (1)
 obytes(current_byte + 18) = 137 'Key Information (byte 2)
Else 'WPA2
 obytes(current_byte + 17) = 0   'Key Information (byte 1) 0x008a or 0x010a = AES Cipher, HMAC-SHA1 MIC (2)
 obytes(current_byte + 18) = 138 'Key Information (byte 2)
End If

obytes(current_byte + 19) = 0  'Length: 32
obytes(current_byte + 20) = 32 'Length: 32
obytes(current_byte + 21) = 0  'Replay Counter: 1
obytes(current_byte + 22) = 0  'Replay Counter: 1
obytes(current_byte + 23) = 0  'Replay Counter: 1
obytes(current_byte + 24) = 0  'Replay Counter: 1
obytes(current_byte + 25) = 0  'Replay Counter: 1
obytes(current_byte + 26) = 0  'Replay Counter: 1
obytes(current_byte + 27) = 0  'Replay Counter: 1
obytes(current_byte + 28) = 1  'Replay Counter: 1

current_byte = current_byte + 28

'ANONCE (32 bytes) - random salt used for handshake by both parties
For i = 0 To 31
 If Len(anonce) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(anonce, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 33

obytes(current_byte) = 0      'Key IV: 00000000000000000000000000000000
obytes(current_byte + 1) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 2) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 3) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 4) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 5) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 6) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 7) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 8) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 9) = 0  'Key IV: 00000000000000000000000000000000
obytes(current_byte + 10) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 11) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 12) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 13) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 14) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 15) = 0 'Key IV: 00000000000000000000000000000000
obytes(current_byte + 16) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 17) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 18) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 19) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 20) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 21) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 22) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 23) = 0 'WPA Key RSC: 0000000000000000
obytes(current_byte + 24) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 25) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 26) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 27) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 28) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 29) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 30) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 31) = 0 'WPA Key ID: 0000000000000000
obytes(current_byte + 32) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 33) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 34) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 35) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 36) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 37) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 38) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 39) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 40) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 41) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 42) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 43) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 44) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 45) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 46) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 47) = 0 'WPA Key MIC: 00000000000000000000000000000000
obytes(current_byte + 48) = 0 'WPA Key Data Length: 0
obytes(current_byte + 49) = 0 'WPA Key Data Length: 0

'------------------------------------------------------------------------------------------------------
'PACKET #3 HEADER - Message 2 of 4 (16 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 330 to 333 (ts_sec) (timestamp seconds) (packet header) (4 bytes) (ex. C2 F1 68 55)
obytes(current_byte + 50) = 0
obytes(current_byte + 51) = 0
obytes(current_byte + 52) = 0
obytes(current_byte + 53) = 0

'bytes 334 to 337 (ts_usec) (timestamp microseconds) (packet header) (4 bytes) (ex. 28 48 07 00)
obytes(current_byte + 54) = 0
obytes(current_byte + 55) = 0
obytes(current_byte + 56) = 0
obytes(current_byte + 57) = 0

'bytes 338 to 341 (incl_len) (saved packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(current_byte + 58) = 34 + eapol_length
obytes(current_byte + 59) = 0
obytes(current_byte + 60) = 0
obytes(current_byte + 61) = 0

'bytes 342 to 345 (orig_len) (actual packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(current_byte + 62) = 34 + eapol_length
obytes(current_byte + 63) = 0
obytes(current_byte + 64) = 0
obytes(current_byte + 65) = 0

'------------------------------------------------------------------------------------------------------
'PACKET #3 DATA - Message 2 of 4 (34 + eapol_length bytes) (34 + 121 = 155)
'------------------------------------------------------------------------------------------------------
obytes(current_byte + 66) = 136  'QoS Data
obytes(current_byte + 67) = 1    'Flags: 0x01
obytes(current_byte + 68) = 58   'Duration: 314 microseconds
obytes(current_byte + 69) = 1    'Duration: 314 microseconds

current_byte = current_byte + 69

'Receiver Address (BSSID) (bytes 1 to 6)
For i = 0 To 5
 If Len(mac1) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac1, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 6

'Transmitter Address (STATION ADDRESS) (bytes 1 to 6)
For i = 0 To 5
 If Len(mac2) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac2, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 6

'Destination Address (BSSID) (bytes 1 to 6)
For i = 0 To 5
 If Len(mac1) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(mac1, (i - 0) + (i - 0) + 1, 2))
 End If
Next
current_byte = current_byte + 7

obytes(current_byte) = 0        'Fragment Number: 0
obytes(current_byte + 1) = 0    'Sequence Number: 0
obytes(current_byte + 2) = 0    'QoS Control: 0x0000
obytes(current_byte + 3) = 0    'QoS Control: 0x0000
obytes(current_byte + 4) = 170  'DSAP: SNAP (0xaa)
obytes(current_byte + 5) = 170  'SSAP: SNAP (0xaa)
obytes(current_byte + 6) = 3    'Control field: U, func=UI (0x03)
obytes(current_byte + 7) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 8) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 9) = 0    'Organization Code: Encapsulated Ethernet (0x000000)
obytes(current_byte + 10) = 136 '802.1X Authentication (0x888e)
obytes(current_byte + 11) = 142 '802.1X Authentication (0x888e)

current_byte = current_byte + 11

'EAPOL (bytes 1 to 256)
For i = 0 To 255
 If (i >= 81) And (i < 97) Then 'fill in keymic
  If Len(keymic) >= ((i - 81) + (i - 81) + 2) Then
   obytes(current_byte + 1 + i) = hex2dec(Mid$(keymic, (i - 81) + (i - 81) + 1, 2))
  End If
 Else 'EAPOL
  If Len(eapol) >= ((i - 0) + (i - 0) + 2) Then
   obytes(current_byte + 1 + i) = hex2dec(Mid$(eapol, (i - 0) + (i - 0) + 1, 2))
  End If
 End If
Next

'write byte array to file
Dim iFile As Integer
iFile = FreeFile
Open CommonDialog.fileName For Binary Access Write As #iFile
Put #iFile, 1, obytes
Close #iFile
  
End If

Exit Sub
write_cap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnReadCAP_Click()
On Error GoTo read_cap_error

CommonDialog.fileName = ""
CommonDialog.Filter = "CAP Files (*.cap;*.pcap;*.dmp)|*.cap;*.pcap;*.dmp|All files (*.*)|*.*"
CommonDialog.FilterIndex = 1
CommonDialog.DefaultExt = "cap"
CommonDialog.DialogTitle = "Choose CAP File"
CommonDialog.InitDir = IIf((last_path <> ""), last_path, App.Path)
CommonDialog.ShowOpen
If (CommonDialog.CancelError = False) And (CommonDialog.fileName <> "") Then
last_path = get_path_from_file(CommonDialog.fileName)
  
btnReadCAP.Enabled = False
Dim ibytes() As Byte
Dim use_little_endian As Boolean
Dim total_bytes As Long
Dim max_packet_length As Long
Dim packet_length As Long
Dim packet_count As Long
Dim current_byte As Long
Dim ssid_is_blank As Boolean
Dim eapol_has_been_set As Boolean
Dim anonce_has_been_set As Boolean
Dim non_qos_offset As Integer
Dim eapol_length_to_use As Long
Dim i As Long

'read in file to byte array
Dim iFile As Integer
iFile = FreeFile
Open CommonDialog.fileName For Binary Access Read As #iFile
total_bytes = LOF(iFile)
If (total_bytes < 1) Then
 Close #iFile
 MsgBox "File is empty!", vbCritical + vbOKOnly, "Invalid File"
 btnReadCAP.Enabled = True
 Exit Sub
End If
ReDim ibytes(0 To total_bytes - 1 + 100) As Byte
Get #iFile, 1, ibytes
Close #iFile

If (total_bytes <= 40) Then
 MsgBox "Invalid CAP file!", vbCritical + vbOKOnly, "Invalid File"
 btnReadCAP.Enabled = True
 Exit Sub
End If

'If (total_bytes > 10240) Then
' MsgBox "This program can only accept .CAP files that are <10kb in size." & vbNewLine & vbNewLine & "Please use pyrit to clean/strip your cap:" & vbNewLine & "pyrit -r IN.CAP -o OUT.CAP strip" & vbNewLine & vbNewLine & "Or use Wireshark by following this tutorial:" & vbNewLine & "http://hackforums.net/showthread.php?tid=2974396", vbCritical + vbOKOnly, "File Too Large"
' btnReadCAP.Enabled = True
' Exit Sub
'End If

'If (total_bytes > UBound(ibytes) + 1) Then
' total_bytes = UBound(ibytes) + 1
'End If

'------------------------------------------------------------------------------------------------------
'GLOBAL HEADER (24 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 0 to 3 (magic_number) (global header) (4 bytes) (ex. D4 C3 B2 A1)
 'D4 C3 B2 A1 = WinDump (winpcap) capture file (Windows) (little endian)
 '4D 3C B2 A1 = WinDump (winpcap) capture file (Windows) (nanosecond-resolution) (little endian)
 'A1 B2 C3 D4 = tcpdump (libpcap) capture file (Linux/Unix) (big endian)
 'A1 B2 3C 4D = tcpdump (libpcap) capture file (Linux/Unix) (nanosecond-resolution) (big endian)
 '34 CD B2 A1 = Extended tcpdump (libpcap) capture file (Linux/Unix) (big endian)
 'A1 B2 CD 34 = Extended tcpdump (libpcap) capture file (Linux/Unix) (big endian)
 If ibytes(0) = 212 And ibytes(1) = 195 And ibytes(2) = 178 And ibytes(3) = 161 Then 'D4 C3 B2 A1
 ElseIf ibytes(0) = 77 And ibytes(1) = 60 And ibytes(2) = 178 And ibytes(3) = 161 Then '4D 3C B2 A1
 ElseIf ibytes(0) = 161 And ibytes(1) = 178 And ibytes(2) = 195 And ibytes(3) = 212 Then 'A1 B2 C3 D4
 ElseIf ibytes(0) = 161 And ibytes(1) = 178 And ibytes(2) = 60 And ibytes(3) = 77 Then 'A1 B2 3C 4D
 ElseIf ibytes(0) = 52 And ibytes(1) = 205 And ibytes(2) = 178 And ibytes(3) = 161 Then '34 CD B2 A1
 ElseIf ibytes(0) = 161 And ibytes(1) = 178 And ibytes(2) = 205 And ibytes(3) = 52 Then 'A1 B2 CD 34
 Else
  MsgBox "Invalid file signature!", vbCritical + vbOKOnly, "Invalid File"
  btnReadCAP.Enabled = True
  Exit Sub
 End If
  
 If (ibytes(0) = 212 And ibytes(1) = 195 And ibytes(2) = 178 And ibytes(3) = 161) Or (ibytes(0) = 77 And ibytes(1) = 60 And ibytes(2) = 178 And ibytes(3) = 161) Then
  use_little_endian = True
 Else
  use_little_endian = False
 End If

'bytes 4 to 5 (version_major) (global header) (2 bytes) (ex. 02 00)

'bytes 6 to 7 (version_minor) (global header) (2 bytes) (ex. 04 00)

'bytes 8 to 11 (thiszone) (GMT to local correction) (global header) (4 bytes) (ex. 00 00 00 00)

'bytes 12 to 15 (sigfigs) (accuracy of timestamps) (global header) (4 bytes) (ex. 00 00 00 00)

'bytes 16 to 19 (snaplen) (max length of captured packets) (global header) (4 bytes) (ex. FF FF 00 00)

If (use_little_endian = True) Then
 max_packet_length = hex2dec_lng(dec2hex(ibytes(16)) & dec2hex(ibytes(17)))
Else
 max_packet_length = hex2dec_lng(dec2hex(ibytes(18)) & dec2hex(ibytes(19)))
End If

'bytes 20 to 23 (network) (Link-Layer Header Type) (global header) (4 bytes) (ex. 69 00 00 00)
 '69 (hex) = 105 (dec) = IEEE 802.11 wireless LAN
 '77 (hex) = 119 (dec) = Prism monitor mode information followed by an 802.11 header
 '7F (hex) = 127 (dec) = Radiotap link-layer information followed by an 802.11 header
 'A3 (hex) = 163 (dec) = AVS monitor mode information followed by an 802.11 header
 If ((use_little_endian = True) And (ibytes(20) = 105)) Or ((use_little_endian = False) And (ibytes(23) = 105)) Then
 ElseIf ((use_little_endian = True) And (ibytes(20) = 119)) Or ((use_little_endian = False) And (ibytes(23) = 119)) Then
 ElseIf ((use_little_endian = True) And (ibytes(20) = 127)) Or ((use_little_endian = False) And (ibytes(23) = 127)) Then
 ElseIf ((use_little_endian = True) And (ibytes(20) = 163)) Or ((use_little_endian = False) And (ibytes(23) = 163)) Then
 Else
  MsgBox "Invalid Link Layer!" & vbNewLine & vbNewLine & "This program only accepts IEEE 802.11 wireless LAN .cap files.", vbCritical + vbOKOnly, "Invalid File"
  btnReadCAP.Enabled = True
  Exit Sub
 End If
 
 current_byte = 24
 Do While (current_byte < total_bytes)
  '------------------------------------------------------------------------------------------------------
  'PACKET HEADER (16 bytes)
  '------------------------------------------------------------------------------------------------------
  'current_byte+0 to current_byte+3 (ts_sec) (timestamp seconds) (packet header) (4 bytes) (ex. C2 F1 68 55)
  'current_byte+4 to current_byte+7 (ts_usec) (timestamp microseconds) (packet header) (4 bytes) (ex. 28 48 07 00)
  'current_byte+8 to current_byte+11 (incl_len) (saved packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
  'current_byte+12 to current_byte+15 (orig_len) (actual packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
  packet_count = packet_count + 1
  If (use_little_endian = True) Then
   packet_length = bytes2num(ibytes(current_byte + 8), ibytes(current_byte + 9))
  Else
   packet_length = bytes2num(ibytes(current_byte + 10), ibytes(current_byte + 11))
  End If
  
  current_byte = current_byte + 16
  '------------------------------------------------------------------------------------------------------
  'PACKET DATA (variable length)
  '------------------------------------------------------------------------------------------------------
  If (packet_length > 0) Then
   
   'beacon frame
   If ibytes(current_byte) = 128 Then 'byte 1 = 80 (hex) 128 (dec)
    'grab SSID
    If ibytes(current_byte + 37) > 0 And ibytes(current_byte + 37) <= 36 Then 'valid SSID length (1-36)
     ssid_is_blank = True
     For i = 1 To ibytes(current_byte + 37)
      If ibytes(current_byte + 37 + i) <> 0 Then
       ssid_is_blank = False
      End If
     Next
     If (ssid_is_blank = False) Then
      txtESSID.Text = ""
      For i = 1 To ibytes(current_byte + 37)
       txtESSID.Text = txtESSID.Text & Chr(ibytes(current_byte + 37 + i))
      Next
     End If
    End If
    'grab BSSID
    txtBSSID.Text = ""
    For i = 1 To 6
     txtBSSID.Text = txtBSSID.Text & dec2hex(ibytes(current_byte + 15 + i))
     If (i < 6) Then
      txtBSSID.Text = txtBSSID.Text & ":"
     End If
    Next
   End If
   
   'probe response
   If ibytes(current_byte) = 80 Then 'byte 1 = 50 (hex) 80 (dec)
    'grab SSID
    If ibytes(current_byte + 37) > 0 And ibytes(current_byte + 37) <= 36 Then 'valid SSID length (1-36)
     txtESSID.Text = ""
     For i = 1 To ibytes(current_byte + 37)
      txtESSID.Text = txtESSID.Text & Chr(ibytes(current_byte + 37 + i))
     Next
    End If
    'grab BSSID
    txtBSSID.Text = ""
    For i = 1 To 6
     txtBSSID.Text = txtBSSID.Text & dec2hex(ibytes(current_byte + 15 + i))
     If (i < 6) Then
      txtBSSID.Text = txtBSSID.Text & ":"
     End If
    Next
   End If
   
   'Message 1 of 4
   'Frame Control Field: 0x8802 QoS Data (AP -> STA) (bytes 1 + 2)
   'Frame Type: 802.1X Authentication (0x888e) (bytes 33 + 34)
   'or
   'Frame Control Field: 0x0802 Data (AP -> STA) (bytes 1 + 2)
   'Frame Type: 802.1X Authentication (0x888e) (bytes 31 + 32)
   If (ibytes(current_byte) = 136 And ((ibytes(current_byte + 1) = 2) Or (ibytes(current_byte + 1) = 10)) And ibytes(current_byte + 32) = 136 And ibytes(current_byte + 33) = 142) Or (ibytes(current_byte) = 8 And ((ibytes(current_byte + 1) = 2) Or (ibytes(current_byte + 1) = 10)) And ibytes(current_byte + 30) = 136 And ibytes(current_byte + 31) = 142) Then
    If ibytes(current_byte) = 8 Then
     non_qos_offset = 2
    Else
     non_qos_offset = 0
    End If
    'Station Address
    txtSTA.Text = ""
    For i = 1 To 6
     'Receiver Address (bytes 5 to 10)
     txtSTA.Text = txtSTA.Text & dec2hex(ibytes(current_byte + 3 + i))
     If (i < 6) Then
      txtSTA.Text = txtSTA.Text & ":"
     End If
    Next
    'BSSID (backup) (bytes 11 to 16)
    txtBSSID.Text = ""
    For i = 1 To 6
     txtBSSID.Text = txtBSSID.Text & dec2hex(ibytes(current_byte + 9 + i))
     If (i < 6) Then
      txtBSSID.Text = txtBSSID.Text & ":"
     End If
    Next
    'ANONCE
    If (ibytes(current_byte + 51) = 0) And (ibytes(current_byte + 52) = 0) And (ibytes(current_byte + 53) = 0) And (ibytes(current_byte + 54) = 0) And (ibytes(current_byte + 55) = 0) And (ibytes(current_byte + 56) = 0) And (ibytes(current_byte + 57) = 0) And (ibytes(current_byte + 58) = 0) And (ibytes(current_byte + 59) = 0) And (ibytes(current_byte + 60) = 0) And (ibytes(current_byte + 61) = 0) And (ibytes(current_byte + 62) = 0) And (ibytes(current_byte + 63) = 0) And (ibytes(current_byte + 64) = 0) And (ibytes(current_byte + 65) = 0) And (ibytes(current_byte + 66) = 0) And (ibytes(current_byte + 67) = 0) And (ibytes(current_byte + 68) = 0) And (ibytes(current_byte + 69) = 0) And (ibytes(current_byte + 70) = 0) And (ibytes(current_byte + 71) = 0) And (ibytes(current_byte + 72) = 0) And (ibytes(current_byte + 73) = 0) And (ibytes(current_byte + 74) = 0) And (ibytes(current_byte + 75) = 0) And (ibytes(current_byte + 76) = 0) And (ibytes(current_byte + 77) = 0) _
    And (ibytes(current_byte + 78) = 0) And (ibytes(current_byte + 79) = 0) And (ibytes(current_byte + 80) = 0) And (ibytes(current_byte + 81) = 0) And (ibytes(current_byte + 82) = 0) Then
    Else
     If (anonce_has_been_set = False) Then
      txtANONCE.Text = ""
      For i = 1 To 32
       'ANONCE (bytes 52 to 83)
       txtANONCE.Text = txtANONCE.Text & dec2hex(ibytes(current_byte + 50 - non_qos_offset + i))
       If (i = 16) Then
        txtANONCE.Text = txtANONCE.Text & vbNewLine
       ElseIf (i < 32) Then
        txtANONCE.Text = txtANONCE.Text & " "
       End If
      Next
      anonce_has_been_set = True
     End If
    End If
   
   'Message 2 of 4
   'Frame Control Field: 0x8801 QoS Data (STA -> AP) (bytes 1 + 2)
   'Frame Type: 802.1X Authentication (0x888e) (bytes 33 + 34)
   'or
   'Frame Control Field: 0x0801 Data (AP -> STA) (bytes 1 + 2)
   'Frame Type: 802.1X Authentication (0x888e) (bytes 31 + 32)
   ElseIf (ibytes(current_byte) = 136 And ((ibytes(current_byte + 1) = 1) Or (ibytes(current_byte + 1) = 9)) And ibytes(current_byte + 32) = 136 And ibytes(current_byte + 33) = 142) Or (ibytes(current_byte) = 8 And ((ibytes(current_byte + 1) = 1) Or (ibytes(current_byte + 1) = 9)) And ibytes(current_byte + 30) = 136 And ibytes(current_byte + 31) = 142) Then
    If ibytes(current_byte) = 8 Then
     non_qos_offset = 2
    Else
     non_qos_offset = 0
    End If
    If (ibytes(current_byte + 51 - non_qos_offset) = 0) And (ibytes(current_byte + 52 - non_qos_offset) = 0) And (ibytes(current_byte + 53 - non_qos_offset) = 0) And (ibytes(current_byte + 54 - non_qos_offset) = 0) And (ibytes(current_byte + 55 - non_qos_offset) = 0) And (ibytes(current_byte + 56 - non_qos_offset) = 0) And (ibytes(current_byte + 57 - non_qos_offset) = 0) And (ibytes(current_byte + 58 - non_qos_offset) = 0) And (ibytes(current_byte + 59 - non_qos_offset) = 0) And (ibytes(current_byte + 60 - non_qos_offset) = 0) And (ibytes(current_byte + 61 - non_qos_offset) = 0) And (ibytes(current_byte + 62 - non_qos_offset) = 0) And (ibytes(current_byte + 63 - non_qos_offset) = 0) And (ibytes(current_byte + 64 - non_qos_offset) = 0) And (ibytes(current_byte + 65 - non_qos_offset) = 0) And (ibytes(current_byte + 66 - non_qos_offset) = 0) _
    And (ibytes(current_byte + 67 - non_qos_offset) = 0) And (ibytes(current_byte + 68 - non_qos_offset) = 0) And (ibytes(current_byte + 69 - non_qos_offset) = 0) And (ibytes(current_byte + 70 - non_qos_offset) = 0) And (ibytes(current_byte + 71 - non_qos_offset) = 0) And (ibytes(current_byte + 72 - non_qos_offset) = 0) And (ibytes(current_byte + 73 - non_qos_offset) = 0) And (ibytes(current_byte + 74 - non_qos_offset) = 0) And (ibytes(current_byte + 75 - non_qos_offset) = 0) And (ibytes(current_byte + 76 - non_qos_offset) = 0) And (ibytes(current_byte + 77 - non_qos_offset) = 0) And (ibytes(current_byte + 78 - non_qos_offset) = 0) And (ibytes(current_byte + 79 - non_qos_offset) = 0) And (ibytes(current_byte + 80 - non_qos_offset) = 0) And (ibytes(current_byte + 81 - non_qos_offset) = 0) And (ibytes(current_byte + 82 - non_qos_offset) = 0) Then
    Else
     'EAPOL
     If packet_length > (34 - non_qos_offset) Then
      If (eapol_has_been_set = False) Then
      
      'SNONCE
      txtSNONCE.Text = ""
      For i = 1 To 32
      'SNONCE (bytes 52 to 83)
       txtSNONCE.Text = txtSNONCE.Text & dec2hex(ibytes(current_byte + 50 - non_qos_offset + i))
       If (i = 16) Then
        txtSNONCE.Text = txtSNONCE.Text & vbNewLine
       ElseIf (i < 32) Then
        txtSNONCE.Text = txtSNONCE.Text & " "
       End If
      Next
      
      txtEAPOLSIZE.Text = hex2dec_lng(dec2hex(ibytes(current_byte + 36 - non_qos_offset)) & dec2hex(ibytes(current_byte + 37 - non_qos_offset))) + 4
      'txtEAPOLSIZE.Text = packet_length - (34 - non_qos_offset)
      
      If CLng(txtEAPOLSIZE.Text) > 0 Then
       eapol_length_to_use = CLng(txtEAPOLSIZE.Text)
      Else
       eapol_length_to_use = packet_length - (34 - non_qos_offset)
      End If
      
      txtEAPOL.Text = ""
      txtKEYMIC.Text = ""
      For i = 1 To eapol_length_to_use
       'Key Version
       If (i = 7) Then
        txtKEYVER.Text = hex2dec_lng(dec2hex(ibytes(current_byte + 33 - non_qos_offset + i - 1)) & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))) And 7
       End If
       'Key MIC
       If (i > 81) And (i < 98) Then
        txtEAPOL.Text = txtEAPOL.Text & "00"
        txtKEYMIC.Text = txtKEYMIC.Text & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))
        If (i < 97) Then
         txtKEYMIC.Text = txtKEYMIC.Text & " "
        End If
       Else
        txtEAPOL.Text = txtEAPOL.Text & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))
       End If
       
       If (i Mod 16 = 0) Then
        txtEAPOL.Text = txtEAPOL.Text & vbNewLine
       ElseIf i < eapol_length_to_use Then
        txtEAPOL.Text = txtEAPOL.Text & " "
       End If
      Next
      eapol_has_been_set = True
      End If
     End If
    End If
   End If
  
   current_byte = current_byte + packet_length 'move to next packet
  End If
 Loop

btnReadCAP.Enabled = True
btnWriteHCCAP.SetFocus

End If

Exit Sub
read_cap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnReadHCCAP_Click()
 On Error GoTo read_hccap_error
 
 CommonDialog.fileName = ""
 CommonDialog.Filter = "HCCAP Files (*.hccap)|*.hccap"
 CommonDialog.DefaultExt = "hccap"
 CommonDialog.DialogTitle = "Choose HCCAP File"
 CommonDialog.InitDir = IIf((last_path <> ""), last_path, App.Path)
 CommonDialog.ShowOpen
 If (CommonDialog.CancelError = False) And (CommonDialog.fileName <> "") Then
  last_path = get_path_from_file(CommonDialog.fileName)
  
  Dim ibytes(0 To 391) As Byte
  Dim file_length As Long

  'read in file to byte array
  Dim iFile As Integer
  iFile = FreeFile
  Open CommonDialog.fileName For Binary Access Read As #iFile
  file_length = LOF(iFile)
  Get #iFile, 1, ibytes
  Close #iFile
  
  If (file_length = 392) Then

  'bytes 0 to 35 = essid (36 bytes) - the essid (name) of the access point
  txtESSID.Text = ""
  For i = 0 To 35
   If (ibytes(i) <> 0) Then
    txtESSID.Text = txtESSID.Text & Chr(ibytes(i))
   End If
  Next

  'bytes 36 to 41 = mac1 (6 bytes) - the bssid (MAC) of the access point
  txtBSSID.Text = ""
  For i = 36 To 41
   txtBSSID.Text = txtBSSID.Text & dec2hex(ibytes(i))
   If (i < 41) Then
    txtBSSID.Text = txtBSSID.Text & ":"
   End If
  Next

  'bytes 42 to 47 = mac2 (6 bytes) - the MAC address of a client connecting to the access point
  txtSTA.Text = ""
  For i = 42 To 47
   txtSTA.Text = txtSTA.Text & dec2hex(ibytes(i))
   If (i < 47) Then
    txtSTA.Text = txtSTA.Text & ":"
   End If
  Next

  'bytes 48 to 79 = snonce (32 bytes) - random salt used for handshake by both parties
  txtSNONCE.Text = ""
  For i = 48 To 79
   txtSNONCE.Text = txtSNONCE.Text & dec2hex(ibytes(i))
   If (i <> 63) And (i <> 79) Then
    txtSNONCE.Text = txtSNONCE.Text & " "
   End If
   If (i = 63) Then
    txtSNONCE.Text = txtSNONCE.Text & vbNewLine
   End If
  Next

  'bytes 80 to 111 = anonce (32 bytes) - random salt used for handshake by both parties
  txtANONCE.Text = ""
  For i = 80 To 111
   txtANONCE.Text = txtANONCE.Text & dec2hex(ibytes(i))
   If (i <> 95) And (i <> 111) Then
    txtANONCE.Text = txtANONCE.Text & " "
   End If
   If (i = 95) Then
    txtANONCE.Text = txtANONCE.Text & vbNewLine
   End If
  Next

  'bytes 112 to 367 = eapol (256 bytes) - EAPOL
  Dim eapol_size As Long
  eapol_size = bytes2num(ibytes(368), ibytes(369))
  txtEAPOL.Text = ""
  For i = 112 To 367
   If ((i - 111) <= eapol_size) Then
    txtEAPOL.Text = txtEAPOL.Text & dec2hex(ibytes(i))
    If (i <> 127) And (i <> 143) And (i <> 159) And (i <> 175) And (i <> 191) And (i <> 207) And (i <> 223) And (i <> 239) And (i <> 255) And (i <> 271) And (i <> 287) And (i <> 303) And (i <> 319) And (i <> 335) And (i <> 351) And (i <> 367) And ((i - 111) <> eapol_size) Then
     txtEAPOL.Text = txtEAPOL.Text & " "
    End If
    If (i = 127) Or (i = 143) Or (i = 159) Or (i = 175) Or (i = 191) Or (i = 207) Or (i = 223) Or (i = 239) Or (i = 255) Or (i = 271) Or (i = 287) Or (i = 303) Or (i = 319) Or (i = 335) Or (i = 351) Then
     txtEAPOL.Text = txtEAPOL.Text & vbNewLine
    End If
   End If
  Next

  'bytes 368 to 371 = eapol_size (4 bytes) - size of eapol
  txtEAPOLSIZE.Text = eapol_size

  'bytes 372 to 375 = keyver (4 bytes) - the flag used to distinguish WPA from WPA2 ciphers. Value of 1 means WPA, other - WPA2
  txtKEYVER.Text = ibytes(372)

  'bytes 376 to 391 = keymic (16 bytes) - the final hash value. MD5 for WPA and SHA-1 for WPA2 (truncated to 128 bit)
  txtKEYMIC.Text = ""
  For i = 376 To 391
   txtKEYMIC.Text = txtKEYMIC.Text & dec2hex(ibytes(i))
   If (i <> 391) Then
    txtKEYMIC.Text = txtKEYMIC.Text & " "
   End If
  Next
  
  Else
   MsgBox "Invalid HCCAP file!", vbCritical + vbOKOnly, "Invalid File"
  End If
  
  btnWriteCAP.SetFocus
  
 End If
 
Exit Sub
read_hccap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnWriteHCCAP_Click()
 On Error GoTo write_hccap_error
 
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
 
 CommonDialog.fileName = ""
 CommonDialog.Filter = "HCCAP Files (*.hccap)|*.hccap"
 CommonDialog.DefaultExt = "hccap"
 CommonDialog.DialogTitle = "Save HCCAP As"
 CommonDialog.InitDir = IIf((last_path <> ""), last_path, App.Path)
 CommonDialog.ShowSave
 If (CommonDialog.CancelError = False) And (CommonDialog.fileName <> "") Then
  last_path = get_path_from_file(CommonDialog.fileName)
  Dim obytes(0 To 391) As Byte
  Dim i As Integer

  'bytes 0 to 35 = essid (36 bytes) - the essid (name) of the access point
  If (txtESSID.Text <> "") Then
  For i = 0 To 35
   If Len(txtESSID.Text) >= (i + 1) Then
    obytes(i) = Asc(Mid$(txtESSID.Text, (i + 1), 1))
   End If
  Next
  End If

  'bytes 36 to 41 = mac1 (6 bytes) - the bssid (MAC) of the access point
  If (txtBSSID.Text <> "") Then
  Dim mac1 As String
  mac1 = hex_digits_only(txtBSSID.Text)
  For i = 36 To 41
   If Len(mac1) >= ((i - 36) + (i - 36) + 2) Then
    obytes(i) = hex2dec(Mid$(mac1, (i - 36) + (i - 36) + 1, 2))
   End If
  Next
  End If

  'bytes 42 to 47 = mac2 (6 bytes) - the MAC address of a client connecting to the access point
  If (txtSTA.Text <> "") Then
  Dim mac2 As String
  mac2 = hex_digits_only(txtSTA.Text)
  For i = 42 To 47
   If Len(mac2) >= ((i - 42) + (i - 42) + 2) Then
    obytes(i) = hex2dec(Mid$(mac2, (i - 42) + (i - 42) + 1, 2))
   End If
  Next
  End If

  'bytes 48 to 79 = snonce (32 bytes) - random salt used for handshake by both parties
  If (txtSNONCE.Text <> "") Then
  Dim snonce As String
  snonce = hex_digits_only(txtSNONCE.Text)
  For i = 48 To 79
   If Len(snonce) >= ((i - 48) + (i - 48) + 2) Then
    obytes(i) = hex2dec(Mid$(snonce, (i - 48) + (i - 48) + 1, 2))
   End If
  Next
  End If

  'bytes 80 to 111 = anonce (32 bytes) - random salt used for handshake by both parties
  If (txtANONCE.Text <> "") Then
  Dim anonce As String
  anonce = hex_digits_only(txtANONCE.Text)
  For i = 80 To 111
   If Len(anonce) >= ((i - 80) + (i - 80) + 2) Then
    obytes(i) = hex2dec(Mid$(anonce, (i - 80) + (i - 80) + 1, 2))
   End If
  Next
  End If

  'bytes 112 to 367 = eapol (256 bytes) - EAPOL
  If (txtEAPOL.Text <> "") Then
  Dim eapol As String
  eapol = hex_digits_only(txtEAPOL.Text)
  For i = 112 To 367
   If Len(eapol) >= ((i - 112) + (i - 112) + 2) Then
    obytes(i) = hex2dec(Mid$(eapol, (i - 112) + (i - 112) + 1, 2))
   End If
  Next
  End If

  'bytes 368 to 371 = eapol_size (4 bytes) - size of eapol
  If (txtEAPOLSIZE.Text <> "") Then
  obytes(368) = CInt(txtEAPOLSIZE.Text) And &HFF&
  obytes(369) = (CInt(txtEAPOLSIZE.Text) And &HFF00&) / 256
  End If

  'bytes 372 to 375 = keyver (4 bytes) - the flag used to distinguish WPA from WPA2 ciphers. Value of 1 means WPA, other - WPA2
  If (txtKEYVER.Text <> "") Then
  obytes(372) = IIf(txtKEYVER.Text = "1", 1, 2)
  End If

  'bytes 376 to 391 = keymic (16 bytes) - the final hash value. MD5 for WPA and SHA-1 for WPA2 (truncated to 128 bit)
  If (txtKEYMIC.Text <> "") Then
  Dim keymic As String
  keymic = hex_digits_only(txtKEYMIC.Text)
  For i = 376 To 391
   If Len(keymic) >= ((i - 376) + (i - 376) + 2) Then
    obytes(i) = hex2dec(Mid$(keymic, (i - 376) + (i - 376) + 1, 2))
   End If
  Next
  End If

  'write byte array to file
  Dim iFile As Integer
  iFile = FreeFile
  Open CommonDialog.fileName For Binary Access Write As #iFile
  Put #iFile, 1, obytes
  Close #iFile
  
  'MsgBox "Done!", vbInformation + vbOKOnly, "HCCAP Editor"
  
 End If
 
Exit Sub
write_hccap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
 txtESSID.Text = "hashcat.net"
 txtBSSID.Text = "B0:48:7A:D6:76:E2"
 txtSTA.Text = "00:25:CF:2D:B4:89"
 txtSNONCE.Text = "70 00 3E 0A D1 1B C0 A9 E4 86 79 45 9E BC BF FD" & vbNewLine & "7E E7 56 97 62 8C 37 13 65 D7 A0 5E 1B 35 D7 D8"
 txtANONCE.Text = "2F 0F 76 4C 66 32 D5 57 9C 57 C3 A9 FE 06 7A 84" & vbNewLine & "5E 22 D6 43 59 41 C1 84 38 45 DB 34 A2 F8 0D DE"
 txtEAPOL.Text = "01 03 00 75 02 01 0A 00 00 00 00 00 00 00 00 00" & vbNewLine & "01 70 00 3E 0A D1 1B C0 A9 E4 86 79 45 9E BC BF" & vbNewLine & "FD 7E E7 56 97 62 8C 37 13 65 D7 A0 5E 1B 35 D7" & vbNewLine & "D8 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & vbNewLine & "00 00 16 30 14 01 00 00 0F AC 04 01 00 00 0F AC" & vbNewLine & "04 01 00 00 0F AC 02 00 00"
 txtEAPOLSIZE.Text = "121"
 txtKEYVER.Text = "2"
 txtKEYMIC.Text = "D9 F3 B5 B6 F7 44 C6 62 51 84 58 AC 6C C7 9F 11"
 
 RemoveMenu GetSystemMenu(Me.hwnd, 0), 2, &H400& 'prevent resizing
 
 'Set Available Monospace Font
 Dim has_fixedsys As Boolean
 Dim has_lucida_console As Boolean
 Dim has_courier_new As Boolean
 Dim i As Integer
 For i = 0 To Screen.FontCount - 1
  If Screen.Fonts(i) = "Fixedsys" Then
   has_fixedsys = True
  ElseIf Screen.Fonts(i) = "Lucida Console" Then
   has_lucida_console = True
  ElseIf Screen.Fonts(i) = "Courier New" Then
   has_courier_new = True
  End If
 Next i
 If (has_fixedsys = True) Then
  txtESSID.Font.Name = "Fixedsys": txtESSID.Font.Size = 9
  txtBSSID.Font.Name = "Fixedsys": txtBSSID.Font.Size = 9
  txtSTA.Font.Name = "Fixedsys": txtSTA.Font.Size = 9
  txtSNONCE.Font.Name = "Fixedsys": txtSNONCE.Font.Size = 9
  txtANONCE.Font.Name = "Fixedsys": txtANONCE.Font.Size = 9
  txtEAPOL.Font.Name = "Fixedsys": txtEAPOL.Font.Size = 9
  txtEAPOLSIZE.Font.Name = "Fixedsys": txtEAPOLSIZE.Font.Size = 9
  txtKEYVER.Font.Name = "Fixedsys": txtKEYVER.Font.Size = 9
  txtKEYMIC.Font.Name = "Fixedsys": txtKEYMIC.Font.Size = 9
 ElseIf (has_lucida_console = True) Then
  txtESSID.Font.Name = "Lucida Console": txtESSID.Font.Size = 10
  txtBSSID.Font.Name = "Lucida Console": txtBSSID.Font.Size = 10
  txtSTA.Font.Name = "Lucida Console": txtSTA.Font.Size = 10
  txtSNONCE.Font.Name = "Lucida Console": txtSNONCE.Font.Size = 10
  txtANONCE.Font.Name = "Lucida Console": txtANONCE.Font.Size = 10
  txtEAPOL.Font.Name = "Lucida Console": txtEAPOL.Font.Size = 10
  txtEAPOLSIZE.Font.Name = "Lucida Console": txtEAPOLSIZE.Font.Size = 10
  txtKEYVER.Font.Name = "Lucida Console": txtKEYVER.Font.Size = 10
  txtKEYMIC.Font.Name = "Lucida Console": txtKEYMIC.Font.Size = 10
 Else
  txtESSID.Font.Name = "Courier New": txtESSID.Font.Size = 10
  txtBSSID.Font.Name = "Courier New": txtBSSID.Font.Size = 10
  txtSTA.Font.Name = "Courier New": txtSTA.Font.Size = 10
  txtSNONCE.Font.Name = "Courier New": txtSNONCE.Font.Size = 10
  txtANONCE.Font.Name = "Courier New": txtANONCE.Font.Size = 10
  txtEAPOL.Font.Name = "Courier New": txtEAPOL.Font.Size = 10
  txtEAPOLSIZE.Font.Name = "Courier New": txtEAPOLSIZE.Font.Size = 10
  txtKEYVER.Font.Name = "Courier New": txtKEYVER.Font.Size = 10
  txtKEYMIC.Font.Name = "Courier New": txtKEYMIC.Font.Size = 10
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

Private Function get_path_from_file(fileName As String) As String
 Dim pos As Integer
 pos = InStrRev(fileName, "\")
 If pos > 0 Then
  get_path_from_file = Left$(fileName, pos)
 Else
  get_path_from_file = ""
 End If
End Function

Private Function dec2hex(d As Byte) As String
 dec2hex = Right$("0" & Hex$(d), 2)
End Function

Private Function hex2dec(h As String) As Byte
 hex2dec = IIf(Len(h) > 0, CByte("&H" & h), 0)
End Function

Private Function hex2dec_lng(h As String) As Long
 hex2dec_lng = IIf(Len(h) > 0, CLng("&H" & h), 0)
End Function

Private Function bytes2num(LoByte As Byte, HiByte As Byte) As Long
 If HiByte And &H80 Then
  bytes2num = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
 Else
  bytes2num = (HiByte * &H100) Or LoByte
 End If
 'bytes2num = CLng("&H" & Right$("0" & Hex$(LoByte), 2) & Right$("0" & Hex$(HiByte), 2))
End Function

Private Function hex_digits_only(input_str As String) As String
 Dim output_str As String
 Dim c As String
 If Len(input_str) > 0 Then
  Dim i As Integer
  For i = 1 To Len(input_str)
   c = UCase$(Mid$(input_str, i, 1))
   If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Or c = "A" Or c = "B" Or c = "C" Or c = "D" Or c = "E" Or c = "F" Then
    output_str = output_str & c
   End If
  Next
 End If
 hex_digits_only = output_str
End Function
