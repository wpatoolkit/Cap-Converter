Attribute VB_Name = "Write_CAP"
Option Explicit

Public Sub WriteCAP(filepath As String, ESSID As String, BSSID As String, STA As String, SNONCE As String, ANONCE As String, EAPOL As String, EAPOL_SIZE As String, KEY_VERSION As String, KEYMIC As String, Optional write_single As Boolean = False)
On Error GoTo write_cap_error

Dim obytes() As Byte
Dim i As Long, j As Long, offset As Long, current_byte As Long, max_record As Long
offset = 0

Dim ssid_length As Long
Dim eapol_length As Long
Dim actual_eapol_len As Long
Dim mac1 As String
Dim mac2 As String
Dim ANONCE_clean As String
Dim EAPOL_clean As String
Dim KEYMIC_clean As String

ssid_length = Len(ESSID)
actual_eapol_len = Len(hex_digits_only(EAPOL))
If actual_eapol_len Mod 2 <> 0 Then
 actual_eapol_len = actual_eapol_len + 1
End If
eapol_length = actual_eapol_len / 2
'eapol_length = CInt(EAPOL_SIZE)
mac1 = hex_digits_only(BSSID)
mac2 = hex_digits_only(STA)
ANONCE_clean = hex_digits_only(ANONCE)
EAPOL_clean = hex_digits_only(EAPOL)
KEYMIC_clean = Left$(hex_digits_only(KEYMIC) & "00000000000000000000000000000000", 32)

If (num_hccap_records > 0) And (write_single = False) Then
 max_record = num_hccap_records - 1
 Dim byte_tally As Long
 byte_tally = 24 'Global Header = 24 bytes
 For j = 0 To max_record
  ssid_length = Len(tmp_hccap_records(j).ESSID)
  actual_eapol_len = Len(hex_digits_only(tmp_hccap_records(j).EAPOL))
  If actual_eapol_len Mod 2 <> 0 Then
   actual_eapol_len = actual_eapol_len + 1
  End If
  eapol_length = actual_eapol_len / 2
  'eapol_length = CInt(tmp_hccap_records(j).EAPOL_SIZE)
  byte_tally = byte_tally + 16 'PACKET #1 HEADER Beacon Frame = 16 bytes
  byte_tally = byte_tally + 130 + ssid_length 'PACKET #1 DATA   Beacon Frame = 130 + ssid_length bytes
  byte_tally = byte_tally + 16 'PACKET #2 HEADER Message 1 of 4 = 16 bytes
  byte_tally = byte_tally + 133 'PACKET #2 DATA   Message 1 of 4 = 133 bytes
  byte_tally = byte_tally + 16 'PACKET #3 HEADER Message 2 of 4 = 16 bytes
  byte_tally = byte_tally + 34 + eapol_length 'PACKET #3 DATA   Message 2 of 4 = 34 + eapol_length bytes
 Next
 ReDim obytes(0 To byte_tally - 1) As Byte
Else
 max_record = 0
 ReDim obytes(0 To 369 + ssid_length + eapol_length - 1) As Byte
End If

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

For j = 0 To max_record
 If (num_hccap_records > 0) And (write_single = False) Then
  ESSID = tmp_hccap_records(j).ESSID
  BSSID = tmp_hccap_records(j).BSSID
  STA = tmp_hccap_records(j).STATION_MAC
  SNONCE = tmp_hccap_records(j).SNONCE
  ANONCE = tmp_hccap_records(j).ANONCE
  EAPOL = tmp_hccap_records(j).EAPOL
  EAPOL_SIZE = tmp_hccap_records(j).EAPOL_SIZE
  KEY_VERSION = tmp_hccap_records(j).KEY_VERSION
  KEYMIC = tmp_hccap_records(j).KEY_MIC
  ssid_length = Len(ESSID)
  actual_eapol_len = Len(hex_digits_only(EAPOL))
  If actual_eapol_len Mod 2 <> 0 Then
   actual_eapol_len = actual_eapol_len + 1
  End If
  eapol_length = actual_eapol_len / 2
  'eapol_length = CInt(EAPOL_SIZE)
  mac1 = hex_digits_only(BSSID)
  mac2 = hex_digits_only(STA)
  ANONCE_clean = hex_digits_only(ANONCE)
  EAPOL_clean = hex_digits_only(EAPOL)
  KEYMIC_clean = Left$(hex_digits_only(KEYMIC) & "00000000000000000000000000000000", 32)
 End If
 
'------------------------------------------------------------------------------------------------------
'PACKET #1 HEADER - Beacon Frame (16 bytes)
'------------------------------------------------------------------------------------------------------
'bytes 24 to 27 (ts_sec) (timestamp seconds) (packet header) (4 bytes) (ex. C2 F1 68 55)
obytes(24 + offset) = 0
obytes(25 + offset) = 0
obytes(26 + offset) = 0
obytes(27 + offset) = 0

'bytes 28 to 31 (ts_usec) (timestamp microseconds) (packet header) (4 bytes) (ex. 28 48 07 00)
obytes(28 + offset) = 0
obytes(29 + offset) = 0
obytes(30 + offset) = 0
obytes(31 + offset) = 0

'bytes 32 to 35 (incl_len) (saved packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(32 + offset) = 130 + ssid_length
obytes(33 + offset) = 0
obytes(34 + offset) = 0
obytes(35 + offset) = 0

'bytes 36 to 39 (orig_len) (actual packet length in bytes) (packet header) (4 bytes) (ex. F8 00 00 00)
obytes(36 + offset) = 130 + ssid_length
obytes(37 + offset) = 0
obytes(38 + offset) = 0
obytes(39 + offset) = 0

'------------------------------------------------------------------------------------------------------
'PACKET #1 DATA - Beacon Frame (130 + ssid_length bytes) (130 + 11 = 141)
'------------------------------------------------------------------------------------------------------
obytes(40 + offset) = 128 'Beacon frame
obytes(41 + offset) = 0   'Flags: 0x00
obytes(42 + offset) = 0   'Duration: 0 microseconds
obytes(43 + offset) = 0   'Duration: 0 microseconds
obytes(44 + offset) = 255 'Receiver Address (byte 1)
obytes(45 + offset) = 255 'Receiver Address (byte 2)
obytes(46 + offset) = 255 'Receiver Address (byte 3)
obytes(47 + offset) = 255 'Receiver Address (byte 4)
obytes(48 + offset) = 255 'Receiver Address (byte 5)
obytes(49 + offset) = 255 'Receiver Address (byte 6)

'Trasmitter Address (BSSID) (bytes 1 to 6)
For i = 50 To 55
 If Len(mac1) >= ((i - 50) + (i - 50) + 2) Then
  obytes(i + offset) = hex2dec(Mid$(mac1, (i - 50) + (i - 50) + 1, 2))
 End If
Next

'BSSID (bytes 1 to 6)
For i = 56 To 61
 If Len(mac1) >= ((i - 56) + (i - 56) + 2) Then
  obytes(i + offset) = hex2dec(Mid$(mac1, (i - 56) + (i - 56) + 1, 2))
 End If
Next

obytes(62 + offset) = 32  'Fragment Number: 0 Sequence Number: 1282
obytes(63 + offset) = 80  'Fragment Number: 0 Sequence Number: 1282
obytes(64 + offset) = 0   'Timestamp 0x00
obytes(65 + offset) = 0   'Timestamp 0x00
obytes(66 + offset) = 148 'Timestamp 0x94
obytes(67 + offset) = 17  'Timestamp 0x11
obytes(68 + offset) = 125 'Timestamp 0x7D
obytes(69 + offset) = 0   'Timestamp 0x00
obytes(70 + offset) = 0   'Timestamp 0x00
obytes(71 + offset) = 0   'Timestamp 0x00
obytes(72 + offset) = 100 'Beacon Interval 0x64
obytes(73 + offset) = 0   'Beacon Interval 0x00
obytes(74 + offset) = 17  'Capabilities Information 0x11
obytes(75 + offset) = 4   'Capabilities Information 0x04
obytes(76 + offset) = 0   'Tag Number: SSID Parameter Set (0)

'ESSID
obytes(77 + offset) = ssid_length 'Tag Length: 11 (SSID LENGTH)
current_byte = 77 + offset
If (ESSID <> "") Then
 For i = 0 To 35
  If Len(ESSID) >= (i + 1) Then
   obytes(current_byte + 1 + i) = Asc(Mid$(ESSID, (i + 1), 1))
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
If (Trim$(KEY_VERSION) = "1") Then 'WPA
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
 If Len(ANONCE_clean) >= ((i - 0) + (i - 0) + 2) Then
  obytes(current_byte + 1 + i) = hex2dec(Mid$(ANONCE_clean, (i - 0) + (i - 0) + 1, 2))
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
  If Len(KEYMIC_clean) >= ((i - 81) + (i - 81) + 2) Then
   obytes(current_byte + 1 + i) = hex2dec(Mid$(KEYMIC_clean, (i - 81) + (i - 81) + 1, 2))
  End If
 Else 'EAPOL
  If Len(EAPOL_clean) >= ((i - 0) + (i - 0) + 2) Then
   obytes(current_byte + 1 + i) = hex2dec(Mid$(EAPOL_clean, (i - 0) + (i - 0) + 1, 2))
  End If
 End If
Next

'PACKET #1 HEADER - Beacon Frame (16 bytes)
'PACKET #1 DATA - Beacon Frame (130 + ssid_length bytes)
'PACKET #2 HEADER - Message 1 of 4 (16 bytes)
'PACKET #2 DATA - Message 1 of 4 (133 bytes)
'PACKET #3 HEADER - Message 2 of 4 (16 bytes)
'PACKET #3 DATA - Message 2 of 4 (34 + eapol_length bytes)
offset = offset + 16 + 130 + ssid_length + 16 + 133 + 16 + 34 + eapol_length

Next

'write byte array to file
If is_file(filepath) Then
 Kill filepath
End If
Dim iFile As Integer
iFile = FreeFile
Open filepath For Binary Access Write As #iFile
Put #iFile, 1, obytes
Close #iFile

Exit Sub
write_cap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub
