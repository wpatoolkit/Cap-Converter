Attribute VB_Name = "Read_CAP"
Option Explicit

Public Function ReadCAP(filepath As String) As hccap_record()
On Error GoTo read_cap_error

Dim tmp_hccap_records() As hccap_record
Dim unique_bssids As String
Dim unique_bssids_count As Long
Dim bssid_index As Long
num_hccap_records = 0

Dim ibytes() As Byte
Dim use_little_endian As Boolean
Dim total_bytes As Long
Dim max_packet_length As Long
Dim packet_length As Long
Dim packet_count As Long
Dim current_byte As Long
Dim tmp_bssid As String
Dim ssid_is_blank As Boolean
Dim essid_has_been_set As Boolean
Dim bssid_has_been_set As Boolean
Dim sta_mac_has_been_set As Boolean
Dim eapol_has_been_set As Boolean
Dim anonce_has_been_set As Boolean
Dim non_qos_offset As Integer
Dim eapol_length_to_use As Long
Dim i As Long, j As Long, k As Long

'read in file to byte array
Dim iFile As Integer
iFile = FreeFile
Open filepath For Binary Access Read As #iFile
total_bytes = LOF(iFile)
If (total_bytes < 1) Then
 Close #iFile
 MsgBox "File is empty!", vbCritical + vbOKOnly, "Invalid File"
 Exit Function
End If
ReDim ibytes(0 To total_bytes - 1 + 100) As Byte
Get #iFile, 1, ibytes
Close #iFile

If (total_bytes <= 40) Then
 MsgBox "Invalid CAP file!", vbCritical + vbOKOnly, "Invalid File"
 Exit Function
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
  Exit Function
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
  Exit Function
 End If
 
 '------------------------------------------------------------------------------------------------------
 'COUNT UNIQUE BSSIDS
 '------------------------------------------------------------------------------------------------------
 current_byte = 24
 Do While (current_byte < total_bytes)
  If (use_little_endian = True) Then
   packet_length = bytes2num(ibytes(current_byte + 8), ibytes(current_byte + 9))
  Else
   packet_length = bytes2num(ibytes(current_byte + 10), ibytes(current_byte + 11))
  End If
  current_byte = current_byte + 16 'jump to packet data
  If (packet_length > 0) Then
   'BEACON FRAME
   If ibytes(current_byte) = 128 Then 'byte 1 = 80 (hex) 128 (dec)
    tmp_bssid = dec2hex(ibytes(current_byte + 15 + 1)) & ":" & dec2hex(ibytes(current_byte + 15 + 2)) & ":" & dec2hex(ibytes(current_byte + 15 + 3)) & ":" & dec2hex(ibytes(current_byte + 15 + 4)) & ":" & dec2hex(ibytes(current_byte + 15 + 5)) & ":" & dec2hex(ibytes(current_byte + 15 + 6)) & ","
    If InStr(unique_bssids, tmp_bssid) = 0 Then
     unique_bssids = unique_bssids & tmp_bssid
     unique_bssids_count = unique_bssids_count + 1
     ReDim Preserve tmp_hccap_records(0 To unique_bssids_count - 1) As hccap_record
     tmp_hccap_records(unique_bssids_count - 1).BSSID = Left$(tmp_bssid, 17)
    End If
   End If
   'PROBE RESPONSE
   If ibytes(current_byte) = 80 Then 'byte 1 = 50 (hex) 80 (dec)
    tmp_bssid = dec2hex(ibytes(current_byte + 15 + 1)) & ":" & dec2hex(ibytes(current_byte + 15 + 2)) & ":" & dec2hex(ibytes(current_byte + 15 + 3)) & ":" & dec2hex(ibytes(current_byte + 15 + 4)) & ":" & dec2hex(ibytes(current_byte + 15 + 5)) & ":" & dec2hex(ibytes(current_byte + 15 + 6)) & ","
    If InStr(unique_bssids, tmp_bssid) = 0 Then
     unique_bssids = unique_bssids & tmp_bssid
     unique_bssids_count = unique_bssids_count + 1
     ReDim Preserve tmp_hccap_records(0 To unique_bssids_count - 1) As hccap_record
     tmp_hccap_records(unique_bssids_count - 1).BSSID = Left$(tmp_bssid, 17)
    End If
   End If
   'MESSAGE 1 of 4
   If (ibytes(current_byte) = 136 And ((ibytes(current_byte + 1) = 2) Or (ibytes(current_byte + 1) = 10)) And ibytes(current_byte + 32) = 136 And ibytes(current_byte + 33) = 142) Or (ibytes(current_byte) = 8 And ((ibytes(current_byte + 1) = 2) Or (ibytes(current_byte + 1) = 10)) And ibytes(current_byte + 30) = 136 And ibytes(current_byte + 31) = 142) Then
    tmp_bssid = dec2hex(ibytes(current_byte + 9 + 1)) & ":" & dec2hex(ibytes(current_byte + 9 + 2)) & ":" & dec2hex(ibytes(current_byte + 9 + 3)) & ":" & dec2hex(ibytes(current_byte + 9 + 4)) & ":" & dec2hex(ibytes(current_byte + 9 + 5)) & ":" & dec2hex(ibytes(current_byte + 9 + 6)) & ","
    If InStr(unique_bssids, tmp_bssid) = 0 Then
     unique_bssids = unique_bssids & tmp_bssid
     unique_bssids_count = unique_bssids_count + 1
     ReDim Preserve tmp_hccap_records(0 To unique_bssids_count - 1) As hccap_record
     tmp_hccap_records(unique_bssids_count - 1).BSSID = Left$(tmp_bssid, 17)
    End If
   'Message 2 of 4
   ElseIf (ibytes(current_byte) = 136 And ((ibytes(current_byte + 1) = 1) Or (ibytes(current_byte + 1) = 9)) And ibytes(current_byte + 32) = 136 And ibytes(current_byte + 33) = 142) Or (ibytes(current_byte) = 8 And ((ibytes(current_byte + 1) = 1) Or (ibytes(current_byte + 1) = 9)) And ibytes(current_byte + 30) = 136 And ibytes(current_byte + 31) = 142) Then
    tmp_bssid = dec2hex(ibytes(current_byte + 3 + 1)) & ":" & dec2hex(ibytes(current_byte + 3 + 2)) & ":" & dec2hex(ibytes(current_byte + 3 + 3)) & ":" & dec2hex(ibytes(current_byte + 3 + 4)) & ":" & dec2hex(ibytes(current_byte + 3 + 5)) & ":" & dec2hex(ibytes(current_byte + 3 + 6)) & ","
    If InStr(unique_bssids, tmp_bssid) = 0 Then
     unique_bssids = unique_bssids & tmp_bssid
     unique_bssids_count = unique_bssids_count + 1
     ReDim Preserve tmp_hccap_records(0 To unique_bssids_count - 1) As hccap_record
     tmp_hccap_records(unique_bssids_count - 1).BSSID = Left$(tmp_bssid, 17)
    End If
   End If
   current_byte = current_byte + packet_length 'move to next packet
  End If
 Loop
 If (unique_bssids_count = 0) Then
  MsgBox "No BSSIDs found!", vbCritical + vbOKOnly, "Invalid File"
  Exit Function
 End If
 num_hccap_records = unique_bssids_count
 '------------------------------------------------------------------------------------------------------
 
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
    'grab BSSID
     tmp_bssid = dec2hex(ibytes(current_byte + 15 + 1)) & ":" & dec2hex(ibytes(current_byte + 15 + 2)) & ":" & dec2hex(ibytes(current_byte + 15 + 3)) & ":" & dec2hex(ibytes(current_byte + 15 + 4)) & ":" & dec2hex(ibytes(current_byte + 15 + 5)) & ":" & dec2hex(ibytes(current_byte + 15 + 6))
    'find this BSSIDs index in the array
    For k = 0 To unique_bssids_count - 1
     If (tmp_hccap_records(k).BSSID = tmp_bssid) Then
      bssid_index = k
      Exit For
     End If
    Next
    'grab SSID
    If ibytes(current_byte + 37) > 0 And ibytes(current_byte + 37) <= 36 Then 'valid SSID length (1-36)
     ssid_is_blank = True
     For i = 1 To ibytes(current_byte + 37)
      If ibytes(current_byte + 37 + i) <> 0 Then
       ssid_is_blank = False
       Exit For
      End If
     Next
     If (ssid_is_blank = False) Then
      If (tmp_hccap_records(bssid_index).ESSID = "") Then
       For i = 1 To ibytes(current_byte + 37)
        tmp_hccap_records(bssid_index).ESSID = tmp_hccap_records(bssid_index).ESSID & Chr(ibytes(current_byte + 37 + i))
       Next
      End If
     End If
    End If
   End If
   
   'probe response
   If ibytes(current_byte) = 80 Then 'byte 1 = 50 (hex) 80 (dec)
    'grab BSSID
    tmp_bssid = dec2hex(ibytes(current_byte + 15 + 1)) & ":" & dec2hex(ibytes(current_byte + 15 + 2)) & ":" & dec2hex(ibytes(current_byte + 15 + 3)) & ":" & dec2hex(ibytes(current_byte + 15 + 4)) & ":" & dec2hex(ibytes(current_byte + 15 + 5)) & ":" & dec2hex(ibytes(current_byte + 15 + 6))
    'find this BSSIDs index in the array
    For k = 0 To unique_bssids_count - 1
     If (tmp_hccap_records(k).BSSID = tmp_bssid) Then
      bssid_index = k
      Exit For
     End If
    Next
    'grab SSID
    If ibytes(current_byte + 37) > 0 And ibytes(current_byte + 37) <= 36 Then 'valid SSID length (1-36)
     If (tmp_hccap_records(bssid_index).ESSID = "") Then
      For i = 1 To ibytes(current_byte + 37)
       tmp_hccap_records(bssid_index).ESSID = tmp_hccap_records(bssid_index).ESSID & Chr(ibytes(current_byte + 37 + i))
      Next
     End If
    End If
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
    'BSSID (bytes 11 to 16)
    tmp_bssid = dec2hex(ibytes(current_byte + 9 + 1)) & ":" & dec2hex(ibytes(current_byte + 9 + 2)) & ":" & dec2hex(ibytes(current_byte + 9 + 3)) & ":" & dec2hex(ibytes(current_byte + 9 + 4)) & ":" & dec2hex(ibytes(current_byte + 9 + 5)) & ":" & dec2hex(ibytes(current_byte + 9 + 6))
    'find this BSSIDs index in the array
    For k = 0 To unique_bssids_count - 1
     If (tmp_hccap_records(k).BSSID = tmp_bssid) Then
      bssid_index = k
      Exit For
     End If
    Next
    'Station Address
    'Receiver Address (bytes 5 to 10)
    If (tmp_hccap_records(bssid_index).STATION_MAC = "") Then
     tmp_hccap_records(bssid_index).STATION_MAC = dec2hex(ibytes(current_byte + 3 + 1)) & ":" & dec2hex(ibytes(current_byte + 3 + 2)) & ":" & dec2hex(ibytes(current_byte + 3 + 3)) & ":" & dec2hex(ibytes(current_byte + 3 + 4)) & ":" & dec2hex(ibytes(current_byte + 3 + 5)) & ":" & dec2hex(ibytes(current_byte + 3 + 6))
    End If
    'ANONCE
    If (ibytes(current_byte + 51) = 0) And (ibytes(current_byte + 52) = 0) And (ibytes(current_byte + 53) = 0) And (ibytes(current_byte + 54) = 0) And (ibytes(current_byte + 55) = 0) And (ibytes(current_byte + 56) = 0) And (ibytes(current_byte + 57) = 0) And (ibytes(current_byte + 58) = 0) And (ibytes(current_byte + 59) = 0) And (ibytes(current_byte + 60) = 0) And (ibytes(current_byte + 61) = 0) And (ibytes(current_byte + 62) = 0) And (ibytes(current_byte + 63) = 0) And (ibytes(current_byte + 64) = 0) And (ibytes(current_byte + 65) = 0) And (ibytes(current_byte + 66) = 0) And (ibytes(current_byte + 67) = 0) And (ibytes(current_byte + 68) = 0) And (ibytes(current_byte + 69) = 0) And (ibytes(current_byte + 70) = 0) And (ibytes(current_byte + 71) = 0) And (ibytes(current_byte + 72) = 0) And (ibytes(current_byte + 73) = 0) And (ibytes(current_byte + 74) = 0) And (ibytes(current_byte + 75) = 0) And (ibytes(current_byte + 76) = 0) And (ibytes(current_byte + 77) = 0) _
    And (ibytes(current_byte + 78) = 0) And (ibytes(current_byte + 79) = 0) And (ibytes(current_byte + 80) = 0) And (ibytes(current_byte + 81) = 0) And (ibytes(current_byte + 82) = 0) Then
    Else
     If (tmp_hccap_records(bssid_index).ANONCE = "") Then
      For i = 1 To 32
       'ANONCE (bytes 52 to 83)
       tmp_hccap_records(bssid_index).ANONCE = tmp_hccap_records(bssid_index).ANONCE & dec2hex(ibytes(current_byte + 50 - non_qos_offset + i))
       If (i = 16) Then
        tmp_hccap_records(bssid_index).ANONCE = tmp_hccap_records(bssid_index).ANONCE & vbNewLine
       ElseIf (i < 32) Then
        tmp_hccap_records(bssid_index).ANONCE = tmp_hccap_records(bssid_index).ANONCE & " "
       End If
      Next
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
    'BSSID (bytes 5 to 10)
    tmp_bssid = dec2hex(ibytes(current_byte + 3 + 1)) & ":" & dec2hex(ibytes(current_byte + 3 + 2)) & ":" & dec2hex(ibytes(current_byte + 3 + 3)) & ":" & dec2hex(ibytes(current_byte + 3 + 4)) & ":" & dec2hex(ibytes(current_byte + 3 + 5)) & ":" & dec2hex(ibytes(current_byte + 3 + 6))
    'find this BSSIDs index in the array
    For k = 0 To unique_bssids_count - 1
     If (tmp_hccap_records(k).BSSID = tmp_bssid) Then
      bssid_index = k
      Exit For
     End If
    Next
    If (ibytes(current_byte + 51 - non_qos_offset) = 0) And (ibytes(current_byte + 52 - non_qos_offset) = 0) And (ibytes(current_byte + 53 - non_qos_offset) = 0) And (ibytes(current_byte + 54 - non_qos_offset) = 0) And (ibytes(current_byte + 55 - non_qos_offset) = 0) And (ibytes(current_byte + 56 - non_qos_offset) = 0) And (ibytes(current_byte + 57 - non_qos_offset) = 0) And (ibytes(current_byte + 58 - non_qos_offset) = 0) And (ibytes(current_byte + 59 - non_qos_offset) = 0) And (ibytes(current_byte + 60 - non_qos_offset) = 0) And (ibytes(current_byte + 61 - non_qos_offset) = 0) And (ibytes(current_byte + 62 - non_qos_offset) = 0) And (ibytes(current_byte + 63 - non_qos_offset) = 0) And (ibytes(current_byte + 64 - non_qos_offset) = 0) And (ibytes(current_byte + 65 - non_qos_offset) = 0) And (ibytes(current_byte + 66 - non_qos_offset) = 0) _
    And (ibytes(current_byte + 67 - non_qos_offset) = 0) And (ibytes(current_byte + 68 - non_qos_offset) = 0) And (ibytes(current_byte + 69 - non_qos_offset) = 0) And (ibytes(current_byte + 70 - non_qos_offset) = 0) And (ibytes(current_byte + 71 - non_qos_offset) = 0) And (ibytes(current_byte + 72 - non_qos_offset) = 0) And (ibytes(current_byte + 73 - non_qos_offset) = 0) And (ibytes(current_byte + 74 - non_qos_offset) = 0) And (ibytes(current_byte + 75 - non_qos_offset) = 0) And (ibytes(current_byte + 76 - non_qos_offset) = 0) And (ibytes(current_byte + 77 - non_qos_offset) = 0) And (ibytes(current_byte + 78 - non_qos_offset) = 0) And (ibytes(current_byte + 79 - non_qos_offset) = 0) And (ibytes(current_byte + 80 - non_qos_offset) = 0) And (ibytes(current_byte + 81 - non_qos_offset) = 0) And (ibytes(current_byte + 82 - non_qos_offset) = 0) Then
    Else
     'EAPOL
     If packet_length > (34 - non_qos_offset) Then
      'SNONCE
      If (tmp_hccap_records(bssid_index).SNONCE = "") Then
       For i = 1 To 32
       'SNONCE (bytes 52 to 83)
        tmp_hccap_records(bssid_index).SNONCE = tmp_hccap_records(bssid_index).SNONCE & dec2hex(ibytes(current_byte + 50 - non_qos_offset + i))
        If (i = 16) Then
         tmp_hccap_records(bssid_index).SNONCE = tmp_hccap_records(bssid_index).SNONCE & vbNewLine
        ElseIf (i < 32) Then
         tmp_hccap_records(bssid_index).SNONCE = tmp_hccap_records(bssid_index).SNONCE & " "
        End If
       Next
       tmp_hccap_records(bssid_index).EAPOL_SIZE = hex2dec_lng(dec2hex(ibytes(current_byte + 36 - non_qos_offset)) & dec2hex(ibytes(current_byte + 37 - non_qos_offset))) + 4
       'tmp_hccap_records(bssid_index).EAPOL_SIZE = packet_length - (34 - non_qos_offset)
       If CLng(tmp_hccap_records(bssid_index).EAPOL_SIZE) > 0 Then
        eapol_length_to_use = CLng(tmp_hccap_records(bssid_index).EAPOL_SIZE)
       Else
        eapol_length_to_use = packet_length - (34 - non_qos_offset)
       End If
       tmp_hccap_records(bssid_index).EAPOL = ""
       tmp_hccap_records(bssid_index).KEY_MIC = ""
       For i = 1 To eapol_length_to_use
        'Key Version
        If (i = 7) Then
         tmp_hccap_records(bssid_index).KEY_VERSION = hex2dec_lng(dec2hex(ibytes(current_byte + 33 - non_qos_offset + i - 1)) & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))) And 7
        End If
        'Key MIC
        If (i > 81) And (i < 98) Then
         tmp_hccap_records(bssid_index).EAPOL = tmp_hccap_records(bssid_index).EAPOL & "00"
         tmp_hccap_records(bssid_index).KEY_MIC = tmp_hccap_records(bssid_index).KEY_MIC & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))
         If (i < 97) Then
          tmp_hccap_records(bssid_index).KEY_MIC = tmp_hccap_records(bssid_index).KEY_MIC & " "
         End If
        Else
         tmp_hccap_records(bssid_index).EAPOL = tmp_hccap_records(bssid_index).EAPOL & dec2hex(ibytes(current_byte + 33 - non_qos_offset + i))
        End If
        
        If (i Mod 16 = 0) Then
         tmp_hccap_records(bssid_index).EAPOL = tmp_hccap_records(bssid_index).EAPOL & vbNewLine
        ElseIf i < eapol_length_to_use Then
         tmp_hccap_records(bssid_index).EAPOL = tmp_hccap_records(bssid_index).EAPOL & " "
        End If
       Next
      End If
     End If
    End If
   End If
  
   current_byte = current_byte + packet_length 'move to next packet
  End If
 Loop
 
ReadCAP = tmp_hccap_records
Exit Function
read_cap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Function
