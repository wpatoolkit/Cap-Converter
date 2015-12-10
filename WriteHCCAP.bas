Attribute VB_Name = "Write_HCCAP"
Option Explicit

Public Sub WriteHCCAP(filepath As String, ESSID As String, BSSID As String, STA As String, SNONCE As String, ANONCE As String, EAPOL As String, EAPOL_SIZE As String, KEY_VERSION As String, KEYMIC As String, Optional write_single As Boolean = False)
On Error GoTo write_hccap_error

Dim obytes() As Byte
Dim i As Long, j As Long, offset As Long, max_record As Long
offset = 0

If (num_hccap_records > 0) And (write_single = False) Then
 ReDim obytes(0 To (num_hccap_records * 392) - 1) As Byte
 max_record = num_hccap_records - 1
Else
 ReDim obytes(0 To 391) As Byte
 max_record = 0
End If

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
End If

'bytes 0 to 35 = essid (36 bytes) - the essid (name) of the access point
If (ESSID <> "") Then
 For i = 0 To 35
  If Len(ESSID) >= (i + 1) Then
   obytes(i + offset) = Asc(Mid$(ESSID, (i + 1), 1))
  End If
 Next
End If

'bytes 36 to 41 = mac1 (6 bytes) - the bssid (MAC) of the access point
If (BSSID <> "") Then
 Dim mac1 As String
 mac1 = hex_digits_only(BSSID)
 For i = 36 To 41
  If Len(mac1) >= ((i - 36) + (i - 36) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(mac1, (i - 36) + (i - 36) + 1, 2))
  End If
 Next
End If

'bytes 42 to 47 = mac2 (6 bytes) - the MAC address of a client connecting to the access point
If (STA <> "") Then
 Dim mac2 As String
 mac2 = hex_digits_only(STA)
 For i = 42 To 47
  If Len(mac2) >= ((i - 42) + (i - 42) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(mac2, (i - 42) + (i - 42) + 1, 2))
  End If
 Next
End If

'bytes 48 to 79 = snonce (32 bytes) - random salt used for handshake by both parties
If (SNONCE <> "") Then
 Dim SNONCE_clean As String
 SNONCE_clean = hex_digits_only(SNONCE)
 For i = 48 To 79
  If Len(SNONCE_clean) >= ((i - 48) + (i - 48) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(SNONCE_clean, (i - 48) + (i - 48) + 1, 2))
  End If
 Next
End If

'bytes 80 to 111 = anonce (32 bytes) - random salt used for handshake by both parties
If (ANONCE <> "") Then
 Dim ANONCE_clean As String
 ANONCE_clean = hex_digits_only(ANONCE)
 For i = 80 To 111
  If Len(ANONCE_clean) >= ((i - 80) + (i - 80) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(ANONCE_clean, (i - 80) + (i - 80) + 1, 2))
  End If
 Next
End If

'bytes 112 to 367 = eapol (256 bytes) - EAPOL
If (EAPOL <> "") Then
 Dim EAPOL_clean As String
 EAPOL_clean = hex_digits_only(EAPOL)
 For i = 112 To 367
  If Len(EAPOL_clean) >= ((i - 112) + (i - 112) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(EAPOL_clean, (i - 112) + (i - 112) + 1, 2))
  End If
 Next
End If

'bytes 368 to 371 = eapol_size (4 bytes) - size of eapol
If (EAPOL_SIZE <> "") Then
 obytes(368 + offset) = CInt(EAPOL_SIZE) And &HFF&
 obytes(369 + offset) = (CInt(EAPOL_SIZE) And &HFF00&) / 256
End If

'bytes 372 to 375 = keyver (4 bytes) - the flag used to distinguish WPA from WPA2 ciphers. Value of 1 means WPA, other - WPA2
If (KEY_VERSION <> "") Then
 obytes(372 + offset) = IIf(KEY_VERSION = "1", 1, 2)
End If

'bytes 376 to 391 = keymic (16 bytes) - the final hash value. MD5 for WPA and SHA-1 for WPA2 (truncated to 128 bit)
If (KEYMIC <> "") Then
 Dim KEYMIC_clean As String
 KEYMIC_clean = hex_digits_only(KEYMIC)
 For i = 376 To 391
  If Len(KEYMIC_clean) >= ((i - 376) + (i - 376) + 2) Then
   obytes(i + offset) = hex2dec(Mid$(KEYMIC_clean, (i - 376) + (i - 376) + 1, 2))
  End If
 Next
End If

offset = offset + 392 'move to next record
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
write_hccap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Sub
