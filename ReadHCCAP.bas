Attribute VB_Name = "Read_HCCAP"
Option Explicit

Public Function ReadHCCAP(filepath As String) As hccap_record()
 On Error GoTo read_hccap_error
 
 Dim tmp_hccap_records() As hccap_record
 
 Dim ibytes() As Byte
 Dim total_bytes As Long
 Dim i As Long, j As Long, offset As Long
 num_hccap_records = 0

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
 ReDim ibytes(0 To total_bytes - 1) As Byte
 Get #iFile, 1, ibytes
 Close #iFile
 
 If (total_bytes <= 391) Or (total_bytes Mod 392 <> 0) Then
  MsgBox "Invalid HCCAP file!", vbCritical + vbOKOnly, "Invalid File"
  Exit Function
 End If
 
 num_hccap_records = total_bytes / 392
 ReDim tmp_hccap_records(0 To num_hccap_records - 1) As hccap_record
 offset = 0
 
 For j = 0 To num_hccap_records - 1
 
  'bytes 0 to 35 = essid (36 bytes) - the essid (name) of the access point
  tmp_hccap_records(j).ESSID = ""
  For i = 0 To 35
   If (ibytes(i + offset) <> 0) Then
    tmp_hccap_records(j).ESSID = tmp_hccap_records(j).ESSID & Chr(ibytes(i + offset))
   End If
  Next
  
  'bytes 36 to 41 = mac1 (6 bytes) - the bssid (MAC) of the access point
  tmp_hccap_records(j).BSSID = ""
  For i = 36 To 41
   tmp_hccap_records(j).BSSID = tmp_hccap_records(j).BSSID & dec2hex(ibytes(i + offset))
   If (i < 41) Then
    tmp_hccap_records(j).BSSID = tmp_hccap_records(j).BSSID & ":"
   End If
  Next
  
  'bytes 42 to 47 = mac2 (6 bytes) - the MAC address of a client connecting to the access point
  tmp_hccap_records(j).STATION_MAC = ""
  For i = 42 To 47
   tmp_hccap_records(j).STATION_MAC = tmp_hccap_records(j).STATION_MAC & dec2hex(ibytes(i + offset))
   If (i < 47) Then
    tmp_hccap_records(j).STATION_MAC = tmp_hccap_records(j).STATION_MAC & ":"
   End If
  Next
  
  'bytes 48 to 79 = snonce (32 bytes) - random salt used for handshake by both parties
  tmp_hccap_records(j).SNONCE = ""
  For i = 48 To 79
   tmp_hccap_records(j).SNONCE = tmp_hccap_records(j).SNONCE & dec2hex(ibytes(i + offset))
   If (i <> 63) And (i <> 79) Then
    tmp_hccap_records(j).SNONCE = tmp_hccap_records(j).SNONCE & " "
   End If
   If (i = 63) Then
    tmp_hccap_records(j).SNONCE = tmp_hccap_records(j).SNONCE & vbNewLine
   End If
  Next
  
  'bytes 80 to 111 = anonce (32 bytes) - random salt used for handshake by both parties
  tmp_hccap_records(j).ANONCE = ""
  For i = 80 To 111
   tmp_hccap_records(j).ANONCE = tmp_hccap_records(j).ANONCE & dec2hex(ibytes(i + offset))
   If (i <> 95) And (i <> 111) Then
    tmp_hccap_records(j).ANONCE = tmp_hccap_records(j).ANONCE & " "
   End If
   If (i = 95) Then
    tmp_hccap_records(j).ANONCE = tmp_hccap_records(j).ANONCE & vbNewLine
   End If
  Next
  
  'bytes 112 to 367 = eapol (256 bytes) - EAPOL
  Dim EAPOL_SIZE As Long
  EAPOL_SIZE = bytes2num(ibytes(368 + offset), ibytes(369 + offset))
  tmp_hccap_records(j).EAPOL = ""
  For i = 112 To 367
   If ((i - 111) <= EAPOL_SIZE) Then
    tmp_hccap_records(j).EAPOL = tmp_hccap_records(j).EAPOL & dec2hex(ibytes(i + offset))
    If (i <> 127) And (i <> 143) And (i <> 159) And (i <> 175) And (i <> 191) And (i <> 207) And (i <> 223) And (i <> 239) And (i <> 255) And (i <> 271) And (i <> 287) And (i <> 303) And (i <> 319) And (i <> 335) And (i <> 351) And (i <> 367) And ((i - 111) <> EAPOL_SIZE) Then
     tmp_hccap_records(j).EAPOL = tmp_hccap_records(j).EAPOL & " "
    End If
    If (i = 127) Or (i = 143) Or (i = 159) Or (i = 175) Or (i = 191) Or (i = 207) Or (i = 223) Or (i = 239) Or (i = 255) Or (i = 271) Or (i = 287) Or (i = 303) Or (i = 319) Or (i = 335) Or (i = 351) Then
     tmp_hccap_records(j).EAPOL = tmp_hccap_records(j).EAPOL & vbNewLine
    End If
   End If
  Next
  
  'bytes 368 to 371 = eapol_size (4 bytes) - size of eapol
  tmp_hccap_records(j).EAPOL_SIZE = EAPOL_SIZE
 
  'bytes 372 to 375 = keyver (4 bytes) - the flag used to distinguish WPA from WPA2 ciphers. Value of 1 means WPA, other - WPA2
  tmp_hccap_records(j).KEY_VERSION = ibytes(372 + offset)
 
  'bytes 376 to 391 = keymic (16 bytes) - the final hash value. MD5 for WPA and SHA-1 for WPA2 (truncated to 128 bit)
  tmp_hccap_records(j).KEY_MIC = ""
  For i = 376 To 391
   tmp_hccap_records(j).KEY_MIC = tmp_hccap_records(j).KEY_MIC & dec2hex(ibytes(i + offset))
   If (i <> 391) Then
    tmp_hccap_records(j).KEY_MIC = tmp_hccap_records(j).KEY_MIC & " "
   End If
  Next
  
  offset = offset + 392 'move to next record
  
 Next
 
ReadHCCAP = tmp_hccap_records
Exit Function
read_hccap_error:
MsgBox "The following error occured: " & vbNewLine & "Error # " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
End Function
