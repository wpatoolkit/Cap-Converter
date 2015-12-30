Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type tagInitCommonControlsEx
 lngSize As Long
 lngICC As Long
End Type

Public Type hccap_record
 ESSID As String
 BSSID As String
 STATION_MAC As String
 SNONCE As String
 ANONCE As String
 EAPOL As String
 EAPOL_SIZE As String
 KEY_VERSION As String
 KEY_MIC As String
End Type

Public last_path As String
Public num_hccap_records As Long
Public tmp_hccap_records() As hccap_record
Public current_index As Long
Public current_file As String

Public Sub Main()
 On Error Resume Next
 Dim iccex As tagInitCommonControlsEx
 iccex.lngSize = LenB(iccex)
 iccex.lngICC = &H200
 InitCommonControlsEx iccex
 Load Form1
 Form1.Show
 'Load Form2
 'Form2.Show
End Sub

Public Function hex_digits_only(input_str As String) As String
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

Public Function dec2hex(d As Byte) As String
 dec2hex = Right$("0" & Hex$(d), 2)
End Function

Public Function hex2dec(h As String) As Byte
 hex2dec = IIf(Len(h) > 0, CByte("&H" & h), 0)
End Function

Public Function hex2dec_lng(h As String) As Long
 hex2dec_lng = IIf(Len(h) > 0, CLng("&H" & h), 0)
End Function

Public Function bytes2num(LoByte As Byte, HiByte As Byte) As Long
 If HiByte And &H80 Then
  bytes2num = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
 Else
  bytes2num = (HiByte * &H100) Or LoByte
 End If
 'bytes2num = CLng("&H" & Right$("0" & Hex$(LoByte), 2) & Right$("0" & Hex$(HiByte), 2))
End Function
