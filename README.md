# Cap-Converter
<p align="center"><img src="https://raw.githubusercontent.com/wpatoolkit/Cap-Converter/master/screenshot.png" /></p>
This is a small GUI tool for Windows that can open, edit and save WPA hash information from CAP and HCCAP files. It can...

1. Convert CAP files to HCCAP files<br>
2. Convert HCCAP files back to CAP files<br>
3. Preview and edit the contents of an HCCAP file<br>

It does not require any external programs or dependencies to be installed (such as <a href="https://www.wireshark.org/">Wireshark</a>, <a href="http://www.winpcap.org/">WinPcap</a> or <a href="http://www.tcpdump.org/">libpcap</a>). It directly reads and writes the raw bytes of the files.

Before converting it is recommended that you clean your caps first <a href="http://hackforums.net/showthread.php?tid=2974396">manually using Wireshark</a> or with <a href="https://code.google.com/p/pyrit/">pyrit</a> using the command...

`pyrit -r INPUT.CAP -o OUTPUT.CAP strip`

to only contain one handshake for one network but even if you don't this program will still make it's best attempt at extracting the correct WPA handshake information.

<b>How to Convert CAP to HCCAP</b><br>
1. Press the "Open CAP..." button to open a .CAP file<br>
2. Verify the hash information looks correct in the "HCCAP Info" box<br>
3. Press the "Save As HCCAP..." button to save the information to an .HCCAP file<br>

<b>How to Convert HCCAP to CAP</b><br>
1. Press the "Open HCCAP..." button to open a .HCCAP file<br>
2. Verify the hash information looks correct in the "HCCAP Info" box<br>
3. Press the "Save As CAP..." button to save the information to a .CAP file<br>

<b>Background</b><br>
Cracking WPA/WPA2 with <a href="https://hashcat.net/wiki/doku.php?id=cracking_wpawpa2">oclHashcat</a> requires the use of an <a href="https://hashcat.net/wiki/doku.php?id=hccap">HCCAP</a> file which is a custom file format designed specifically for hashcat. Typically this file is created using <a href="http://www.aircrack-ng.org/">aircrack-ng</a> (v1.2-beta1 or later) using the command...

`aircrack-ng -J HCCAP_FILE CAP_FILE.CAP`

or by using <a href="https://hashcat.net/cap2hccap/">this online converter</a> (which is powered by <a href="http://sourceforge.net/projects/cap2hccap/">cap2hccap</a>). However if you aren't comfortable with either of those options this program allows you to perform your CAP-to-HCCAP conversions offline with a familiar Windows interface.

This program effectively merges together the functionality and adds a GUI to these existing projects:<br>
1. <a href="https://github.com/philsmd/analyze_hccap">analyze_hccap</a><br>
2. <a href="https://github.com/philsmd/craft_hccap">craft_hccap</a><br>
3. <a href="https://github.com/philsmd/hccap2cap">hccap2cap</a><br>

<b>VB6 Runtimes</b><br>
This program was written in Visual Basic 6 which means it should work on any modern version of Windows but just in case you need to download the VB6 runtimes you can do so from here:

<a href="https://www.microsoft.com/en-us/download/details.aspx?id=24417">https://www.microsoft.com/en-us/download/details.aspx?id=24417</a>
