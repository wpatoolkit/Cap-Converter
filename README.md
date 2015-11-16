# Cap-Converter
<p align="center"><img src="https://raw.githubusercontent.com/wpatoolkit/Cap-Converter/master/screenshot.png" /></p>
This is a simple GUI tool for Windows that can convert back and forth between the CAP and HCCAP file formats. It can...

1. Convert CAP files to HCCAP files
2. Convert HCCAP files to CAP files
3. Preview and edit the contents of an HCCAP file

It was written using Visual Basic 6 and should work on any version of Windows.

<b>Background</b><br>
Cracking WPA/WPA2 with <a href="https://hashcat.net/wiki/doku.php?id=cracking_wpawpa2">oclHashcat</a> requires the use of an <a href="https://hashcat.net/wiki/doku.php?id=hccap">HCCAP</a> file which is a custom file format designed specifically for hashcat. Typically this file is created using <a href="http://www.aircrack-ng.org/">aircrack-ng</a> (v1.2-beta1 or later) using the command...

`aircrack-ng -J HCCAP_FILE CAP_FILE.CAP`

or by using <a href="https://hashcat.net/cap2hccap/">this online converter</a> (which is powered by <a href="http://sourceforge.net/projects/cap2hccap/">cap2hccap</a>).
