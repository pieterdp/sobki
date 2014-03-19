'    (c) 2013, 2014 Pieter De Praetere
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of version 3 of the GNU General Public License
'    as published by the Free Software Foundation.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'
' This script allows the creation of the configuration file for fastscan-3.0

Sub touchFile (fileName)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists (fileName) <> True Then
		Set file = fso.CreateTextFile(fileName, True)
		file.WriteLine("TO_BE_OVERWRITTEN 00:00:00:00:00:00:xx")
		file.Close
	End If
End Sub

set shell = CreateObject ("WScript.Shell")
set fso = CreateObject ("Scripting.FileSystemObject")
' Check for windows versions > XP
if fso.FolderExists ("C:\Program Files\IrfanView") = True then
	config_default = array ("##", "# Configuratiebestand voor FastScan", "# Vorm: key='value'", "##", "##", "# Output-dir: in die map worden mappen aangemaakt met volgend masker:", "# jjjj-mm-dd-%user%", "##", "base_output_dir='K:\Cultuur\PBC\Beeldbank\99sys_SCANS'", "fastscan_dir='K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\FastScan'", "iview_dir='C:\Program Files\IrfanView'", "im_dir='K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick'", "exf_dir='L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool'", "exv_dir='L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\exiv2'")
else
	config_default = array ("##", "# Configuratiebestand voor FastScan", "# Vorm: key='value'", "##", "##", "# Output-dir: in die map worden mappen aangemaakt met volgend masker:", "# jjjj-mm-dd-%user%", "##", "base_output_dir='K:\Cultuur\PBC\Beeldbank\99sys_SCANS'", "fastscan_dir='K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\FastScan'", "iview_dir='C:\Program Files (x86)\IrfanView'", "im_dir='K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick'", "exf_dir='L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool'", "exv_dir='L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\exiv2'")
end if


config_path = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan"
' Directory checking
if fso.FolderExists (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties") <> True then
	fso.CreateFolder (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties")
end if
if fso.FolderExists (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan") <> True then
	fso.CreateFolder (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan")
end if
if fso.FolderExists (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\log") <> True then
	fso.CreateFolder (shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\log")
end if

config_file = config_path & "\config.txt"
set ObjConfig_file = fso.OpenTextFile (config_file, 2, true)
For Each line in config_default
	ObjConfig_file.WriteLine (line)
Next
ObjConfig_file.close

' Create some required files, or else we crash
touchFile shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\md_cmd_output.txt"