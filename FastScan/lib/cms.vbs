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
' input_filename scanner_name

' <<<<<<<<<<<<<<<<< Functions >>>>>>>>>>>>>>>>>>>>
Function read_config_value (line, pattern)
	set r = new Regexp
	r.IgnoreCase = True
	r.Pattern = pattern
	r.Global = False
	if r.Test (line) = True then
		set match = r.Execute (line)
		set submatch = match.Item(0).SubMatches
		if submatch.Count = 1 then
			read_config_value = submatch.Item(0)
		end if
	end if
End Function

Function read_config_file (pattern, file)
	set ObjConfig_file = fso.OpenTextFile (file)
	dim line
	do while not ObjConfig_file.AtEndOfStream
		line = ObjConfig_file.ReadLine ()
		if read_config_value (line, pattern) <> Empty then
			read_config_file = read_config_value (line, pattern)
		end if
	loop
	ObjConfig_file.Close
	set ObjConfig_file = Nothing
	' catch-all
End Function

' Function to read the scanners.xml-file
Function read_scanners_xml ()
	' scanners.xml is always in ../etc/
	'dim scanner_dir = fso.GetParentFolderName (shell.CurrentDirectory) & "\etc\"
	scanner_dir = fs_dir & "\etc\"
	scanner_file = scanner_dir & "scanners.xml"
	set oXML = CreateObject ("MSXML.DOMDocument")
	oXML.Load scanner_file
	set read_scanners_xml = oXML
End Function

' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>

' Read configuration file
set shell = CreateObject ("WScript.Shell")
set fso = CreateObject ("Scripting.FileSystemObject")
username = shell.ExpandEnvironmentStrings ("%USERNAME%")
config_file = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\config.txt"
' CMS Directory
cms_dir = read_config_file ("^cms_dir='(.*)'$", config_file) & "\"
' FS Directory
fs_dir = read_config_file ("^fastscan_dir='(.*)'$", config_file) & "\"
set sXML = read_scanners_xml

if Wscript.Arguments.Count <> 2 Then
	Wscript.Echo "Opgelet! cms.vbs input_filepath scanner_name"
	Wscript.Sleep 5000
	Wscript.Quit
end if
input_filepath = Wscript.Arguments(0)
scanner_name = Wscript.Arguments(1)

' Get scanner profile
set xIProf = sXML.selectSingleNode ("/scanners/scanner[@name='" & scanner_name & "']/icc[@relation='source-profile']")
lIProf = xIProf.Text
if xIProf.getAttribute ("type") = "included" then
	' Some profiles are included with fastscan.
	lIProf = fs_dir & "lib\icc-profiles\" & lIProf
end if
if XIProf.getAttribute ("type") = "none" then
	lIProf = ""
end if

' Get output profile
set xOProf = sXML.selectSingleNode ("/scanners/scanner[@name='" & scanner_name & "']/icc[@relation='destination-profile']")
lOProf = xOProf.Text
if xOProf.getAttribute ("type") = "included" then
	' Some profiles are included with fastscan.
	lOProf = fs_dir & "lib\icc-profiles\" & lOProf
end if
if XOProf.getAttribute ("type") = "none" then
	' Always use the default output profile
	lOProf = fs_dir & "lib\icc-profiles\eciRGB_v2_ICCv4.icc"
end if
' Move the file
fso.MoveFile input_filepath, input_filepath & ".ick.tif"
cms_string = chr(34) & cms_dir & "tifficc.exe" & chr(34) & " "
if lIProf <> "" then
	cms_string = cms_string & "-i" & chr(34) & lIProf & chr(34)
end if
cms_string = cms_string & " -o" & chr(34) & lOProf & chr(34) & " -e " & chr(34) & input_filepath & ".ick.tif" & chr(34) & " " & chr(34) & input_filepath & chr(34)
shell.Run cms_string, 0, true
fso.DeleteFile input_filepath & ".ick.tif"