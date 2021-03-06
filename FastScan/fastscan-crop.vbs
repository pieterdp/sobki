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
' This script uses image_magick to auto-crop images that have been scanned,
' removing the huge whitespace around them (which should be black for white
' paper and white for everything else)
' Usage: cscript fastscan-crop.vbs input_file, output_file and fuzz_factor (full path!)


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
	set fso = CreateObject ("Scripting.FileSystemObject")
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


' Function to get the output from a command
' http://stackoverflow.com/questions/5690134/running-command-line-silently-with-vbscript-and-getting-output
Function run_and_get (command)
	'Wscript.Echo command
	set shell = WScript.CreateObject("WScript.Shell")
	set fso = CreateObject ("Scripting.FileSystemObject")
	output = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\cmd_output.txt"
	If fso.FileExists (output) <> True Then
		set op = fso.CreateTextFile (output)
		op.WriteLine ("0")
		op.Close
	End If
	c_command = "cmd /c " & command & " > " & chr(34) & output & chr(34)
	shell.Run c_command, 0, true
	set shell = nothing
	set file = fso.OpenTextFile (output, 1)
	text = file.ReadAll
	file.Close
	run_and_get = text
End Function

' Function to get the dimensions (in image_magick "crop" format) of
' the scan without the whitespace
Function get_dimensions (input, fuzz)
	im_command = im_dir & "im_convert.exe" & " " & chr(34) & input & chr(34) & " -virtual-pixel edge -blur 0x15 -fuzz " & fuzz & " -trim -format %[fx:w]x%[fx:h]+%[fx:page.x]+%[fx:page.y] info:"
	get_dimensions = run_and_get (im_command)
End Function
' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>
' im_convert.exe compare_with.tif -virtual-pixel edge -blur 0x15 -fuzz 10% -trim info:
' compare_with.tif TIFF 1598x1002 2551x4200+0+0 8-bit sRGB 0.109u 0:00.093
' im_convert.exe compare_with.tif -crop 1598x1002+0+0 +repage compare_with_result.png
' http://stackoverflow.com/questions/5690134/running-command-line-silently-with-vbscript-and-getting-output

' Read configuration file
set shell = CreateObject ("WScript.Shell")
config_file = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\config.txt"
' FS Directory
fs_dir = read_config_file ("^fastscan_dir='(.*)'$", config_file) & "\"
' IM Directory
im_dir = read_config_file ("^im_dir='(.*)'$", config_file) & "\"
' IV Directory
iv_dir = read_config_file ("^iview_dir='(.*)'$", config_file) & "\"


' Check arguments
If Wscript.Arguments.Count <> 3 Then
	Wscript.Echo "Error: this script requires 3 arguments: input_file, output_file and fuzz_factor. Exiting program."
	Wscript.Sleep 5000
	Wscript.Quit
End If
input_file = Wscript.Arguments.Item (0)
output_file = Wscript.Arguments.Item (1)
fuzz_factor = Wscript.Arguments.Item (2)

set fso = CreateObject ("Scripting.FileSystemObject")
If fso.FileExists (input_file) <> True Then
	Wscript.Echo "Error: input file " & input_file & " does not seem to exist. Exiting program."
	Wscript.Sleep 5000
	Wscript.Quit
End If

' Guessing dimensions of new image
Wscript.Echo "Guessing dimensions of sub-image ..."
im_dimensions = get_dimensions (input_file, fuzz_factor)

' Creating new file
set shell = WScript.CreateObject("WScript.Shell")
Wscript.Echo "Cropping input file with dimensions " & im_dimensions & "..."
im_commandx =  chr(34) & im_dir & "im_convert.exe" & chr(34) & " " & chr(34) & input_file & chr(34) & " -crop " & im_dimensions & " +repage " & chr(34) & output_file & chr(34)
shell.Run im_commandx, 0, true
If fso.FileExists (output_file) <> True Then
	Wscript.Echo "Error: output file " & output_file & " was not created. Perhaps the drive is full (or something else borked). Exiting program."
	Wscript.Sleep 5000
	Wscript.Quit
End If
Wscript.Echo "Cropping complete."
Wscript.Quit