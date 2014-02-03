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
	If r.Test (line) = True then
		set match = r.Execute (line)
		set submatch = match.Item(0).SubMatches
		If submatch.Count = 1 then
			read_config_value = submatch.Item(0)
		End If
	End If
End Function


' Function to get the output from a command
' http://stackoverflow.com/questions/5690134/running-command-line-silently-with-vbscript-and-getting-output
Function run_and_get (command)
	'Wscript.Echo command
	set shell = WScript.CreateObject("WScript.Shell")
	set fso = CreateObject ("Scripting.FileSystemObject")
	username = shell.ExpandEnvironmentStrings ("%USERNAME%")
	output = "C:\Users\" & username & "\Applicaties\FastScan\cmd_output.txt"
	c_command = "cmd /c " & chr(34) & command & chr(34) & " > " & output
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
	im_command = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\im_convert.exe " & chr(34) & input & chr(34) & " -virtual-pixel edge -blur 0x15 -fuzz " & fuzz & " -trim -format %[fx:w]x%[fx:h]+%[fx:page.x]+%[fx:page.y] info:"
	get_dimensions = run_and_get (im_command)
End Function
' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>
' im_convert.exe compare_with.tif -virtual-pixel edge -blur 0x15 -fuzz 10% -trim info:
' compare_with.tif TIFF 1598x1002 2551x4200+0+0 8-bit sRGB 0.109u 0:00.093
' im_convert.exe compare_with.tif -crop 1598x1002+0+0 +repage compare_with_result.png
' http://stackoverflow.com/questions/5690134/running-command-line-silently-with-vbscript-and-getting-output

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
im_commandx = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\im_convert.exe " & chr(34) & input_file & chr(34) & " -crop " & im_dimensions & " +repage " & chr(34) & output_file & chr(34)
shell.Run im_commandx, 0, true
If fso.FileExists (output_file) <> True Then
	Wscript.Echo "Error: output file " & output_file & " was not created. Perhaps the drive is full (or something else borked). Exiting program."
	Wscript.Sleep 5000
	Wscript.Quit
End If
Wscript.Echo "Cropping complete."
Wscript.Quit