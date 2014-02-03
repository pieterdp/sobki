'    (c) 2014 Pieter De Praetere
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

'	Script to add as much metadata as possible to TIFF-files (using FADGI-guidelines)
dim rTags
'rTags = Array ("IFD0:ImageWidth", "IFD0:ImageHeight", "IFD0:BitsPerSample", "IFD0:Compression", "IFD0:PhotometricInterpretation", "IFD0:ImageDescription", "IFD0:Make", "IFD0:Model", "IFD0:SamplesPerPixel", "IFD0:XResolution", "IFD0:YResolution", "IFD0:ResolutionUnit", "IFD0:Software", "IFD0:ModifyDate", "IFD0:Artist", "ExifIFDColorSpace", "ExifIFDImageUniqueID", "ICC_Profile") ' Required tags
rTags = Array ("ImageWidth", "ImageHeight", "BitsPerSample", "Compression", "PhotometricInterpretation", "ImageDescription", "Make", "Model", "SamplesPerPixel", "XResolution", "YResolution", "ResolutionUnit", "Software", "ModifyDate", "Artist", "ColorSpace", "ImageUniqueID", "ICC_Profile") ' Required tags

' Function to get the output from a command
' http://stackoverflow.com/questions/5690134/running-command-line-silently-with-vbscript-and-getting-output
Function run_and_get (command)
	'Wscript.Echo command
	set shell = WScript.CreateObject("WScript.Shell")
	set fso = CreateObject ("Scripting.FileSystemObject")
	username = shell.ExpandEnvironmentStrings ("%USERNAME%")
	output = "C:\Users\" & username & "\Applicaties\FastScan\md_cmd_output.txt"
	c_command = "cmd /c " & chr(34) & command & chr(34) & " > " & output
	shell.Run c_command, 0, true
	set shell = nothing
	set file = fso.OpenTextFile (output, 1)
	text = file.ReadLine
	file.Close
	' Remove trailing newline http://blogs.technet.com/b/heyscriptingguy/archive/2005/05/20/how-can-i-remove-the-last-carriage-return-linefeed-in-a-text-file.aspx
	iTLength = Len (text)
	iTEnd = Right (iTLength, 2)
	If iTEnd = vbCrLf Then
		text = Left (text, iTLength - 2)
	End If
	run_and_get = text
End Function

' Function to check for every required tag
' whether any information is stored in it
' if so, add it to a dictionary object
' @param string filename (full path)
Function CheckTags (FileName)
	dim cTags
	dim bCommand
	Set cTags = CreateObject ("Scripting.Dictionary")
	' Setting the -s3 flag removes the name of the tag
	' Using -f means empty values are equal to '-'
	' Using -n means values are given in numbers when appropriate
	bCommand = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\exiftool.exe -n -s3 -f "
	For Each Tag in rTags
		tValue = run_and_get (bCommand & "-" & Tag & " " & FileName)
		If tValue <> "-" Then
			cTags.Add Tag, tValue
			Wscript.Echo Tag & ": " & tValue
		End If
	Next
	Set CheckTags = cTags
End Function

' Function to us imagemagick to get a lot of info
' all collected in 1 string because IM is slow
' Collect everything you can, values separated by ;
' and then split into a dictionary
Function IMIdentify (iFileName)
	Set iValues = CreateObject ("Scripting.Dictionary")
	Dim iCommand, iFormat
	iFormat = "-format " & chr(34) & "%[w];%[h];%[colorspace]" & chr(34)
	iCommand = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\identify.exe " & " " & iFormat & " " & chr(34) & iFileName & chr(34)
	Wscript.Echo iCommand
	Set IMIdentify = iValues
End Function

' Application below
FilePath = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\PKT004560.tif"
dim oTags, nTags, tKeys
Set oTags = CheckTags (FilePath)
Set nTags = CreateObject ("Scripting.Dictionary")
Set IMInfo = IMIdentify (FilePath)
For Each rTag in rTags
	If oTags.Item (rTag) <> "" Then
		nTags.Add rTag, oTags.Item (rTag)
	Else
		Select Case rTag
			Case "ImageWidth"
			Case "ImageHeight"
		End Select
	End If
Next