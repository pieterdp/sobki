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
'	Parameters:
'	fastscan-metadata.vbs filename number username
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
	Dim iCommand, iFormat
	iFormat = "-format " & chr(34) & "%[w];%[h];%[colorspace];%[profiles];%[x];%[y];%[C];%[units];%[depth];%[channels]" & chr(34)
	iCommand = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\identify.exe " & " " & iFormat & " " & chr(34) & iFileName & chr(34)
	Dim iReturn, iOptions
	iReturn = run_and_get (iCommand)
	iOptions = Split (iReturn, ";", -1, 1)
	IMIdentify = iOptions
End Function

' Function to convert the colorspace
' from the tiff-colorspace field to
' a numeric value
Function ConvertColorSpace (cSpace)
	Dim cSpaces, cOutput
	Set cSpaces = CreateObject ("Scripting.Dictionary")
	cSpaces ("whiteiszero") = 0
	cSpaces ("blackiszero") = 1
	cSpaces ("rgb") = 2
	cSpaces ("srgb") = 2
	cSpaces ("rgb palette") = 3
	cSpaces ("transparency mask") = 4
	cSpaces ("cmyk") = 5
	cSpaces ("ycbcr") = 6
	cSpaces ("cielab") = 8
	cSpaces ("icclab") = 9
	cSpaces ("itulab") = 10
	cSpaces ("color filter array") = 32803
	cSpaces ("pixar logl") = 32844
	cSpaces ("pixar logluv") = 32845
	cSpaces ("libear raw") = 34892
	cOutput = cSpaces (LCase (cSpace))
	If cOutput = "" Then
		cOutput = 2 ' In this case, a sensible default - TODO config file
	End If
	ConvertColorSpace = cOutput
End Function

' Application below
If Wscript.Arguments.Count <> 3 Then
	Wscript.Echo "Opgelet! Te weinig argumenten: cscript fastscan-metadata.vbs bestandsnaam nummer gebruikersnaam. Programma afgesloten."
	Wscript.Sleep 5000
	Wscript.Quit
End If
FilePath = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\PKT004560.tif"
dim oTags, nTags, tKeys, IMInfo, Number, UserName, FileName
Number = Wscript.Arguments ()
UserName = Wscript.Arguments ()
FileName = Wscript.Arguments ()
Set oTags = CheckTags (FilePath)
Set nTags = CreateObject ("Scripting.Dictionary")
IMInfo = IMIdentify (FilePath)
For Each rTag in rTags
	If oTags.Item (rTag) <> "" Then
		nTags.Add rTag, oTags.Item (rTag)
	Else
		Select Case rTag
			Case "ImageWidth"
				nTags.Add rTag, IMInfo (0)
			Case "ImageHeight"
				nTags.Add rTag, IMInfo (1)
			Case "Compression"
				nTags.Add rTag, IMInfo (6)
			Case "BitsPerSample"
				nTags.Add rTag, IMInfo (8)
			Case "XResolution"
				nTags.Add rTag, IMInfo (4)
			Case "YResolution"
				nTags.Add rTag, IMInfo (5)
			Case "ColorSpace"
				nTags.Add rTag, IMInfo (2)
			Case "ICC_Profile"
				nTags.Add rTag, IMInfo (3)
			Case "PhotometicInterpretation"
				' Convert between this field (int16u) and colorspace (string) -> 2 (RGB) is the default in this case
				nTags.Add rTag, ConvertColorSpace (IMInfo (2))
			Case "SamplesPerPixel"
				' Don't now this, but convert sets this to 3 (RGB), and convert is used in FastScan to crop stuff
			Case "ResolutionUnit"
				nTags.Add rTag, IMInfo (7)
			Case "ImageUniqueId"
				nTags.Add rTag, Number
			Case "ModifyDate"
			Case "ImageDescription"
			Case "Artist"
			Case "Make"
			Case "Model"
			Case "Software"
		End Select
	End If
Next
For Each elem In nTags
	Wscript.Echo elem & " - " & nTags(elem)
Next