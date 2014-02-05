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

'Exif.Image.ImageHistory ?
dim rTags, evTags
'rTags = Array ("IFD0:ImageWidth", "IFD0:ImageHeight", "IFD0:BitsPerSample", "IFD0:Compression", "IFD0:PhotometricInterpretation", "IFD0:ImageDescription", "IFD0:Make", "IFD0:Model", "IFD0:SamplesPerPixel", "IFD0:XResolution", "IFD0:YResolution", "IFD0:ResolutionUnit", "IFD0:Software", "IFD0:ModifyDate", "IFD0:Artist", "ExifIFD:ColorSpace", "ExifIFD:ImageUniqueID", "ICC_Profile") ' Required tags
rTags = Array ("ImageWidth", "ImageHeight", "BitsPerSample", "Compression", "PhotometricInterpretation", "ImageDescription", "Make", "Model", "SamplesPerPixel", "XResolution", "YResolution", "ResolutionUnit", "Software", "ModifyDate", "Artist", "ColorSpace", "ImageUniqueID", "ICC_Profile", "CreateDate") ' Required tags
evTags = Array ("ImageWidth", "ImageLength", "BitsPerSample", "Compression", "PhotometricInterpretation", "ImageDescription", "Make", "Model", "SamplesPerPixel", "XResolution", "YResolution", "ResolutionUnit", "Software", "DateTime", "Artist", "ColorSpace", "ImageUniqueID") ' Required tags
Use_Fast = true
If Use_Fast = true Then
	aTags = evTags
Else
	aTags = rTags
End If

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
	set f = fso.GetFile (output)
	If f.Size = 0 Then
		text = ""
	Else
		set file = fso.OpenTextFile (output, 1)
		text = file.ReadLine
		file.Close
		' Remove trailing newline http://blogs.technet.com/b/heyscriptingguy/archive/2005/05/20/how-can-i-remove-the-last-carriage-return-linefeed-in-a-text-file.aspx
		iTLength = Len (text)
		iTEnd = Right (iTLength, 2)
		If iTEnd = vbCrLf Then
			text = Left (text, iTLength - 2)
		End If
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
	bCommand = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\exiftool.exe -s3 -f "
	For Each Tag in rTags
		If Tag = "ColorSpace" Then
			tValue = run_and_get (bCommand & "-" & Tag & " " & chr(34) & FileName & chr(34))
		Else
			tValue = run_and_get (bCommand & "-" & Tag & " " & chr(34) & FileName & chr(34))
		End If
		If tValue <> "-" Then
			cTags.Add Tag, tValue
			'Wscript.Echo Tag & ": " & tValue
		End If
	Next
	Set CheckTags = cTags
End Function

' Function to read all available tags using exiv2.exe
' which should be way faster than exiftool.exe
' @param string filename (full path)
Function ExivCheckTags (eFileName)
	dim eTags, eCommand
	Set eTags = CreateObject ("Scripting.Dictionary")
	eCommand = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\exiv2\exiv2.exe pr -P v -g "
	For Each eTag in evTags
		eValue = run_and_get (eCommand & chr(34) & eTag & chr(34) & " " & chr(34) & FileName & chr(34))
		If eValue <> "" Then
			eTags.Add eTag, eValue
		'	Wscript.Echo eTag & ": " & eValue
		End If
	Next
	Set ExivCheckTags = eTags
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

' Function to pad the date
Function d_pad (d)
	d_pad = Right (String (2, "0") & d, 2)
End Function

' Convert a date to ISO8601
' @param string date
' @return string iso-date
Function ISODate (iDate)
'2014-02-04T14:27:16+00:00
	iYear = Year (iDate)
	iMonth = d_pad (Month (iDate))
	iDay = d_pad (Day (iDate))
	iHour = d_pad (Hour (iDate))
	iMin = d_pad (Minute (iDate))
	iSec = d_pad (Second (iDate))
	iTZ = "+01:00"
	iISODate = iYear & "-" & iMonth & "-" & iDay & "T" & iHour & ":" & iMin & ":" & iSec & iTZ
	ISODate = iISODate
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
'FilePath = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\PKT004560.tif"
dim oTags, nTags, tKeys, IMInfo, Number, UserName, FileName
Number = Wscript.Arguments (1)
UserName = Wscript.Arguments (2)
FileName = Wscript.Arguments (0)
FilePath = FileName
Set fso = CreateObject ("Scripting.FileSystemObject")
Set f = fso.GetFile (FilePath)
Set nTags = CreateObject ("Scripting.Dictionary") ' Based on the name of the computer to which the scanner is connected
Set nMakes = CreateObject ("Scripting.Dictionary")
nMakes ("PC1240047") = "Mikrotek"
nMakes ("PC1040198") = "Canon"
nMakes ("PC0840196") = "HP"
Set nModels = CreateObject ("Scripting.Dictionary")
nModels ("PC1240047") = "ScanMaker 9800 XL+"
nModels ("PC1040198") = "CanoScan 3200F"
nModels ("PC0840196") = "Scanjet 8200"
Set Shell = CreateObject ("WScript.Shell")
ComputerName = Shell.ExpandEnvironmentStrings ("%computername%")

' Real app starts about here
Wscript.Echo "Parsing metadata ... "
Set Shell = nothing
IMInfo = IMIdentify (FilePath)
If Use_Fast = true Then
	Set oTags = ExivCheckTags (FilePath)
Else
	Set oTags = CheckTags (FilePath)
End If
'Set oTags = CheckTags (FilePath)
For Each rTag in aTags
	If rTag = "Software" Then
		nTags.Add rTag, "Sobki"
	End If
	If oTags.Item (rTag) <> "" And rTag <> "Software" Then
		nTags.Add rTag, oTags.Item (rTag)
	Else
		Select Case rTag
			Case "ImageWidth"
				nTags.Add rTag, IMInfo (0)
			Case "ImageHeight"
				nTags.Add rTag, IMInfo (1)
			Case "ImageLength"
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
				' Don't know this, but convert sets this to 3 (RGB), and convert is used in FastScan to crop stuff
			Case "ResolutionUnit"
				nTags.Add rTag, IMInfo (7)
			Case "ImageUniqueID"
				nTags.Add rTag, Number
			Case "ModifyDate"
				nTags.Add rTag, ISODate (f.DateLastModified)
			Case "CreateDate"
				nTags.Add rTag, ISODate (f.DateCreated)
			Case "DateTime"
				nTags.Add rTag, ISODate (f.DateLastModified)
			Case "ImageDescription"
				' Leave this empty
			Case "Artist"
				' 2 Artists: Original & Scan
				nTags.Add rTag, "Original: ; Scan: PBT: " & UserName
			Case "Make"
				nTags.Add rTag, nMakes (ComputerName)
			Case "Model"
				nTags.Add rTag, nModels (ComputerName)
		End Select
	End If
Next
' Add new metadata to the file
Wscript.Echo "Adding metadata to file ... "
Dim eCommand
eCommand = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\EXIFTool\exiftool.exe -n "
If Use_Fast = true Then
	Dim wCommand
	wCommand = "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\exiv2\exiv2.exe mo -M"
	Set Shell = CreateObject ("WScript.Shell")
	For Each nTag in nTags
		Select Case nTag
			Case "ModifyDate"
			' Folded into Xmp.tiff.DateTime
				wvCommand = wCommand & chr(34) & "set Xmp.tiff.DateTime " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case "CreateDate"
			'Exif.Image.DateTimeOriginal
			'Exif.Photo.DateTimeDigitized
				wvCommand = wCommand & chr(34) & "set Exif.Photo.DateTimeDigitized " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case "ColorSpace"
			'Exif.Photo.ColorSpace
				wvCommand = wCommand & chr(34) & "set Exif.Photo.ColorSpace " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case "ImageUniqueID"
			'Exif.Photo.ImageUniqueID
				wvCommand = wCommand & chr(34) & "set Exif.Photo.ImageUniqueID " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case "ICC_Profile"
			'Exif.Image.InterColorProfile
				wvCommand = wCommand & chr(34) & "set Exif.Image.InterColorProfile " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case "Software"
			'Exif.Image.Software & Xmp.tiff.Software - exists multiple times
				wvCommand = wCommand & chr(34) & "set Exif.Image.Software " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " -M" & chr(34) & "set Xmp.tiff.Software " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
			Case Else
				wvCommand = wCommand & chr(34) & "set Xmp.tiff." & nTag & " " & chr(39) & nTags (nTag) & chr(39) & chr(34) & " " & chr(34) & FilePath & chr(34)
		End Select
		'Wscript.Echo wvCommand
		Shell.Run wvCommand, 0, true
	Next
Else
	Set Shell = CreateObject ("WScript.Shell")
	For Each nTag in nTags
		'Wscript.Echo "Writing " & nTag & " to value " & nTags (nTag) '& ": " & eCommand & "-" & nTag & "=" & nTags (nTag) & " " & chr(34) & FilePath & chr(34)
		Shell.Run eCommand & "-" & nTag & "=" & chr(34) & nTags (nTag) & chr(34) & " " & chr(34) & FilePath & chr(34), 0, true
	Next
End If

Set Shell = nothing
Set Shell = CreateObject ("WScript.Shell")
' Now split this off in the same directory (to keep them together)
Wscript.Echo "Splitting off metadata file ... "
' Writing a more readable format may be done using exiftool (reads an .exv like any other image)
If Use_Fast = true Then
	Shell.Run "L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\exiv2\exiv2.exe ex -e a " & chr(34) & FilePath & chr(34), 0, true
	Else
	Shell.Run "cmd /c " & eCommand & " -j --FileSize --FileModifyDate --FileAccessDate --FilePermissions " & chr(34) & FilePath & chr(34) & " > " & chr(34) & FilePath & ".json" & chr(34), 0, true
End If
' Quit
Wscript.Echo "Finished."
