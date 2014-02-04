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

'VB-script to batch-convert a set of images from <input> to jpg
'with a quality setting of 90 using a local installation of
'imagemagick
'invoke as cscript convert_images.vbs dir1 dir2 ...

If Wscript.Arguments.Count = 0 Then
	'Arguments = Array ("L:\PBC\Beeldbank\Postkaarten\98_RAW_scans") ' Default
	Wscript.Echo "Usage: csscript convert_images.vbs dir1[ dir2 ...]"
	Arguments = Array (InputBox ("Gelieve de map met scans in te geven:"))
Else
	set Arguments = Wscript.Arguments
End If

Dim fso
set fso = Wscript.CreateObject ("Scripting.FileSystemObject")


For Each Dir in Arguments
	Dim objFolder
	Dim colFiles
	set objFolder = fso.GetFolder (Dir)
	set colFiles = objFolder.Files
	Wscript.echo objFolder.Path
	For Each objFile in colFiles
		convert_to_jpg (objFile)
	Next
	folder_walk (objFolder)
Next

'http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx
Sub folder_walk (objFolder)
	Set Folders = objFolder.SubFolders
	For Each SubFolder in Folders
		Wscript.echo SubFolder.Path
		set Files = SubFolder.Files
		For Each File in Files
			convert_to_jpg (File)
		Next
		folder_walk (SubFolder)
	Next
End Sub

'Convert images to JPG
'Requires a File Object
Sub convert_to_jpg (File)
	set fso = Wscript.CreateObject ("Scripting.FileSystemObject")
	If LCase (fso.GetExtensionName (File)) = "tiff" or LCase (fso.GetExtensionName (File)) = "tif" Then
		Wscript.echo File.Name & " -> " & fso.GetBaseName (File) & ".jpg"
		Dim shell
		set shell = Wscript.CreateObject ("Wscript.Shell")
		' TMP
'		If LCase (Right (fso.GetBaseName (File), 1)) = "a" Then
			' 1x Left
'			shell.Run "L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe " & """" & fso.GetParentFolderName (File) & "\" & File.Name & """" & " -rotate " & """" & "-90" & """" & " " & """" & fso.GetParentFolderName (File) & "\" & File.Name & """", 0, true
'		ElseIf LCase (Right (fso.GetBaseName (File), 1)) = "b" Then
'			' 1x Right
'			shell.Run "L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe " & """" & fso.GetParentFolderName (File) & "\" & File.Name & """" & " -rotate " & """" & "-90" & """" & " " & """" & fso.GetParentFolderName (File) & "\" & File.Name & """", 0, true
'		End If
		If fso.FolderExists (fso.GetParentFolderName (File) & "\" & "JPGS\") <> True Then
			fso.CreateFolder (fso.GetParentFolderName (File) & "\" & "JPGS\")
		End If
		shell.Run "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\im_convert.exe " & """" & fso.GetParentFolderName (File) & "\" & File.Name & """" & " -quality 95 " & """" & fso.GetParentFolderName (File) & "\" & "JPGS\" & fso.GetBaseName (File) & ".jpg" & """", 0, true
		set shell = nothing
	End If
End Sub