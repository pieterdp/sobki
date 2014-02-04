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
'	meta-dater.vbs dir

If Wscript.Arguments.Count = 0 Or Wscript.Arguments.Count = 1 Then
	'Arguments = Array ("L:\PBC\Beeldbank\Postkaarten\98_RAW_scans") ' Default
	Wscript.Echo "Usage: csscript metadater.vbs type dir1[ dir2 ...]"
	Arguments = Array (InputBox ("Gelieve de map met scans in te geven (tiffs):"))
Else
	set Arguments = Wscript.Arguments
End If

Dim fso
Set fso = Wscript.CreateObject ("Scripting.FileSystemObject")
Set shell = WScript.CreateObject("WScript.Shell")

For Each Dir in Arguments
	If fso.FolderExists (Dir) = True Then
		Dim objFolder
		Dim colFiles
		set objFolder = fso.GetFolder (Dir)
		set colFiles = objFolder.Files
		Wscript.echo objFolder.Path
		For Each objFile in colFiles
			AddMetadata (objFile)
		Next
		folder_walk (objFolder)
	End If
Next

'http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx
Sub folder_walk (objFolder)
	Set Folders = objFolder.SubFolders
	For Each SubFolder in Folders
		Wscript.echo SubFolder.Path
		set Files = SubFolder.Files
		For Each File in Files
			AddMetadata (File)
		Next
		folder_walk (SubFolder)
	Next
End Sub

Sub AddMetadata (aObjFile)
	Dim mType
	mType = Wscript.Arguments (0)
	Select Case mType
		Case "init"
			' Initial creation of metadata => username is in the name of the directory
			' Split on '-', last item is the username
			Dim sFolder, mUserName, mFileName, mNumber
			sFolder = Split (fso.GetParentFolderName (fso.GetParentFolderName (aObjFile.Path)), "-")
			mUserName = sFolder (3)
			mFileName = aObjFile.Path
			mNumber = fso.GetBaseName (mFileName)
			Wscript.Echo "Scan " & mNumber
			' Use fastscan-metadata.vbs to actually add the metadata
			shell.Run "cscript fastscan-metadata.vbs " & chr(34) & mFileName & chr(34) & " " & chr(34) & mNumber & chr(34) & " " & chr(34) & mUserName & chr(34), 0, true
		Case "final"
			' JPG's have been created, add metadata to JPGS in JPGS-subfolder
	End Select
End Sub