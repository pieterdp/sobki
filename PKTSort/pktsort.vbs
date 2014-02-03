'    (c) 2013 Pieter De Praetere
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


'	Invocation: cscript pktsort.vbs [dir ...]
'	If no directory provided, uses internally set default (TODO: change)
'	This script runs through a set of folders returned by fastscan and
'	tries to extract the postcard from the scan, asks whether this is H or L
'	and creates a JPG.

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

' Function to get user input
Function u_input (prompt)
	set shell = CreateObject ("WScript.Shell")
	username = shell.ExpandEnvironmentStrings ("%USERNAME%")
	if LCase (username) = "pdpr" then
		Wscript.StdOut.Write prompt & " "
		u_input = Wscript.StdIn.ReadLine
	else
		u_input = InputBox (prompt)
	end if
End Function

'http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx
Sub folder_walk (objFolder)
	Set Folders = objFolder.SubFolders
	For Each SubFolder in Folders
		If LCase (Right (Subfolder.Path, 6)) <> "edited" and LCase (Right (Subfolder.Path, 4)) <> "jpgs" Then
			Wscript.Echo "===================================================="
			Wscript.echo SubFolder.Path
			set Files = SubFolder.Files
			For Each File in Files
				sort (File)
			Next
			folder_walk (SubFolder)
		End If
	Next
End Sub

Sub rate_file (File, Rating)
	Dim fso
	Set fso = Wscript.CreateObject ("Scripting.FileSystemObject")
	' Create a H, L & D folder in the parent folder of the file
	For Each target in Array ("H", "L", "D")
		If fso.FolderExists (fso.GetParentFolderName (File) & "\" & target) <> True Then
			fso.CreateFolder (fso.GetParentFolderName (File) & "\" & target)
		End If
	Next
	If fso.FolderExists (
End Sub

' Add suffixes (DL, L or H)
' Requires a File Object
Sub sort (File)
	set fso = Wscript.CreateObject ("Scripting.FileSystemObject")
	If LCase (fso.GetExtensionName (File)) = "tiff" or LCase (fso.GetExtensionName (File)) = "tif" Then
		set shell = Wscript.CreateObject ("Wscript.Shell")
		basefilename = fso.GetBaseName (File)
		file_ext = "." & fso.GetExtensionName (File)
		file_parent = fso.GetParentFolderName (File)
		Wscript.Echo basefilename & file_ext
		' Experimental resize
		Wscript.Echo "Uitknippen afbeelding ..."
		' Check whether the current folder ends with RAW. If so, create folder EDITED in parent_folder. If not, in current folder
		If Lcase (Right (file_parent, 3)) = "raw" Then
			edit_folder = fso.GetParentFolderName (file_parent) & "\EDITED"
			If fso.FolderExists (edit_folder) <> True Then
				fso.CreateFolder (edit_folder)
			End If
		ElseIf Lcase (Right (file_parent, 6)) = "edited" Then
			edit_folder = file_parent
		Else
			edit_folder = file_parent & "\EDITED"
			If fso.FolderExists (edit_folder) <> True Then
				fso.CreateFolder (edit_folder)
			End If
		End If
		' Do nothing as long as not every scanner uses the same fuzz factor
		' But copy to EDITED if none exists
		If fso.FileExists (edit_folder & "\" & basefilename & file_ext) <> True Then
			fso.CopyFile file_parent & "\" & basefilename & file_ext, edit_folder & "\"
		End If
	'	rcommand = "cscript L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\FastScan\fastscan-crop.vbs " & chr(34) & file_parent & "\" & basefilename & file_ext & chr(34) & " " & chr(34) & edit_folder & "\" & basefilename & file_ext & chr(34) & " 10%"
	'	shell.Run rcommand, 0, true
		' Open in IrfanView
		iview = chr(34) & "C:\Program Files\IrfanView\" & "i_view32.exe" & chr(34)
		iviewc = iview & " " & chr(34) & edit_folder & "\" & basefilename & file_ext & chr(34)
		shell.Run iviewc, 1, true
		psuffix = u_input ("Is deze afbeelding geschikt voor [H]oge kwaliteit, [L]age kwaliteit of een [D]ubbel?")
		If psuffix = "" Then
			psuffix = "H"
		End If
		new_file_name = basefilename & file_ext
		Select Case psuffix
			Case "H"
				
			Case "L"
				
			Case "D"
				
			Case "R"
				Wscript.Echo "Removing file " & basefilename & file_ext & "..."
				fso.DeleteFile edit_folder & "\" & basefilename & file_ext
				Wscript.Echo "[OK]"
			Case Else
			Wscript.Echo "Opgelet! Foute bestemming ingegeven. Programma afgesloten."
			Wscript.Sleep 5000
			Wscript.Quit
		End Select
		If psuffix <> "R" Then
			Wscript.Echo basefilename & file_ext & " -> " & new_file_name
			If fso.FileExists (edit_folder & "\" & new_file_name) = True Then
				fso.DeleteFile (edit_folder & "\" & new_file_name)
			End If
			fso.MoveFile edit_folder & "\" & basefilename & file_ext, edit_folder & "\" & new_file_name
		End If
		' Create JPG's
		Wscript.Echo "Aanmaken JPG: " & fso.GetBaseName (new_file_name) & ".jpg"
		jpg_folder = fso.GetParentFolderName (edit_folder) & "\JPGS"
		If fso.FolderExists (jpg_folder) <> True Then
			fso.CreateFolder (jpg_folder)
		End If
		icommand = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\im_convert.exe " & chr(34) & edit_folder & "\" & new_file_name & chr(34) & " " & chr(34) & jpg_folder & "\" & fso.GetBaseName (new_file_name) & ".jpg" & chr(34)
		shell.Run icommand, 0, true
		Wscript.Echo "----------------------------------------------------"
		set shell = nothing
	End If
End Sub

' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>

basedir = "L:\PBC\Beeldbank\Postkaarten\98_RAW_scans"

If Wscript.Arguments.Count = 0 Then
	Wscript.Echo "Usage: csscript pktsort.vbs dir1[ dir2 ...]"
	Wscript.Echo "Using default " & basedir
	set arguments = Array(basedir)
Else
	set arguments = Wscript.Arguments
End If
Dim fso
set fso = Wscript.CreateObject ("Scripting.FileSystemObject")


For Each Dir in arguments
	Dim objFolder
	Dim colFiles
	set objFolder = fso.GetFolder (Dir)
	set colFiles = objFolder.Files
	Wscript.echo objFolder.Path
	For Each objFile in colFiles
		sort (objFile)
	Next
	folder_walk (objFolder)
Next
