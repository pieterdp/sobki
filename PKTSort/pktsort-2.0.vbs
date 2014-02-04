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


'	Invocation: cscript pktsort.vbs [dir ...]
'	If no directory provided, uses internally set default (TODO: change)
'	This script runs through a set of folders returned by fastscan and
'	tries to extract the postcard from the scan, asks whether this is H or L.
'
'	Directory structure
'	<base>
'		-> EDITED : where fastscan stores its modified scans
'		-> RAW : where fastscan stores its raw scans (removed by fastscan)
'		-> SORTED : where pktsort stores the sorted scans from EDITED
'				-> H, L, D : rated images
'					-> JPGS : where pktsort stores the JPGS

' <<<<<<<<<<<<<<<<< Options >>>>>>>>>>>>>>>>>>>>>>
' FSO
set fso = Wscript.CreateObject ("Scripting.FileSystemObject")

' Shell
set Shell = Wscript.CreateObject ("Wscript.Shell")

' Ratings (i.e. (H)oog, (L)aag, (D)ubbel)
Ratings = Array ("H", "L", "D")


' <<<<<<<<<<<<<<<<< Functions >>>>>>>>>>>>>>>>>>>>
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

' Function to create the subdirectories of SORTED as defined in RATINGS
Sub create_subdirs_sorted (SortDir, fRatings)
	For Each fRating in fRatings
		If fso.FolderExists (SortDir & "\" & fRating) <> True Then
			fso.CreateFolder (SortDir & "\" & fRating)
		End If
	Next
End Sub

' Folder walk function
Sub folder_walk (objFolder)
	set Folders = objFolder.SubFolders
	For Each SubFolder in Folders
		FolderAbs = SubFolder.Path ' Absolute path of the folder containing the images
		FolderName = fso.GetBaseName (FolderAbs)
		' The excludes subdirectories are created by the workflow and must not in itself be traversed
		' This script will take its images from EDITED and store them in SORTED
		If LCase (FolderName) <> "raw" and LCase (FolderName) <> "jpgs" and LCase (FolderName) <> "sorted" Then
			Wscript.Echo "===================================================="
			Wscript.Echo FolderAbs
			set Files = SubFolder.Files
			For Each File in Files
				sort File, FolderAbs
				Wscript.Echo "----------------------------------------------------"
			Next
			folder_walk (SubFolder)
		End If
	Next
End Sub

' Function to rate the image
' 1) open the image in IrfanView & wait for closure
' 2) ask which rating the image deserves (H, L, D)
' 3) move the image to the right folder
Sub rate (fFileObj, fSortPath)
	iview = chr(34) & "C:\Program Files\IrfanView\" & "i_view32.exe" & chr(34)
	' Open File
	Shell.Run iview & " " & chr(34) & fFileObj.Path & chr(34), 1, True
	' Rate
	sRating = "unbound"
	Do Until fso.FolderExists (fSortPath & "\" & sRating) = True or sRating = "R"
		sRating = u_input ("Is deze afbeelding geschikt voor [H]oge kwaliteit, [L]age kwaliteit of een [D]ubbel?")
		If sRating = "" Then
			sRating = "H"
		End If
		' Default to uppercase
		sRating = UCase (sRating)
		' Files are moved to the folder corresponding with the letter-rating they have been assigned
	Loop
	
	'If fso.FolderExists (fSortPath & "\" & sRating) <> True and sRating <> "R" Then
	'	' If sRating = R, then remove the file in question and do not error out
	'	' Error out
	'	Wscript.Echo "Opgelet! Foute bestemming ingegeven. Programma afgesloten."
	'	Wscript.Sleep 5000
	'	Wscript.Quit
	'End If
	If sRating = "R" Then
			fso.DeleteFile fFileObj.Path
	Else
		FileName = fso.GetFileName (fFileObj.Path)
		' Overwriting is impossible, so remove already existing file and then move it
		If fso.FileExists (fSortPath & "\" & sRating & "\" & FileName) = True Then
			fso.DeleteFile fSortPath & "\" & sRating & "\" & FileName
		End If
		fso.MoveFile fFileObj.Path, fSortPath & "\" & sRating & "\" & FileName
		' Check-up
		If fso.FileExists (fSortPath & "\" & sRating & "\" & FileName) <> True Then
			Wscript.Echo "Fout: bestand " & fSortPath & "\" & sRating & "\" & FileName & " is niet aangemaakt. Misschien is de schijf vol? Programma afgesloten."
			Wscript.Sleep 5000
			Wscript.Quit
		End If
		' Create JPG'S (inside the ratings folder)
		'create_jpg fSortPath & "\" & sRating & "\" & FileName ' Deferred: other application
	End If
End Sub

' Function create JPG-version of a rated image (fImage is not an object, but a path!)
Sub create_jpg (fImage)
	fRatedPath = fso.GetParentFolderName (fImage)
	If fso.FolderExists (fRatedPath & "\JPGS") <> True Then
		fso.CreateFolder (fRatedPath & "\JPGS")
	End If
	fNImageName = fso.GetBaseName (fImage) & ".jpg" ' Name of the JPG
	Wscript.Echo "Aanmaken JPG " & fNImageName & " in JPGS\ ... "
	icommand = "K:\Cultuur\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\ImageMagick\im_convert.exe " & chr(34) & fImage & chr(34) & " " & chr(34) & fRatedPath & "\JPGS\" & fNImageName & chr(34)
	shell.Run icommand, 0, true
	' Check
	If fso.FileExists (fRatedPath & "\JPGS\" & fNImageName) <> True Then
		Wscript.Echo "Fout: bestand " & fRatedPath & "\JPGS\" & fNImageName & " kon niet worden aangamaakt. Misschien is de schijf vol? Programma afgesloten."
		Wscript.Sleep 5000
		Wscript.Quit
	End If
		
End Sub

' Function responsable for overall sorting (has subfunctions)
Sub sort (objFile, FolderPath)
	' Check whether we are in the EDITED-subfolder. If so, create (if not exists) a subfolder in the parent called SORTED. Else, create in current folder
	If LCase (fso.GetBaseName (FolderPath)) = "edited" Then
		ParentFolder = fso.GetParentFolderName (FolderPath)
		If fso.FolderExists (ParentFolder & "\SORTED") <> True Then
			fso.CreateFolder (ParentFolder & "\SORTED")
		End If
		' Create subdirectories of SORTED (1 per rating - eases uploading in MEMORIX because we can simply auto-fill the ID-number from the file name)
		create_subdirs_sorted ParentFolder & "\SORTED", Ratings
		SortPath = ParentFolder & "\SORTED"
		BasePath = ParentFolder
	Else
		' Create a subfolder called SORTED
		If fso.FolderExists (FolderPath & "\SORTED") <> True Then
			fso.CreateFolder (FolderPath & "\SORTED")
		End If
		' Create subdirectories of SORTED (1 per rating - eases uploading in MEMORIX because we can simply auto-fill the ID-number from the file name)
		create_subdirs_sorted FolderPath & "\SORTED", Ratings
		SortPath = FolderPath & "\SORTED"
		BasePath = FolderPath
	End If
	Wscript.Echo "Behandelen afbeelding " & objFile.Name & " ..."
	' Rate the images & create JPG'S
	rate objFile, SortPath
	' Remove RAW-folder if exists (should be done by FastScan)
	If fso.FolderExists (BasePath & "\RAW") = True Then
		fso.DeleteFolder (BasePath & "\RAW")
	End If
End Sub

' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>
If Wscript.Arguments.Count = 0 Then
	'Arguments = Array ("L:\PBC\Beeldbank\Postkaarten\98_RAW_scans") ' Default
	Arguments = Array (InputBox ("Gelieve de map met scans in te geven:"))
Else
	set Arguments = Wscript.Arguments
End If


' Start the loop
For Each Dir in Arguments
	' Below only happens for the directories in the argument list
	Dim objFolder, colFiles
	set objFolder = fso.GetFolder (Dir)
	set colFiles = objFolder.Files
	Wscript.echo objFolder.Path
	' Below happens for all sub-directories
	For Each objFile in colFiles
		sort objFile, objFolder.Path
		Wscript.Echo "----------------------------------------------------"
	Next
	folder_walk (objFolder)
Next

Wscript.Echo "Opgelet: vergeet niet de scans te verplaatsen naar de dagmap op"
Wscript.Echo "[K:\Cultuur\»mediatheek\PB_Tolhuis\SCANS\]"
Wscript.Sleep 5000
Wscript.Quit