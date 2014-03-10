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
' This script allows to batch-scan images using Irfan View
' Invocation: cscript fastscan-3.0.vbs type
'		where type can be: postkaart, affiche, foto, bidprent

' 2e opl http://stackoverflow.com/questions/15621395/vbscript-relative-path

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


' Function to determine the last used number
' Using logdir\prefix_lastlog.txt <= contains the last used number
Function last_number (logfile)
	set fso = CreateObject ("Scripting.FileSystemObject")
	if fso.FileExists (logfile) <> True then
		' File does not exist, create it & add 0
		set lastlog = fso.CreateTextFile (logfile)
		lastlog.WriteLine ("0")
		lastlog.Close
		set lastlog = Nothing
		last_number = 0
	else
		set lastlog = fso.OpenTextFile (logfile)
		x_last_number = lastlog.ReadLine ()
		lastlog.Close
		set lastlog = Nothing
		last_number = x_last_number
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

' Function to pad the number up to 6 items
Function pad (number)
	pad = Right (String (8, "0") & number, 6)
End Function

' Function to pad the date
Function d_pad (d)
	d_pad = Right (String (2, "0") & d, 2)
End Function

' Subroutine to update prefix_lastlog.txt
Sub update_lastlog (logfile, n)
	set lastlog = fso.CreateTextFile (logfile, true)
	lastlog.WriteLine (n)
	lastlog.Close
	set lastlog = Nothing
End Sub

' Subroutine to mimic the "pause"-key in .BAT
Sub Pause(strPause)
      MsgBox strPause, 1
End Sub
' <<<<<<<<<<<<<<<<< Application >>>>>>>>>>>>>>>>>>

' Read configuration file
set shell = CreateObject ("WScript.Shell")
set fso = CreateObject ("Scripting.FileSystemObject")
username = shell.ExpandEnvironmentStrings ("%USERNAME%")
config_file = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\config.txt"
' Base output directory "^base_output_dir='(.*)'$"
base_out_dir =  read_config_file ("^base_output_dir='(.*)'$", config_file)
outdir = base_out_dir & "\"
' Log Directory
logdir = shell.ExpandEnvironmentStrings ("%USERPROFILE%") & "\Applicaties\FastScan\log\"
' FS Directory
fs_dir = read_config_file ("^fastscan_dir='(.*)'$", config_file) & "\"
' IM Directory
im_dir = read_config_file ("^im_dir='(.*)'$", config_file) & "\"
' IV Directory
iv_dir = read_config_file ("^iview_dir='(.*)'$", config_file) & "\"

' Alexander
if LCase (username) = "tolvrij" then
	base_logdir = "L:\PBC\Beeldbank\99sys_SCANS\log_a"
	logdir = base_logdir & "\"
end if

' Create output directories
working_dir = outdir & DatePart ("yyyy", Date) & "-" & d_pad (DatePart ("m", Date)) & "-" & d_pad (DatePart ("d", Date)) & "-" & username
if fso.FolderExists (working_dir) <> true then
	fso.CreateFolder (working_dir)
end if
raw_dir = working_dir & "\RAW"
if fso.FolderExists (raw_dir) <> true then
	fso.CreateFolder (raw_dir)
end if
edit_dir = working_dir & "\EDITED"
if fso.FolderExists (edit_dir) <> true then
	fso.CreateFolder (edit_dir)
end if

' Prefix determination
'	postkaart => PKT
'	affiche => AFF
'	foto => FOT
'	bidprent => BID
if Wscript.Arguments.Count = 0 Then
	Wscript.Echo "Opgelet! Geen type gespecifieerd: script fastscan-3.0.vbs type. Programma afgesloten."
	Wscript.Sleep 5000
	Wscript.Quit
end if
Select Case Wscript.Arguments(0)
	Case "postkaart"
		prefix = "PKT"
	Case "affiche"
		prefix = "AFF"
	Case "foto"
		prefix = "FOT"
	Case "bidprent"
		prefix = "BID"
	Case Else
	Wscript.Echo "Opgelet! Fout type gespecifieerd: script fastscan-3.0.vbs type. Programma afgesloten."
	Wscript.Sleep 5000
	Wscript.Quit
End Select

' Main loop
' Program works like this:
' Simply pressing [Enter] is the default
' 1) Get the last used number from prefix_lastlog.txt
' 2) Check whether this is a backside of a previously scanned item (check backside=1)
' 3) If 2 = false, increment last_number by 1 & ask whether this item has a backside (backside=?)
' 4) Create the filename: prefix + 6-len (number)x0 + number + (A/B) + extension & ask if correct
' 5) Scan using irfanview (?)
' 6) Based on the prefix, do some additional operations
' 7) Ask whether the user wishes to do another scan or would like to terminate
backside = 0
brun = 0
item = 1
do while 1 = 1
	last = last_number (logdir & prefix & "_lastlog.txt")
	' Backside
	Wscript.Echo "Scan " & item & ":"
	item = item + 1
	if backside <> 1 then
		' Not a backside
		' Reset brun
		brun = 0
		' New number
		number = last + 1
		' Ask whether this image has a backside
		do while 1 = 1
			is_n_correct = u_input ("Automatisch gegenereerd nummer (nieuw nummer ingeven indien niet correct): [" & number & "]")
			if is_n_correct = "" then
				exit do
			end if
			if is_n_correct <> "" then
				if IsNumeric (is_n_correct) then
					number = is_n_correct
					exit do
				end if
			end if
		loop
		has_backside = u_input ("Heeft dit item een achterkant? ([J]a/[N]ee)")
		if has_backside = "" and prefix = "BID" then
			has_backside = "j"
		end if
		if InStr (LCase (has_backside), "j") <> 0 then
			' Yes
			backside = 1
		else
			' No
			backside = 0
		end if
	else
		' Is a backside
		brun = 1 ' So we know when to reset backside
		number = last
	end if
	' Ask whether number is correct
	' Create file name
	if backside = 1 and brun = 0 then
		' A side
		filename = prefix & pad (number) & "A.tif"
	elseif backside = 1 and brun = 1 then
		' B side
		filename = prefix & pad (number) & "B.tif"
	else
		' Normal case
		filename = prefix & pad (number) & ".tif"
	end if
	unique_id = prefix & pad (number)
	' Is the filename correct?
	is_f_correct = u_input ("Automatische bestandsnaam: [" & filename & "]. Correct? ([J]a/[N]ee)")
	if InStr (LCase (is_f_correct), "n") <> 0 then
		' We made a mistake - correct it
		new_number = u_input ("Geef het nummer (zonder 0'en en " & prefix & ") in:")
		new_back = u_input ("Geen achterkant (X), voorkant (A) of achterkant (B)?")
		Select Case new_back
			Case "A"
				brun = 0
				backside = 1
			Case "B"
				brun = 1
				backside = 1
			Case Else
				new_back = ""
		End Select
		filename = prefix & pad (new_number) & new_back & ".tif"
		number = new_number
		unique_id = prefix & pad (new_number)
	end if
	' Scan
	Wscript.Echo "Voorbereiden scan ..."
	set shell = CreateObject ("WScript.Shell")
	if prefix = "FOT" then
		' Ask about the border colour of the photograph
		border_type = "unbound"
		Do Until border_type = "Z" or border_type = "W" or border_type = "G"
			border_type = u_input ("Heeft de foto een zwarte (Z), witte (W) of geen (G) rand?")
			border_type = UCase (border_type)
			if border_type = "" then
				border_type = "G"
			end if
		Loop
		Pause ("Opgelet! Gebruik de juiste achtergrond voor het scannen (zwart voor witte rand en geen rand; wit voor zwarte rand)!")
	end if
	if prefix = "BID" then
		' Ask about the border colour of the item
		border_type = "unbound"
		Do Until border_type = "Z" or border_type = "W" or border_type = "G"
			border_type = u_input ("Heeft de bidprent een zwarte (Z), witte (W) of geen (G) rand?")
			border_type = UCase (border_type)
			if border_type = "" then
				border_type = "G"
			end if
		Loop
		Pause ("Opgelet! Gebruik de juiste achtergrond voor het scannen (zwart voor witte rand en geen rand; wit voor zwarte rand)!")
	end if
	' Some jiggery-pokery because some systems don't quite behave as they should
	Wscript.Echo "Scannen van " & prefix & pad (number) & " naar " & filename & "..."
	Pause ("Leg het item binnen het scanbare gedeelte op de glasplaat en druk op OK om door te gaan")
	iview = chr(34) & iv_dir & "i_view32.exe" & chr(34)
	if shell.ExpandEnvironmentStrings ("%computername%") = "PC1040198" then
		' Difficult case
'		shell.Run iview & " /scanhidden /dpi=(300,300) /convert=" & raw_dir & "\" & filename, 1, true
	'elseif shell.ExpandEnvironmentStrings ("%computername%") = "PC1240047" then
	'	' Expensive scanner with a lot of options
	'	shell.Run iview & " /scanhidden /dpi=(300,300) /convert=" & raw_dir & "\" & filename, 2, true
	else
		' Normal case
'		shell.Run iview & " /scanhidden /dpi=(300,300) /convert=" & raw_dir & "\" & filename, 2, true
	end if
	if fso.FileExists (raw_dir & "\" & filename) <> true then
		Wscript.Echo "Fout: scan niet voltooid. Mogelijk is de schijf vol of zijn er verbindingsproblemen met de scanner. Programma afgesloten."
		Wscript.Sleep 5000
		Wscript.Quit
	end if
	' Update prefix_lastlog.txt to show that this number has been scanned
	update_lastlog logdir & prefix & "_lastlog.txt", number
	' Some additional operations
	Select Case prefix
		Case "PKT"
			' Cuttings
			Wscript.Echo "Bijsnijden van " & filename & "..."
			Wscript.Echo "Bezig met bijsnijden ... "
			' Using the new black cover made everything below useless, but it's kept (one never knows)
			' Use the new cropper - with high fuzz factor due to nice contrast with black background (le expensive scanneur!) => 10% for scanners with white backgrounds
			shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "15%", 0, true
			if fso.FileExists (edit_dir & "\" & filename) <> true then
				Wscript.Echo "Fout: bijsnijden niet voltooid. Mogelijk is de schijf vol. Programma afgesloten."
				Wscript.Sleep 5000
				Wscript.Quit
			else
				Wscript.Echo "Bijsnijden voltooid"
			end if
		Case "AFF"
		Case "FOT"
			' Cuttings
			Wscript.Echo "Bijsnijden van " & filename & "..."
			Wscript.Echo "Bezig met bijsnijden ... "
			Select Case border_type
				Case "W"
					'25%
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "25%", 0, true
				Case "Z"
					'38%
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "38%", 0, true
				Case "G"
					'15% (to be on the safe side)
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "15%", 0, true
			End Select
		Case "BID"
			' Cuttings
			Wscript.Echo "Bijsnijden van " & filename & "..."
			Wscript.Echo "Bezig met bijsnijden ... "
			Select Case border_type
				Case "W"
					'25%
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "25%", 0, true
				Case "Z"
					'38%
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "38%", 0, true
				Case "G"
						'15% (to be on the safe side)
					shell.Run "cscript " & chr(34) & fs_dir & "fastscan-crop.vbs" & chr(34) & " " & chr(34) & raw_dir & "\" & filename & chr(34) & " " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & "15%", 0, true
			End Select
	End Select
	Wscript.Echo "Toevoegen metadata ..."
	shell.Run "cscript " & chr(34) & fs_dir & "fastscan-metadata.vbs" & chr(34) & " divorce " & chr(34) & edit_dir & "\" & filename & chr(34) & " " & unique_id & " " & chr(34) & username & chr(34), 0, true
	' If this image has a backside & brun = 0
	' then don't ask questions, but continue the loop
	' Else, ask questions
	if not (backside = 1 and brun = 0) then
		another = u_input ("Wilt u nog een item scannen? ([J]a/[N]ee)")
		if InStr (LCase (another), "n") <> 0 then
			' You want to stop? Please? Nooo!
			Wscript.Echo "Opgelet! Vergeet niet om na de laatste scan alle items te verplaatsen van "
			Wscript.Echo "[" & edit_dir & "] naar "
			Wscript.Echo "[K:\Cultuur\mediatheek\PB_Tolhuis\SCANS\]!"
			Wscript.Echo "Controleer ook of alle items goed gescand werden."
			' Remove RAW-folder
			fso.DeleteFolder (raw_dir)
			Wscript.Sleep 5000
			Wscript.Quit
		end if
	end if
	' Reset counters
	if brun = 1 then
		backside = 0
	end if
loop