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
' This script allows the creation of the configuration file for fastscan-3.0

config_default = array ("##", "# Configuratiebestand voor FastScan", "# Vorm: key='value'", "##", "##", "# Output-dir: in die map worden mappen aangemaakt met volgend masker:", "# jjjj-mm-dd-%user%", "##", "base_output_dir='K:\Cultuur\PBC\Beeldbank\99sys_SCANS'")

set shell = CreateObject ("WScript.Shell")
set fso = CreateObject ("Scripting.FileSystemObject")
username = shell.ExpandEnvironmentStrings ("%USERNAME%")
config_path = "C:\Users\" & username & "\Applicaties\FastScan"

' Directory checking
if fso.FolderExists ("C:\Users") <> True then
	fso.CreateFolder ("C:\Users")
end if
if fso.FolderExists ("C:\Users\" & username) <> True then
	fso.CreateFolder ("C:\Users" & username)
end if
if fso.FolderExists ("C:\Users\" & username & "\Applicaties") <> True then
	fso.CreateFolder ("C:\Users" & username & "\Applicaties")
end if
if fso.FolderExists ("C:\Users\" & username & "\Applicaties\FastScan") <> True then
	fso.CreateFolder ("C:\Users" & username & "\Applicaties\FastScan")
end if
if fso.FolderExists ("C:\Users\" & username & "\Applicaties\FastScan\log") <> True then
	fso.CreateFolder ("C:\Users" & username & "\Applicaties\FastScan\log")
end if

config_file = config_path & "\config.txt"
set ObjConfig_file = fso.OpenTextFile (config_file, 2, true)
For Each line in config_default
	ObjConfig_file.WriteLine (line)
Next
ObjConfig_file.close
