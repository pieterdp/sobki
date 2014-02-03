@ECHO OFF
Setlocal EnableDelayedExpansion

::    (c) 2013 Pieter De Praetere
::
::    This program is free software: you can redistribute it and/or modify
::    it under the terms of version 3 of the GNU General Public License
::    as published by the Free Software Foundation.
::
::    This program is distributed in the hope that it will be useful,
::    but WITHOUT ANY WARRANTY; without even the implied warranty of
::    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
::    GNU General Public License for more details.
::
::    You should have received a copy of the GNU General Public License
::    along with this program.  If not, see <http://www.gnu.org/licenses/>.
::
::

:: REPLACED WITH pktsort.vbs
start "PKTSort" /wait /D"L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\PKTSort" cscript pktsort-2.0.vbs
exit


:DIRCRE
cls
:: Creation of directories
if defined %1 set u=%1 else (
	set u=%USERNAME%
)
set basedir=L:\PBC\Beeldbank\Postkaarten\98_RAW_scans
if defined %2 set srcdir=%2 else (
	set /p ndir="Geef de bronmap in (relatief t.o.v. [%basedir%] en zonder EDITED):" 
	set srcdir=%basedir%\!ndir!\EDITED
	::echo %srcdir%
)

:FILCRE
:: Create base CSV-file for this operation
set csvpath=I:\PKTSort\Log
if not exist "%csvpath%" mkdir "%csvpath%"
set csvname=!ndir!.csv
if not exist %csvpath%\!csvname! (
	:: CSV file header
	set header=^"Nummer^";^"Dubbel van^";^"Orig. In Memorix^";^"Orig. Hernummerd?^"
	echo !header! > %csvpath%\!csvname!
)

:LOOP
:: Loop through srcdir - may be repeated multiple times during one run of the program
for %%X in (!srcdir!\*.tif) do (
	set fnam=%%~nX%%~xX
	:: View the image
	start /wait /D"C:\Program Files\IrfanView" i_view32.exe "!srcdir!\!fnam!" /one
	:: Ask whether this is a double
	set is_double=N
	echo Afbeelding %%~nX ...
	set /p is_double="Is deze postkaart een dubbel? [J/N] (Gewoon ENTER is N)"
	if /i "!is_double!" EQU "J" (
		:: Hernoem
		:: New filename
		set nfnam=%%~nXDL%%~xX
		echo Hernoemen van [!fnam!] naar [!nfnam!] ...
		ren "!srcdir!\!fnam!" "!nfnam!"
		echo Hernoemd.
	) else (
		set add=H
		set /p add="Opladen in hoge (H) of lage (L) kwaliteit? (Gewoon enter is H)"
		if /i "!add!" EQU "L" (
			set nfnam=%%~nXL%%~xX
		) else (
			set nfnam=%%~nXH%%~xX
		)
		echo Hernoemen van [!fnam!] naar [!nfnam!] ...
		ren "!srcdir!\!fnam!" "!nfnam!"
		echo Hernoemd.
	)
)

:AGAIN
set another=J
set /p another="Wilt u nog een map controleren? [J/N] (gewoon enter is gelijk aan Ja)"
if /i "%another%" EQU "N" (
	echo OPGELET! Vergeet niet om alle scans te verplaatsen van
	echo [%srcdir%] naar [K:\Cultuur\»mediatheek\PB_Tolhuis\SCANS\%date:~9,4%-%date:~6,2%-%date:~3,2%]!
	echo Controleer ook of alle nummers kloppen.
	PAUSE
) else (
	GOTO DIRCRE
)