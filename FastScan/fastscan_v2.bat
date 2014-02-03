::ECHO OFF
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
:: variables between ! & ! are expanded on run time
:: This script allows to batch-scan images using Irfan View
:: Invocation is ./fastscan_v2.bat user (optional)
:: 		user		the person executing the scan (for directory creation)
:: All options are interactive and not CLI
::		start_gen	automatically generated starting number for the scan based on the number of the last scan in the log file
::		
echo Welkom bij FastScan v.2.0 - met dit programma kan je snel en semi-automatisch
echo scannen
set logdir=C:\Users\%USERNAME%\Applicaties\FastScan\log
if not exist %logdir% (
mkdir %logdir%
)
:: REPLACED WITH fastscan-3.0.vbs
start "FastScan" /wait /D"L:\PBC\Beeldbank\1_Digitalisering\0_Scansysteem\2_Scansoftware\FastScan" cscript fastscan-3.0.vbs "postkaart"
exit


:DIRCRE
cls
:: Creation of directories
if defined %1 set u=%1 else (
	set u=%USERNAME%
)
set dirname=L:\PBC\Beeldbank\Postkaarten\98_RAW_scans\%date:~9,4%-%date:~6,2%-%date:~3,2%-%u%\RAW
if not exist "%dirname%" mkdir "%dirname%"
set findir=L:\PBC\Beeldbank\Postkaarten\98_RAW_scans\%date:~9,4%-%date:~6,2%-%date:~3,2%-%u%\EDITED
if not exist "%findir%" mkdir "%findir%"

:OFFSE
:: Creation of the offset
::set logdir=L:\PBC\Beeldbank\98_RAW_scans\log
set logdir=C:\Users\%USERNAME%\Applicaties\FastScan\log
set logfile=%logdir%\lastlog.txt
if not exist %logdir% (
mkdir %logdir%
echo 0 > %logfile%
) else (
	for /f "eol=" %%i in (%logfile%) do (
		set last=%%i
	)
)
:: Ask the user whether the first postcard they will scan will use our auto-generated number or
:: whether they will enter one themselves
:: Doing this is required, as the log file only has the number of the last scanned postcard
:: The next one must have a number that's 1 higher
set /a last_t=%last%+1
set offset=%last_t%
echo Is het automatisch gegenereerde nummer van de eerste te scannen postkaart
echo correct? (zonder PKT en de eerste nullen) (leeg laten indien ja, 
set /p offset_d="anders zelf het nummer invoeren) [%last_t%] "
if not "%offset_d%"=="" set offset=%offset_d%
::if /i NOT %offset_d%=="" set offset=%offset_d%

set /a temp=%offset%-1
:: Set this to the log file or somethings won't work
echo %temp% > %logfile%
:: Default value for has_back
set hasback=0

:: Default value for run
set run=1
set brun=0

:FRBA
if %run% NEQ 1 set /a run=%run%+1
:: Ask whether this postcard has a back side
:: If we asked it the previous execution, don't ask again
if %hasback% EQU 1 set brun=2
:: Above: this is the back side (brun=2 is B, brun=1 is A)
if %hasback% EQU 0 (
	set brun=0
	set hasback=0
	set reph=N
	set /p reph="Heeft deze postkaart een in te scannen achterzijde? [J/N] (gewoon enter is gelijk aan Nee)"
	if /i "!reph!" EQU "J" (
		set brun=1
		set hasback=1 
		)
)
if %brun% EQU 2 set hasback=0
if %brun% EQU 0 set hasback=0

:PROPFN
:: Create the file name, taking in account whether this is a image with a back side
:: Numerical value
:: This is the first run, so use offset
if "!run!" EQU "1" set fnam_n=%offset% ) else (
:: It isn't, use %last% + 1 or %last% depending on brun
	for /f "eol=" %%i in (%logfile%) do (
		set last=%%i
	)
	set /a fnam_n=!last!+1
	:: brun=2 means this is the backside
	if !brun!==2 set fnam_n=!last!
)

:: Padding (see http://stackoverflow.com/questions/13398545/string-processing-in-windows-batch-files-how-to-pad-value-with-leading-zeros)
set fnam=00000%fnam_n%
set fnam=%fnam:~-6%

:: Adding A or B
if %brun% EQU 2 set fnam=%fnam%B
if %brun% EQU 1 set fnam=%fnam%A
:: Auto - the rest of the name
set fnam=PKT%fnam%
set ffnam=%fnam%.tif
:: Ask whether this is correct, else allow the user to enter one
echo De scan zal opgeslagen worden onder de naam [%ffnam%] in map [%dirname%].
set rep=J
set /p rep="Is dit correct? [J/N] (gewoon enter is gelijk aan Ja)"
if "%rep%" EQU "N" (
set /p fnam="Nieuwe bestandsnaam? (zonder extensie) "
if !brun! EQU 1 set fnam=!fnam!A
if !brun! EQU 2 set fnam=!fnam!B
set ffnam=!fnam!.tif
set /p fnam_n="Nummer van de postkaart (zonder PKT en leidende nullen)? "
)

:SCAN
:: Now do the scanning using IrfanView's command line mode
:: Auto-scanning using scan hidden doesn't quite work in the downstairs reading room
echo Leg de postkaart op de glasplaat van de scanner (binnen het scanbare deel!) en druk op Enter
pause
echo Scannen van postkaart %fnam% naar %ffnam% ...
if /i %computername% NEQ PC0840196 (
	start /wait /D"C:\Program Files\IrfanView" i_view32.exe "/scanhidden /dpi=(300,300) /convert=%dirname%\%ffnam%"
) else (
	start /wait /D"C:\Program Files\IrfanView" i_view32.exe "/scan /dpi=(300,300) /convert=%dirname%\%ffnam%"
)
:: Logging - VERY important (or nothing works)
echo %fnam_n% > %logfile%

:: Set run to run+1
set /a run=%run%+1
if exist %dirname%\%ffnam% (
echo Scan voltooid.
) else (
	echo Scan niet voltooid. Bestand [%dirname%\%ffnam%] kon niet worden aangemaakt. Misschien is de schijf vol?
	pause
)

:CROP
:: Auto-crop using a local installation of ImageMagick
::start /wait /B /DL:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8 im_convert.exe "-extract 1710x1100+0+0 %dirname%\%ffnam% %findir%\%ffnam%"
set stdsize=1
set /p stdsize="Geef de standaardgrootte van de postkaart in (1, 2, 3 of 4): (gewoon enter is 1)"
if /i "%stdsize%" EQU "1" (
	echo Bezig met bijsnijden ...
	L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe -crop 1640x1036+0+0 %dirname%\%ffnam% %findir%\%ffnam%
)
if /i "%stdsize%" EQU "2" (
	echo Bezig met bijsnijden ...
	L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe -crop 1745x1213+0+0 %dirname%\%ffnam% %findir%\%ffnam%
)
if /i "%stdsize%" EQU "3" (
	echo Bezig met bijsnijden ...
	L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe -crop 1745x1268+0+0 %dirname%\%ffnam% %findir%\%ffnam%
)
if /i "%stdsize%" EQU "4" (
	echo Bezig met bijsnijden ...
	L:\PBC\Beeldbank\99_Opvolging_scans\99_Applicaties\ImageMagick-6.8.6-8\im_convert.exe -crop 1773x1243+0+0 %dirname%\%ffnam% %findir%\%ffnam%
)

if exist %findir%\%ffnam% (
echo Bijsnijden voltooid.
) else (
echo Bijsnijden niet voltooid. Bestand [%findir%\%ffnam%] kon niet worden aangemaakt. Misschien is de schijf vol?
pause
)

:ANOTHER
if %hasback% EQU 1 GOTO FRBA
set another=J
set /p another="Wilt u nog een afbeelding scannen? [J/N] (gewoon enter is gelijk aan Ja)"
if /i "%another%" EQU "N" (
	echo OPGELET! Vergeet niet om na de laatste scan alle scans te verplaatsen van
	echo [%findir%] naar [K:\Cultuur\»mediatheek\PB_Tolhuis\SCANS\%date:~9,4%-%date:~6,2%-%date:~3,2%]!
	echo Controleer ook of alle postkaarten goed gescand zijn.
	echo Druk op J + Enter om de RAW-map te verwijderen:
	rmdir %dirname% /s
	PAUSE
) else (
	GOTO FRBA
)