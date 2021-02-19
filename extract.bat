@ECHO OFF

SET source=""
SET outputdir=""

SET tempfolder="%temp%\WordExtractor"
SET wordconv="C:\Program Files\Microsoft Office\Office14\Wordconv.exe"

SET /A c=0

MKDIR %outputdir%

:CopyToTemp
XCOPY /E /I /Y %source% %tempfolder%

:ConvertDocToDocx
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.doc"') DO (
	ECHO %%~fF
	%wordconv% -oice -nme "%%~fF" "%%~fFx"
)



SETLOCAL ENABLEDELAYEDEXPANSION

:ExtractPNG
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.docx"') DO ( 
"C:\Program Files\7-Zip\7z.exe" x "%%~fF" -aou -o"%%~pF\%%~nF" *.png -r > nul
SET /a c=0

FOR /F "TOKENS=*" %%E IN ('DIR /S /B "%%~dF%%~pF%%~nF\word\media\*.png"') DO (
SET /a c+=1
SET imgname=%%~nF_!c!%%~xE
ECHO !imgname!
REN "%%~fE" "!imgname!"
COPY "%%~dE%%~pE\!imgname!" "%outputdir%"
)

)

:ExtractJPG
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.docx"') DO ( 
"C:\Program Files\7-Zip\7z.exe" x "%%~fF" -aou -o"%%~pF\%%~nF" *.jpg -r > nul
SET /a c=0

FOR /F "TOKENS=*" %%E IN ('DIR /S /B "%%~dF%%~pF%%~nF\word\media\*.jpg"') DO (
SET /a c+=1
SET imgname=%%~nF_!c!%%~xE
ECHO !imgname!
REN "%%~fE" "!imgname!"
COPY "%%~dE%%~pE\!imgname!" "%outputdir%"
)

)

:ExtractJPEG
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.docx"') DO ( 
"C:\Program Files\7-Zip\7z.exe" x "%%~fF" -aou -o"%%~pF\%%~nF" *.jpeg -r > nul
SET /a c=0

FOR /F "TOKENS=*" %%E IN ('DIR /S /B "%%~dF%%~pF%%~nF\word\media\*.jpeg"') DO (
SET /a c+=1
SET imgname=%%~nF_!c!%%~xE
ECHO !imgname!
REN "%%~fE" "!imgname!"
COPY "%%~dE%%~pE\!imgname!" "%outputdir%"
)

)

DEL /F /S /Q "%tempfolder%"

ENDLOCAL
PAUSE
