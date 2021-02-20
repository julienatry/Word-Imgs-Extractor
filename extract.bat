@ECHO OFF

::--------------------------------------------------------------------------------------
SET source=""
SET outputdir=""
::--------------------------------------------------------------------------------------


::vars
SET tempfolder="%temp%\WordImgsExtractor"
SET wordconv="C:\Program Files\Microsoft Office\Office14\Wordconv.exe"

:MakeOutputDir
MKDIR %outputdir%

:CopyToTemp
XCOPY /E /I /Y %source% %tempfolder%

:ConvertDocToDocx
ECHO Beginning doc to docx conversion
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.doc"') DO (
	ECHO Processing %%~nF
	%wordconv% -oice -nme "%%~fF" "%%~fFx"
)


SETLOCAL ENABLEDELAYEDEXPANSION

:ExtractPNG
FOR /F "TOKENS=*" %%F IN ('DIR /S /B "%tempfolder%\*.docx"') DO ( 
	"C:\Program Files\7-Zip\7z.exe" x "%%~fF" -aou -o"%%~pF\%%~nF" *.png -r > nul
	SET c=0

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
	SET c=0

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
	SET c=0

	FOR /F "TOKENS=*" %%E IN ('DIR /S /B "%%~dF%%~pF%%~nF\word\media\*.jpeg"') DO (
		SET /a c+=1
		SET imgname=%%~nF_!c!%%~xE
		ECHO !imgname!
		REN "%%~fE" "!imgname!"
		COPY "%%~dE%%~pE\!imgname!" "%outputdir%"
	)

)

:DeleteTemp
RMDIR /S /Q "%tempfolder%"

ENDLOCAL
PAUSE
