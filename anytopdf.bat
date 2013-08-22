@ECHO OFF

REM AnyToPDF Windows Batch Version
REM converted from https://code.google.com/p/anytopdf/
REM depdep.coder@gmail.com

SET VERSION=1.0

REM check parameter
IF [%1]==[] GOTO USAGE
IF [%2]==[] GOTO USAGE
IF "%1"=="help" GOTO USAGE
IF "%1"=="-h" GOTO USAGE

REM check if input file exit
IF NOT EXIST "%1" (
	ECHO input file '%1' not exist
	GOTO END
)

REM check if Open Office 3 installed
SET exe=%ProgramFiles(x86)%\OpenOffice.org 3\program\soffice.exe
IF NOT EXIST "%exe%" (
	ECHO Open Office 3 NotFound
	GOTO END
)

REM check if AnyToPDF.xba exist
SET xba=%APPDATA%\OpenOffice.org\3\user\basic\Standard\AnyToPDF.xba

IF NOT EXIST "%xba%" (
	ECHO Macro file not found, will attempt to create
	REM create macro
	ECHO ^<?xml version="1.0" encoding="UTF-8"?^> > "%xba%"
	ECHO ^<^!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd"^> >> "%xba%"
	ECHO ^<script:module xmlns:script="http://openoffice.org/2000/script" script:name="AnyToPDF" script:language="StarBasic"^>REM  *****  BASIC  ***** >> "%xba%"
	ECHO. >> "%xba%"
	ECHO Sub Main >> "%xba%"
	ECHO. >> "%xba%"
	ECHO End Sub >> "%xba%"
	ECHO Sub ConvertAnyToPDF^(inFile,outFile^) >> "%xba%"
	ECHO    inURL = ConvertToURL^(inFile^) >> "%xba%"
	ECHO    oDoc = StarDesktop.loadComponentFromURL^(inURL, ^&quot;_blank^&quot;, 0, Array^(AnyToPDFMPV^(^&quot;Hidden^&quot;, True^), ^)^) >> "%xba%"
	ECHO    outURL = ConvertToURL^(outFile^) >> "%xba%"
	ECHO    oDoc.storeToURL^(outURL, Array^(AnyToPDFMPV^(^&quot;FilterName^&quot;, ^&quot;writer_pdf_Export^&quot;^), ^)^) >> "%xba%"
	ECHO    oDoc.close^(True^) >> "%xba%"
	ECHO End Sub >> "%xba%"
	ECHO. >> "%xba%"
	ECHO Function AnyToPDFMPV^( Optional cName As String, Optional uValue ^) As com.sun.star.beans.PropertyValue >> "%xba%"
	ECHO    Dim oPropertyValue As New com.sun.star.beans.PropertyValue >> "%xba%"
	ECHO    If Not IsMissing^( cName ^) Then >> "%xba%"
	ECHO       oPropertyValue.Name = cName >> "%xba%"
	ECHO    EndIf >> "%xba%"
	ECHO    If Not IsMissing^( uValue ^) Then >> "%xba%"
	ECHO       oPropertyValue.Value = uValue >> "%xba%"
	ECHO    EndIf >> "%xba%"
	ECHO    AnyToPDFMPV^(^) = oPropertyValue >> "%xba%"
	ECHO End Function >> "%xba%"
	ECHO. >> "%xba%"
	ECHO ^</script:module^> >> "%xba%"
	
	REM update script registration
	ECHO Extending user's openoffice.org scripts registry with AnyToPDF macro module.
	SET script=%APPDATA%\OpenOffice.org\3\user\basic\Standard\script.xlb
	
	ECHO ^<?xml version="1.0" encoding="UTF-8"?^> > "%script%"
	ECHO ^<^!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd"^> >> "%script%"
	ECHO ^<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false"^> >> "%script%"

	FOR %%i IN (%APPDATA%\OpenOffice.org\3\user\basic\Standard\*.*) DO (
		IF ".xba" == "%%~xi" ECHO   ^<library:element library:name="%%~ni"/^> >> "%script%"
	)

	ECHO ^</library:library^> >> "%script%"
)

REM run open office with macro
"%exe%" -writer -headless -invisible "macro:///Standard.AnyToPDF.ConvertAnyToPDF(%~f1,%~f2)"

GOTO END

:USAGE
ECHO AnyToPDF v%VERSION% (LGPL)
ECHO converts arbitrary documents to PDF format using openoffice.org v3 macros.
ECHO http://code.google.com/p/anytopdf
ECHO.
ECHO Usage: anytopdf.bat [infile] [outfile]
ECHO.
ECHO   [infile]        input file in any format that openoffice.org can read
ECHO                   (doc/xls/odt/rtf/html/txt/etc.)
ECHO   [outfile]       output file (will be written in PDF format)
ECHO.

:END
