if not exist "%~1\..\Certification\signfile.bat" goto error
if not exist %3 goto error

cd "%~1\..\Certification"
signfile.bat %2 %3
cd "%~1"
goto end

:error
echo Could not sign %3
goto end

:end
