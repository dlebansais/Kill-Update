if not exist "%~1..\Version Tools\VersionBuilder.exe" goto error

"%~1..\Version Tools\VersionBuilder.exe" %2 %3 -v
goto end

:error
echo Failed to update version.
goto end

:end
