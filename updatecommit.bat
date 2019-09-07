if not exist "%~1..\Version Tools\GitCommitId.exe" goto error
"%~1..\Version Tools\GitCommitId.exe" %2 -u
goto end

:error
echo Failed to update Commit Id.
goto end

:end
