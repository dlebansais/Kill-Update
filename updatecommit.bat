if exist "%~1packages\GitCommitId*" goto run1
if exist "%~1..\Version Tools\GitCommitId.exe" goto run2
goto error

:run1
for /D %%F in (%~1packages\GitCommitId*) do "%%F\lib\net48\GitCommitId.exe" %2 -u
goto end

:run2
"%~1..\Version Tools\GitCommitId.exe" %2 -u
goto end

:error
echo Failed to update Commit Id.
goto end

:end
