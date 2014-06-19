setlocal ENABLEDELAYEDEXPANSION
FOR %%A IN (*.en.*.resx) DO SET tmp=%%A && RENAME %%A !tmp:en.=!
SET tmp=
endlocal