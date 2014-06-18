PUSHD %CD%
CD bin
e:\wsusoffline\bin\wget.exe -S -N "http://tanzil.net/res/text/metadata/quran-data.xml"
RENAME quran-uthmani.xml download.php
e:\wsusoffline\bin\wget.exe -S -N --content-disposition --post-data "quranType=uthmani&marks=true&sajdah=true&rub=true&alef=true&outType=xml&agree=true" "http://tanzil.net/pub/download/download.php"
RENAME download.php quran-uthmani.xml
e:\wsusoffline\bin\wget.exe "http://tanzil.net/trans/"
FOR /F delims^=^"^ tokens^=6^,12 %%I IN ('FIND "http://tanzil.net/trans/" index.html') DO IF "%%J"=="return download(this)" (FOR /F "delims=^/ tokens=4" %%Q IN ("ECHO %%I") DO (e:\wsusoffline\bin\wget.exe -S -N --content-disposition "http://tanzil.net/trans/?transID=%%Q&type=txt")) ELSE (FOR /F "delims=^/ tokens=4" %%Q IN ("ECHO %%J") DO (e:\wsusoffline\bin\wget.exe -S -N --content-disposition "http://tanzil.net/trans/?transID=%%Q&type=txt"))
FOR /F delims^=^"^ tokens^=5^,6^,12 %%H IN ('FIND "http://tanzil.net/trans/" index.html') DO @(@IF "%%J"=="return download(this)" (@FOR /F "delims=^/ tokens=4" %%Q IN ("ECHO %%I") DO @ECHO|SET/P=^^^<translation file="%%Q" ) ELSE (@FOR /F "delims=^/ tokens=4" %%Q IN ("ECHO %%J") DO @ECHO|SET/P=^^^<translation file="%%Q" )) >> translations.xml & (@FOR /F "delims=^<^> tokens=5,8" %%P IN ("ECHO %%H") DO @ECHO name="%%P" translator="%%Q"^>^</translation^>) >> translations.xml
DEL index.html
POPD
GOTO :done
setlocal EnableDelayedExpansion
FOR %%i IN (1, 1, 114) DO (
SET ii=000%%i
FOR %%x IN (1, 1, 286) DO (
SET xx=000%%x
e:\wsusoffline\bin\wget.exe "http://tanzil.net/res/audio/afasy/!ii:~3!!xx:~3!.mp3"
)
)
endlocal
:done